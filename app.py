import os
import re
import pdfplumber
import pandas as pd
from flask import Flask, render_template, request, send_file
import logging

# Suppress pdfminer logs
logging.getLogger('pdfminer').setLevel(logging.WARNING)

# ========== CONFIG ==========
app = Flask(__name__)
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0  # Disable caching
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "converted"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ========== HELPERS ==========
def clean_numeric_value(value):
    if isinstance(value, str):
        return float(re.sub(r"[^\d.]", "", value.replace(",", "")))
    return value

def extract_text(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            return "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())
    except Exception as e:
        logging.error(f"Error extracting text: {e}")
        return ""

def extract_table_data(pdf_path):
    rows = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table and len(table) > 1:
                    for i, row in enumerate(table):
                        # Skip header rows
                        if any(col in str(row).lower() for col in ["code", "item", "description"]):
                            continue
                        if len(row) >= 6:
                            try:
                                rows.append({
                                    "Code Name": row[0].strip() if row[0] else "",
                                    "Item Code": row[1].strip() if row[1] else "",
                                    "Description": row[2].strip() if row[2] else "",
                                    "Quantity / Unit Type": row[3].split("/")[0].strip() if row[3] else "",
                                    "Unit Price (EGP)": clean_numeric_value(row[4]),
                                    "Total Sales Amount (EGP)": clean_numeric_value(row[5])
                                })
                            except Exception as e:
                                logging.warning(f"Skipping row due to error: {e}")
                                continue
    except Exception as e:
        logging.error(f"Table extraction failed: {e}")
    return rows

def extract_value(text, keyword):
    if not text:
        return "N/A"
    pattern = fr"{re.escape(keyword)}[^\S\r\n]*[:—=]?\s*([^\n|]+)"
    match = re.search(pattern, text)
    return match.group(1).strip() if match else "N/A"

def extract_numeric_value(text, keyword):
    if not text:
        return 0.0
    pattern = fr"{re.escape(keyword)}\s*\(?EGP\)?\s*[:=]?\s*([\d,]+\.?\d*)"
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        try:
            return float(match.group(1).replace(",", ""))
        except ValueError:
            return 0.0
    return 0.0

# ========== PDF EXTRACTION ==========
def extract_invoice_line_items(pdf_path):
    text = extract_text(pdf_path)
    base_data = {
        "Status": extract_value(text, "Status"),
        "Submission Date": extract_value(text, "Submission Date"),
        "Issuance Date": extract_value(text, "Issuance Date"),
        "Internal ID": extract_value(text, "Internal ID"),
        "Issuer": extract_value(text, "Taxpayer Name"),
        "Recipients": extract_value(text.split("Recipients")[-1], "Taxpayer Name"),
    }
    line_items = extract_table_data(pdf_path)
    return pd.DataFrame([{**base_data, **item} for item in line_items])

def extract_invoice_summary(pdf_path):
    text = extract_text(pdf_path).replace("|", "").replace("—", ":").replace("–", ":")
    summary = {
        "Status": extract_value(text, "Status"),
        "Submission Date": extract_value(text, "Submission Date"),
        "Issuance Date": extract_value(text, "Issuance Date"),
        "Internal ID": extract_value(text, "Internal ID"),
        "Issuer": extract_value(text, "Taxpayer Name"),
        "Recipients": extract_value(text.split("Recipients")[-1], "Taxpayer Name"),
        "Code Name": "MNOs Services" if "MNOs Services" in text else "N/A",
        "Item Code": extract_value(text, "Item Code"),
        "Description": extract_value(text, "Description"),
        "Quantity / Unit type": extract_value(text, "Quantity/ Unit Type"),
        "Unit price (EGP)": extract_numeric_value(text, "Unit Price"),
        "Total Sale Amount (EGP)": extract_numeric_value(text, "Total Sales Amount"),
        "Total Sales": extract_numeric_value(text, "Total Sales"),
        "Total Discount": extract_numeric_value(text, "Total discount"),
        "Total Item discount": extract_numeric_value(text, "Total Items Discount"),
        "Value added tax": extract_numeric_value(text, "Value added tax"),
        "Extra invoice discount": extract_numeric_value(text, "Extra Invoice Discounts"),
        "Total amount": extract_numeric_value(text, "Total Amount")
    }
    return pd.DataFrame([summary])

# ========== ROUTES ==========
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        files = request.files.getlist("pdf_files")
        if not files or all(file.filename == '' for file in files):
            return render_template("index.html", error="No files selected")
            
        line_items_list = []
        summary_list = []

        for file in files:
            if file and file.filename.lower().endswith(".pdf"):
                try:
                    pdf_path = os.path.join(UPLOAD_FOLDER, file.filename)
                    file.save(pdf_path)
                    
                    df_line = extract_invoice_line_items(pdf_path)
                    if not df_line.empty:
                        df_line["Source File"] = file.filename
                        line_items_list.append(df_line)
                    else:
                        logging.warning(f"No line items extracted from {file.filename}")

                    df_summary = extract_invoice_summary(pdf_path)
                    if not df_summary.empty:
                        df_summary["Source File"] = file.filename
                        summary_list.append(df_summary)
                    else:
                        logging.warning(f"No summary extracted from {file.filename}")

                except Exception as e:
                    logging.error(f"Error processing {file.filename}: {e}")
                    continue

        if not line_items_list or not summary_list:
            return render_template("index.html", error="Failed to extract data from PDF files")

        try:
            df_lines = pd.concat(line_items_list, ignore_index=True)
            df_summaries = pd.concat(summary_list, ignore_index=True)

            output_merged = os.path.join(OUTPUT_FOLDER, "invoices_merged.xlsx")
            extra_columns = [col for col in df_summaries.columns if col not in df_lines.columns]
            
            df_merged = df_lines.copy()
            
            for file_name in df_summaries["Source File"].unique():
                df_lines_group = df_merged[df_merged["Source File"] == file_name]
                df_summary_group = df_summaries[df_summaries["Source File"] == file_name]

                if df_lines_group.empty or df_summary_group.empty:
                    continue

                last_index = df_lines_group.index[-1]
                last_summary_row = df_summary_group.iloc[-1][extra_columns]

                for col in extra_columns:
                    df_merged.at[last_index, col] = last_summary_row[col]

            df_merged.to_excel(output_merged, index=False)
            return send_file(
                output_merged,
                as_attachment=True,
                download_name="invoices_merged.xlsx",
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        except Exception as e:
            logging.error(f"Excel generation failed: {e}")
            return render_template("index.html", error=f"Error generating Excel: {str(e)}")

    return render_template("index.html")

@app.after_request
def add_header(response):
    # Disable caching for all routes
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, post-check=0, pre-check=0, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '-1'
    return response

# ========== ENTRY POINT ==========
if __name__ == "__main__":
    # Start on all network interfaces
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
