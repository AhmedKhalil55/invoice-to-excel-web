import os
import re
import pdfplumber
import pandas as pd
from flask import Flask, render_template, request, send_file
import logging

# Suppress unnecessary logging
logging.getLogger('pdfminer').setLevel(logging.WARNING)

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "converted"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def clean_numeric_value(value):
    if isinstance(value, str):
        return float(re.sub(r"[^\d.]", "", value.replace(",", "")))
    return value

def extract_text(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        return "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

def extract_table_data(pdf_path):
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table and len(table) > 1:
                for i, row in enumerate(table):
                    is_probable_header = (
                        isinstance(row[0], str) and "Code" in row[0] or
                        isinstance(row[1], str) and "Item" in row[1] or
                        isinstance(row[2], str) and "Description" in row[2]
                    )
                    if is_probable_header:
                        continue
                    if len(row) >= 6:
                        try:
                            rows.append({
                                "Code Name": row[0].strip(),
                                "Item Code": row[1].strip(),
                                "Description": row[2].strip(),
                                "Quantity / Unit Type": row[3].split("/")[0].strip(),
                                "Unit Price (EGP)": clean_numeric_value(row[4]),
                                "Total Sales Amount (EGP)": clean_numeric_value(row[5])
                            })
                        except:
                            continue
    return rows

def extract_value(text, keyword):
    match = re.search(fr"{re.escape(keyword)}[^\S\r\n]*[:—=]?\s*([^\n|]+)", text)
    return match.group(1).strip() if match else "N/A"

def extract_numeric_value(text, keyword):
    pattern = fr"{re.escape(keyword)}\s*\(EGP\)\s*([\d,]+\.?\d*)"
    match = re.search(pattern, text)
    if match:
        try:
            return float(match.group(1).replace(",", ""))
        except:
            return 0.0
    return 0.0

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

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        files = request.files.getlist("pdf_files")
        line_items_list = []
        summary_list = []

        for file in files:
            if file and file.filename.endswith(".pdf"):
                pdf_path = os.path.join(UPLOAD_FOLDER, file.filename)
                file.save(pdf_path)

                df_line = extract_invoice_line_items(pdf_path)
                if not df_line.empty:
                    df_line["Source File"] = file.filename
                    line_items_list.append(df_line)

                df_summary = extract_invoice_summary(pdf_path)
                if not df_summary.empty:
                    df_summary["Source File"] = file.filename
                    summary_list.append(df_summary)

        if not line_items_list or not summary_list:
            return "❌ No valid data extracted from the uploaded files."

        df_lines = pd.concat(line_items_list, ignore_index=True)
        df_summaries = pd.concat(summary_list, ignore_index=True)

        output_merged = os.path.join(OUTPUT_FOLDER, "invoices_merged.xlsx")
        extra_columns = df_summaries.columns[12:]
        columns_to_merge = [col for col in extra_columns if col not in df_lines.columns]
        df_merged = df_lines.copy()

        for file_name in df_summaries["Source File"].unique():
            df_lines_group = df_merged[df_merged["Source File"] == file_name]
            df_summary_group = df_summaries[df_summaries["Source File"] == file_name]

            if df_lines_group.empty or df_summary_group.empty:
                continue

            last_index = df_lines_group.index[-1]
            last_summary_row = df_summary_group.iloc[-1][columns_to_merge]

            for col in columns_to_merge:
                df_merged.at[last_index, col] = last_summary_row[col]

        df_merged.to_excel(output_merged, index=False)
        return send_file(output_merged, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    # ✅ This line ensures it reads the port from Railway or defaults to 5000 locally
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
