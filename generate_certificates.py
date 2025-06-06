# This script reads certificates.csv and generates Word documents based on certificate_template.docx
# It uses the python-docx library to replace placeholders {NAME}, {DATE}, and {CERT_ID} in the template.

import csv
import os
from docx import Document

TEMPLATE_FILE = "certificate_template.docx"
CSV_FILE = "certificates.csv"
OUTPUT_DIR = "output"


def replace_placeholders(doc, replacements):
    """Replace placeholders in the entire document including tables."""
    # paragraphs
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            for key, val in replacements.items():
                run.text = run.text.replace(f"{{{key}}}", val)

    # tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        for key, val in replacements.items():
                            run.text = run.text.replace(f"{{{key}}}", val)


if __name__ == "__main__":
    if not os.path.exists(TEMPLATE_FILE):
        raise FileNotFoundError(f"Template file '{TEMPLATE_FILE}' not found")
    if not os.path.exists(CSV_FILE):
        raise FileNotFoundError(f"CSV file '{CSV_FILE}' not found")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with open(CSV_FILE, newline="", encoding="utf-8") as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            name = row.get("NAME", "")
            date = row.get("DATE", "")
            cert_id = row.get("CERT_ID", "")
            doc = Document(TEMPLATE_FILE)
            replace_placeholders(doc, {"NAME": name, "DATE": date, "CERT_ID": cert_id})
            output_filename = f"{cert_id}_{name}.docx"
            doc.save(os.path.join(OUTPUT_DIR, output_filename))
            print(f"Created {output_filename}")

