

import fitz  # PyMuPDF
import pandas as pd
import re

def load_pdf(file_path):
    doc = fitz.open(file_path)
    return doc

def extract_text_and_layout(doc):
    pages = []
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text("text")
        layout = page.get_text("dict")
        pages.append((text, layout))
    return pages

def detect_tables(layout):
    tables = []
    for block in layout["blocks"]:
        if "lines" in block and is_table(block):
            tables.append(block)
    return tables

def is_table(block):
    # Criteria to identify a table block
    # Consider lines, text alignment, and spacing patterns
    if "lines" in block:
        if len(block["lines"]) > 2:  # Assuming a table has more than 2 lines
            return True
    # Check for consistent alignment and spacing to detect borderless tables
    if "lines" in block and len(block["lines"]) > 2 and has_consistent_alignment(block):
        return True
    return False

def has_consistent_alignment(block):
    # Check if the text lines have consistent alignment and spacing
    y_positions = [line["bbox"][1] for line in block["lines"]]
    y_diffs = [y_positions[i+1] - y_positions[i] for i in range(len(y_positions)-1)]
    avg_y_diff = sum(y_diffs) / len(y_diffs)
    for diff in y_diffs:
        if abs(diff - avg_y_diff) > avg_y_diff * 0.5:  # Allow some variation
            return False
    return True

def clean_text(text):
    # Remove unsupported characters
    return re.sub(r'[^\x20-\x7E]', '', text)

def extract_table_data(table_block):
    table_data = []
    for line in table_block["lines"]:
        row = []
        for span in line["spans"]:
            row.append(clean_text(span["text"]))
        table_data.append(row)
    return table_data

def write_to_excel(tables, output_path):
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for idx, table in enumerate(tables):
            df = pd.DataFrame(table)
            df.to_excel(writer, sheet_name=f"Table_{idx+1}", index=False)

def main(pdf_path, output_path):
    doc = load_pdf(pdf_path)
    pages = extract_text_and_layout(doc)
    all_tables = []
    for text, layout in pages:
        tables = detect_tables(layout)
        for table in tables:
            table_data = extract_table_data(table)
            all_tables.append(table_data)
    write_to_excel(all_tables, output_path)

if __name__ == "__main__":
    pdf_path = "test3.pdf"  # Replace with the actual path to your PDF file
    output_path = "output3.xlsx"  # Replace with the desired output path for the Excel file
    main(pdf_path, output_path)