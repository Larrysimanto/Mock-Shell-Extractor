import pdfplumber
import pandas as pd
import re
import os

def extract_titles_and_footnotes(pdf_path):
    """
    Extracts titles and footnotes from each page of a PDF document.

    The logic identifies titles based on a specific text pattern (e.g., "Table X.X.X:")
    and identifies footnotes by scanning the bottom portion of each page for
    relevant keywords (e.g., "Note:", abbreviations).
    """
    if not os.path.exists(pdf_path):
        print(f"Error: The file '{pdf_path}' was not found.")
        return None

    extracted_data = []

    with pdfplumber.open(pdf_path) as pdf:
        print(f"Processing {len(pdf.pages)} pages from '{pdf_path}'...")
        
        for i, page in enumerate(pdf.pages):
            page_number = i + 1
            page_content = ""
            title = ""
            footnotes = []

            # --- Title Identification Logic ---
            # Titles in this document reliably start with "Table", "Figure", or "Listing"
            # followed by a version number. We use a regular expression to find them.
            text_lines = page.extract_text().split('\n')
            
            for line in text_lines:
                # Regex to find patterns like "Table 14.1.1: ...", "Figure 14.2.1: ...", etc.
                if re.match(r'^(Table|Figure|Listing)\s[\d\.]+\s?:.*', line):
                    title = line.strip()
                    break # Assume one main title per page

            # --- Footnote Identification Logic ---
            # Footnotes are located at the bottom of the page. We crop the bottom
            # 40% of the page (from 60% of the page height to the bottom) to isolate
            # the area where footnotes and definitions live.
            page_height = page.height
            footer_crop_box = (0, page_height * 0.60, page.width, page_height)
            footer_text = page.crop(footer_crop_box).extract_text()

            if footer_text:
                footer_lines = footer_text.split('\n')
                for line in footer_lines:
                    # Ignore the standard confidential footer and page number
                    if "Confidential" in line or re.match(r'^Page\s\d+', line):
                        continue
                    
                    # Footnotes often start with "Note:" or are abbreviation definitions
                    if line.strip().startswith("Note:") or "=" in line:
                         footnotes.append(line.strip())

            # Only add data if we found a title
            if title:
                extracted_data.append({
                    "Page": page_number,
                    "Title": title,
                    "Footnotes": " | ".join(footnotes) if footnotes else "N/A"
                })

    print("Extraction complete.")
    return extracted_data

def save_to_excel(data, output_path):
    """
    Saves the extracted data to an Excel file.
    """
    if not data:
        print("No data was extracted to save.")
        return

    df = pd.DataFrame(data)
    
    # Reorder columns for clarity
    df = df[["Page", "Title", "Footnotes"]]
    
    df.to_excel(output_path, index=False)
    print(f"Successfully saved data to '{output_path}'")

# --- Main Execution ---
if __name__ == "__main__":
    PDF_FILE = "tfl_mock_shells.pdf"
    EXCEL_OUTPUT_FILE = "clinical_study_report_summary.xlsx"
    
    report_data = extract_titles_and_footnotes(PDF_FILE)
    
    if report_data:
        save_to_excel(report_data, EXCEL_OUTPUT_FILE)
