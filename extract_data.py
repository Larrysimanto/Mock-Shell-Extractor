import pdfplumber
import pandas as pd
import re
import os

def extract_title_info(page):
    """
    Extracts the title, id, and population from a single PDF page.
    The title is assumed to be a 3-line block.
    Page number patterns are removed from the title components.
    """
    text_lines = page.extract_text().split('\n')
    
    page_num_pattern = r'Page\s*\d+\s*of\s*\d+'

    for i, line in enumerate(text_lines):
        line = line.strip()
        # Pattern for the first line of a title (e.g., "Table 1.1")
        title_start_pattern = r'^(Table|Figure|Listing)\s+[\d\.]+'
        
        if re.match(title_start_pattern, line, re.IGNORECASE):
            # We found the start of a title. The next 2 lines are part of it.
            if i + 2 < len(text_lines):
                
                # Clean page numbers from the title lines
                output_id = re.sub(page_num_pattern, '', line, flags=re.IGNORECASE).strip()
                title_name = re.sub(page_num_pattern, '', text_lines[i+1], flags=re.IGNORECASE).strip()
                population = re.sub(page_num_pattern, '', text_lines[i+2], flags=re.IGNORECASE).strip()
                
                return {
                    "id": output_id,
                    "title": title_name,
                    "population": population
                }
    return None

def extract_footnotes(page):
    """
    Extracts footnotes from a single PDF page. Footnotes are defined as the
    text appearing after the 3rd horizontal line on the page.
    Lines starting with "programming Note" are ignored.
    """
    # 1. Find all horizontal lines and sort them from top to bottom.
    horizontal_lines = sorted([line for line in page.lines if line['y0'] == line['y1']], key=lambda line: line['y0'])

    # 2. Check if there are at least 3 lines.
    if len(horizontal_lines) < 3:
        print(f"Found {len(horizontal_lines)} horizontal lines on page {page.page_number}.")
        return "" # Not enough lines to find the 3rd one.

    # 3. Get the 3rd horizontal line.
    third_line = horizontal_lines[2]
    third_line_y = third_line['y0']

    # 4. Define the bounding box for the footnote area.
    bbox = (0, third_line_y, page.width, page.height)
    
    # 5. Crop the page and extract text.
    footnote_text = page.crop(bbox).extract_text()

    if not footnote_text:
        return ""

    # 6. Process the extracted text.
    footnotes = []
    found_lines = footnote_text.split('\n')
    for line in found_lines:
        clean_line = line.strip()
        if clean_line and not re.match(r'^Page\s*\d+', clean_line, re.IGNORECASE):
            if '─' not in clean_line and '_' not in clean_line:
                if not clean_line.lower().startswith("programming note"):
                    footnotes.append(clean_line)
                            
    return " | ".join(footnotes)

def extract_data_from_pdf(pdf_path):
    """
    Extracts titles and footnotes from each page of a PDF document.
    """
    if not os.path.exists(pdf_path):
        print(f"Error: The file '{pdf_path}' was not found.")
        return None

    extracted_data = []

    with pdfplumber.open(pdf_path) as pdf:
        print(f"Processing {len(pdf.pages)} pages from '{pdf_path}'...")
        
        for i, page in enumerate(pdf.pages):
            page_number = i + 1
            title_info = extract_title_info(page)
            
            if title_info:
                footnotes = extract_footnotes(page)
                
                extracted_data.append({
                    "Page": page_number,
                    "id": title_info["id"],
                    "title": title_info["title"],
                    "population": title_info["population"],
                    "Footnotes": footnotes if footnotes else "N/A"
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
    df = df[['Page', 'id', 'title', 'population', 'Footnotes']]
    df.to_excel(output_path, index=False)
    print(f"Successfully saved data to '{output_path}'")

def main():
    """
    Main function to run the data extraction and saving process.
    """
    PDF_FILE = "DMC_shell.pdf"
    EXCEL_OUTPUT_FILE = "clinical_study_report_summary.xlsx"
    
    report_data = extract_data_from_pdf(PDF_FILE)
    
    if report_data:
        save_to_excel(report_data, EXCEL_OUTPUT_FILE)


if __name__ == "__main__":
    main()