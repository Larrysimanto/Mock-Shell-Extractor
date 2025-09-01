import pdfplumber
import pandas as pd
import re
import os

def extract_title_from_page(page):
    """
    Extracts the title from a single PDF page.
    """
    text_lines = page.extract_text().split('\n')
    for line in text_lines:
        if re.match(r'^(Table|Figure|Listing)\s[\d\.]+\s?.*', line):
            return line.strip()
    return ""

def extract_footnotes_from_page(page):
    """
    Extracts footnotes from a single PDF page by identifying content blocks
    that exist between a horizontal separator line and the next major title.
    
    Args:
        page: A pdfplumber page object.

    Returns:
        list: A list of strings, where each string is a line from any
              footnote block found on the page.
    """
    # 1. Define a pattern for major titles (e.g., "Table 1.1", "Figure 2.3")
    title_pattern = re.compile(r'^(Table|Figure)\s+\d+(\.\d+)*')
    
    # 2. Find all "anchors" on the page: horizontal lines and titles.
    # Get all horizontal lines with their y-coordinate.
    lines = [{'type': 'line', 'y': l['y0']} for l in page.lines if l['y0'] == l['y1']]
    
    # Get all titles with their top y-coordinate using pdfplumber's search.
    # page.search is more robust for finding coordinates of text.
    titles = [{'type': 'title', 'y': t['top']} for t in page.search(title_pattern)]

    # Combine all anchors into one list.
    anchors = lines + titles
    
    # Also add the very bottom of the page as a final anchor to cap the last block.
    anchors.append({'type': 'page_bottom', 'y': page.height})

    if not lines:
        # If there are no separator lines, we can't find any footnotes with this method.
        return []

    # 3. Sort anchors from top to bottom to map the page structure.
    sorted_anchors = sorted(anchors, key=lambda x: x['y'])
    
    all_footnotes_on_page = []

    # 4. Iterate through the map to find blocks that start with a line.
    for i, anchor in enumerate(sorted_anchors):
        if anchor['type'] == 'line':
            # This is the start of a potential footnote block.
            # The block ends at the position of the *next* anchor.
            if i + 1 < len(sorted_anchors):
                start_y = anchor['y']
                end_y = sorted_anchors[i+1]['y']
                
                # Define the precise bounding box for this block.
                bbox = (0, start_y, page.width, end_y)
                
                # Crop, extract, and process the text within this specific block.
                text = page.crop(bbox).extract_text()
                
                if text:
                    found_lines = text.split('\n')
                    for line in found_lines:
                        clean_line = line.strip()
                        if clean_line and not re.match(r'^Page\s*\d+', clean_line, re.IGNORECASE):
                            all_footnotes_on_page.append(clean_line)

    return all_footnotes_on_page

def extract_footnotes_with_line(page):
    """
    Extracts footnotes from a single PDF page by first finding a
    horizontal separator line and then parsing the text below it.
    """
    horizontal_lines = [line for line in page.lines if line['y0'] == line['y1']]

    if not horizontal_lines:
        return {}

    separator_line = max(horizontal_lines, key=lambda line: line['y0'])
    separator_y = separator_line['y0']

    # Find titles below the separator line
    text_lines = page.extract_text(y_start=separator_y).split('\n')
    next_title_y = page.height
    for line in text_lines:
        if re.match(r'^(Table|Figure|Listing)\s[\d\.]+\s?:.*', line):
            # Find the y-coordinate of this line
            for char in page.chars:
                if char['text'] == line[0]:
                    next_title_y = char['top']
                    break
            break

    bbox = (0, separator_y, page.width, next_title_y)
    footnote_area = page.crop(bbox)
    text = footnote_area.extract_text()

    if not text:
        return {}

    footnote_pattern = re.compile(r'^\((\d+)\)(.*)')
    
    footnotes = {}
    current_source_id = None
    lines = text.split('\n')

    for line in lines:
        match = footnote_pattern.match(line.strip())
        if match:
            source_id = int(match.group(1))
            footnote_text = match.group(2).strip()
            footnotes[source_id] = footnote_text
            current_source_id = source_id
        elif current_source_id is not None and line.strip():
            footnotes[current_source_id] += ' ' + line.strip()
            
    return footnotes

def extract_footnotes_with_line2(page):
    """
    Extracts footnotes from a single PDF page by identifying content blocks
    that exist between a horizontal separator line and the next major title.
    It assigns sequential numbers to each found line to create a dictionary.
    """
    # 1. Define a pattern for major titles (Table/Figure/Listing X.X:)
    title_pattern = re.compile(r'^(Table|Figure|Listing)\s+[\d\.]+\s?:.*', re.IGNORECASE)
    
    # 2. Find all "anchors" on the page: horizontal lines and titles.
    lines = [{'type': 'line', 'y': l['y0']} for l in page.lines if l['y0'] == l['y1']]
    titles = [{'type': 'title', 'y': t['top']} for t in page.search(title_pattern)]
    anchors = lines + titles
    anchors.append({'type': 'page_bottom', 'y': page.height})

    if not lines:
        return {} # Return an empty dictionary as required

    # 3. Sort anchors from top to bottom to map the page structure.
    sorted_anchors = sorted(anchors, key=lambda x: x['y'])
    
    final_footnotes_dict = {}
    # Initialize a counter to assign a unique ID to each footnote line.
    footnote_counter = 1

    # 4. Iterate through the map to find and process each block starting with a line.
    for i, anchor in enumerate(sorted_anchors):
        if anchor['type'] == 'line':
            if i + 1 < len(sorted_anchors):
                start_y = anchor['y']
                end_y = sorted_anchors[i+1]['y']
                
                bbox = (0, start_y, page.width, end_y)
                text = page.crop(bbox).extract_text()
                
                if not text:
                    continue

                # 5. Process each line found in the block.
                found_lines = text.split('\n')
                for line in found_lines:
                    clean_line = line.strip()
                    # Filter out any empty lines or lines that are just page numbers.
                    if clean_line and not re.match(r'^Page\s*\d+', clean_line, re.IGNORECASE):
                        # Assign the current count as the key and the line as the value.
                        final_footnotes_dict[footnote_counter] = clean_line
                        # Increment the counter for the next footnote line.
                        footnote_counter += 1
                            
    return final_footnotes_dict


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
            title = extract_title_from_page(page)
            
            if title:
                footnotes_dict = extract_footnotes_with_line2(page)
                footnotes = [f"({key}) {value}" for key, value in footnotes_dict.items()]
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
    df = df[['Page', 'Title', 'Footnotes']]
    df.to_excel(output_path, index=False)
    print(f"Successfully saved data to '{output_path}'")

def main():
    """
    Main function to run the data extraction and saving process.
    """
    PDF_FILE = "sample_mock.pdf"
    EXCEL_OUTPUT_FILE = "clinical_study_report_summary.xlsx"
    
    report_data = extract_data_from_pdf(PDF_FILE)
    print(report_data)
    
    if report_data:
        save_to_excel(report_data, EXCEL_OUTPUT_FILE)


if __name__ == "__main__":
    main()
