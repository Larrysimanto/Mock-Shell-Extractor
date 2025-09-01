# PDF Title and Footnote Extractor

This Python script extracts titles and footnotes from each page of a specified PDF document. It is designed to parse clinical study reports or similar structured documents where titles follow a consistent pattern (e.g., "Table X.X.X") and footnotes are located at the bottom of the page.

The extracted data—page number, title, and concatenated footnotes—is then saved into a well-structured Excel spreadsheet for easy analysis and review.

## Features

- **Title Detection**: Identifies titles using a configurable regular expression.
- **Footnote Isolation**: Scans a specific portion of the page bottom for footnotes, ignoring common footer text like confidentiality notices and page numbers.
- **Structured Output**: Exports the findings to an Excel file with clear column headers (`Page`, `Title`, `Footnotes`).
- **Error Handling**: Checks for the existence of the input PDF file before processing.

## Prerequisites

Before running the script, you need to have Python 3 installed, along with the following packages:

- `pandas`
- `pdfplumber`
- `openpyxl` (required by pandas to write `.xlsx` files)

## Installation

1.  **Clone the repository or download the script.**

2.  **Install the required Python packages.** It is recommended to use a virtual environment. The dependencies are listed in `requirements.txt`.

    ```bash
    pip install -r requirements.txt
    ```

## How to Run

1.  **Place your PDF** file in the same directory as the script or provide a full path to it.

2.  **Configure the script**: Open `extract_data.py` and modify the variables in the main execution block at the bottom of the file:

    ```python
    # --- Main Execution ---
    if __name__ == "__main__":
        PDF_FILE = "tfl_mock_shells.pdf"  # <-- Your input PDF file
        EXCEL_OUTPUT_FILE = "clinical_study_report_summary.xlsx" # <-- Your desired output file name
        
        report_data = extract_titles_and_footnotes(PDF_FILE)
        
        if report_data:
            save_to_excel(report_data, EXCEL_OUTPUT_FILE)
    ```

3.  **Execute the script** from your terminal:

    ```bash
    python extract_data.py
    ```

4.  **Check the output**: A new Excel file (e.g., `clinical_study_report_summary.xlsx`) will be created in the same directory, containing the extracted data.

## How It Works

### Title Identification
The script reads each page and splits the text into lines. It uses the following regular expression to identify lines that look like titles:

`r'^(Table|Figure|Listing)\s[\d\.]+\s?:.*'`

This pattern matches lines starting with "Table", "Figure", or "Listing", followed by a version number (e.g., 14.1.1) and a colon.

### Footnote Identification
To find footnotes, the script isolates the bottom portion of each page. The area is defined by a threshold; currently, it scans the **bottom 40%** of the page (from the 60% height mark downwards).

Within this cropped area, it looks for lines that:
- Start with "Note:"
- Contain an equals sign (`=`), which is common for abbreviation definitions.

It actively ignores lines containing "Confidential" or page number patterns (e.g., "Page 123") to reduce noise.

## Customization

You can easily customize the script's logic by modifying these sections:

- **Title Regex**: Change the regular expression in the `extract_titles_and_footnotes` function to match the title format of your documents.
- **Footnote Area**: Adjust the `footer_crop_box` calculation to change the vertical percentage of the page scanned for footnotes. For example, changing `page_height * 0.60` to `page_height * 0.75` would scan the bottom 25% of the page.
- **Footnote Keywords**: Add or remove conditions in the footnote identification loop to better match the footnote patterns in your specific PDFs.