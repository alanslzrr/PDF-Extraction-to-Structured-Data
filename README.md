#     PDF Extraction to Structured Data

![GIF demonstrating the Streamlit process, Excel, and JSON output](placeholder-for-your-gif.gif)

## Introduction
The PDF Certificate Processor is a sophisticated application designed to extract, process, and store tabular data from PDF certificates. It provides a user-friendly interface for uploading PDFs, processes the content to extract relevant information, and offers the extracted data in both Excel and JSON formats.

## Application Components
1. **Main Application (mainextrat.py)**: The core script handling the Streamlit interface and main processing logic.
2. **PyMuPDF Extractor (AreaExtractionTablePyMuPDF.py)**: A utility script for precise text extraction and coordinate identification.
3. **Tabula Visualizer (AreaExtractionTableTabula.py)**: A tool for visualizing table areas in PDFs to aid in extraction coordinate determination.

## Detailed Process Flow

### 1. Application Initialization and User Interface
- The application is built using Streamlit, providing a web-based interface.
- A clear title "PDF Table Certification Extractor" is displayed.
- A file uploader is presented, specifically configured to accept PDF files.

### 2. PDF Upload and Initial Processing
- When a user uploads a PDF file:
  - The file is temporarily saved to the server for processing.
  - `PyMuPDF` (imported as `fitz`) is used to extract the certificate number.
  - A regular expression (`r'C\d+'`) searches for a pattern like 'C' followed by digits within a specific area of the first page (coordinates: 438.10, 62.91, 603.96, 88.87).
  - This certificate number becomes crucial for naming output files and organizing data.

### 3. Table Extraction and Processing
- The application uses `Camelot` for table extraction, employing different strategies for the first page and subsequent pages:
  - First page: Uses `table_area_first_page = "0,680,310,550"` coordinates.
  - Other pages: Uses `table_area_other_pages = "0,700,600,50"` and specific column coordinates.
- The `process_pdf_table` function handles the extraction:
  - It determines which columns to remove based on the `should_remove_column` function.
  - Columns are removed if they contain primarily checkmarks, 'N/A', or empty cells, with exceptions for specific indices.
- The process repeats for each page of the PDF, accumulating data in `all_processed_data`.

### 4. Excel (XLSX) File Generation
- The extracted and processed data is converted to an Excel format using `openpyxl`.
- Each page's data is placed in a separate worksheet named "Page X".
- Instead of saving to a file, the Excel data is stored in a `BytesIO` object, allowing for in-memory processing and direct download.

### 5. Excel to JSON Conversion
- The Excel data is further processed into a structured JSON format:
  - The first page (sheet) is processed differently, extracting key-value pairs.
  - Subsequent pages are processed to extract measurement data, organized by groups.
  - Special handling is implemented for merged cells and specific exclude lines.
- The resulting JSON structure has two main sections:
  - `"datasheet_info"`: General information from the first page.
  - `"measurements"`: Detailed measurement data from subsequent pages.

### 6. Data Storage and Management
- A cumulative JSON file (`certificate_data.json`) is maintained:
  - It stores data for all processed certificates.
  - The file is updated with each new certificate processed, using the certificate number as the key.
  - If the file doesn't exist, it's created; if it exists, it's updated without overwriting existing entries.

### 7. User Download Options
- The application provides two download options for the user:
  1. Excel file: Named `{cert_number}_extracted_tables.xlsx`
  2. JSON file: Named `{cert_number}_extracted_data.json`
- Both files are offered for download directly from the Streamlit interface.

## Key Functions

1. `process_pdf_table(pdf_path, page_number, ignore_na_indices)`:
   - Extracts table data from a specific page of the PDF.
   - Handles different extraction strategies for first and subsequent pages.
   - Removes unnecessary columns based on content criteria.

2. `extract_text_for_filename(pdf_path)`:
   - Extracts the certificate number from the PDF using PyMuPDF.
   - Uses regex to identify the certificate number pattern.

3. `save_data_to_excel(data_per_page)`:
   - Converts processed data to Excel format.
   - Creates a workbook with multiple sheets, one for each PDF page.

4. `process_workbook_from_stream(excel_stream)`:
   - Converts Excel data to a structured JSON format.
   - Handles different processing logic for the first page and subsequent pages.

5. `update_certificate_data(cert_number, workbook_data)`:
   - Updates the cumulative JSON file with new certificate data.
   - Maintains a collection of all processed certificates.

## User Interface
- The Streamlit interface provides a simple and intuitive user experience.
- Users can upload PDF files through a file uploader component.
- Progress and status messages are displayed during processing.
- Download buttons for Excel and JSON outputs are provided upon successful processing.

## Data Processing
- The application employs sophisticated algorithms to handle various PDF layouts and table structures.
- It uses regular expressions and coordinate-based text extraction for precise data identification.
- The process adapts to different page layouts, ensuring accurate data extraction across various certificate formats.

## Output Generation
- Excel Output:
  - Generated using `openpyxl`.
  - Each PDF page is represented as a separate worksheet.
  - Data is structured in a tabular format, preserving the original layout.
- JSON Output:
  - Structured representation of the extracted data.
  - Includes a `datasheet_info` section for general certificate information.
  - Contains a `measurements` section with detailed data from each page.

## Data Storage
- A cumulative JSON file (`certificate_data.json`) serves as a persistent storage solution.
- Each processed certificate's data is stored using its unique certificate number as the key.
- The storage mechanism allows for easy retrieval and management of historical certificate data.

## Error Handling and Resource Management
- The application implements robust error handling, particularly for file operations.
- Temporary files (like the uploaded PDF) are cleaned up after processing to manage resources efficiently.
- Streams and file handlers are properly managed to prevent resource leaks.

## Development and Debugging Tools
1. `AreaExtractionTablePyMuPDF.py`:
   - Utilizes PyMuPDF for precise text extraction with coordinates.
   - Useful for fine-tuning extraction areas and troubleshooting text location issues.

2. `AreaExtractionTableTabula.py`:
   - Employs Camelot to visualize table areas within PDFs.
   - Aids in determining and adjusting correct extraction coordinates.

## Conclusion
The PDF Certificate Processor demonstrates a sophisticated approach to PDF data extraction, processing, and storage. Its flexibility in handling varying PDF structures and data organizations makes it a powerful tool for automating certificate data management. The combination of user-friendly interface, robust processing capabilities, and multiple output formats ensures that it can meet diverse user needs in certificate data extraction and analysis.
