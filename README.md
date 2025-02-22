# kcb-statement-etl-cli

A tool designed to Extract, Transform, and Load (ETL) transaction data from KCB Bank statements in PDF format. It automates the process of converting raw bank statements into structured Excel reports, providing daily totals, transaction summaries, and financial insights.

## Features

-   **PDF Extraction**: Parses KCB Bank PDF statements to extract transaction data.
-   **Data Transformation**: Cleans and structures the extracted data for analysis.
-   **Excel Report Generation**: Produces Excel reports with:
    -   Transaction details
    -   Daily totals
    -   Monthly summaries
    -   Overall financial insights

## Requirements

-   Python 3.x
-   Flask
-   tabula-py
-   pandas
-   openpyxl
-   xlsxwriter
-   Java Runtime Environment (JRE) for tabula-py

## Installation

1.  **Clone the repository**:
    
    bash
    
    CopyEdit
    
    `git clone https://github.com/Levi-LMN/kcb-statement-etl-cli.git
    cd kcb-statement-etl-cli` 
    
2.  **Set up a virtual environment** (optional but recommended):
    
    bash
    
    CopyEdit
    
    ``python3 -m venv venv
    source venv/bin/activate  # On Windows, use `venv\Scripts\activate` `` 
    
3.  **Install the required packages**:
    
    bash
    
    CopyEdit
    
    `pip install -r requirements.txt` 
    
    Ensure that you have the Java Runtime Environment (JRE) installed, as `tabula-py` depends on it.
    

## Usage

1.  **Start the Flask application**:
    
    bash
    
    CopyEdit
    
    `python app.py` 
    
    By default, the application will run on `http://127.0.0.1:5000/`.
    
2.  **Upload a KCB Bank PDF statement**:
    
    -   Open your web browser and navigate to `http://127.0.0.1:5000/`.
    -   Use the provided interface to upload your PDF statement.
3.  **Download the generated Excel report**:
    
    After processing, the application will provide a download link for the Excel report containing your transaction data and summaries.
    

## How It Works

1.  **PDF Upload**: Users upload their KCB Bank PDF statements through the web interface.
2.  **Data Extraction**: The application uses `tabula-py` to extract tables from the PDF.
3.  **Data Cleaning**: Extracted data is cleaned and formatted using `pandas`.
4.  **Report Generation**: Structured data is written to an Excel file with multiple sheets, including detailed transactions, daily totals, and monthly summaries. The `xlsxwriter` library is used to apply formatting and create the Excel file.
5.  **Download Link**: Once processing is complete, a download link for the Excel report is provided to the user.

## Notes

-   Ensure that the uploaded PDF statements are in the standard format provided by KCB Bank for accurate extraction and processing.
-   The application creates an `uploads` directory to temporarily store uploaded files. Processed files are removed after the Excel report is generated to save space.

## License

This project is licensed under the MIT License. See the LICENSE file for details.

## Acknowledgments

-   [tabula-py](https://github.com/chezou/tabula-py) for PDF table extraction.
-   pandas for data manipulation.
-   Flask for the web framework.
-   [xlsxwriter](https://xlsxwriter.readthedocs.io/) for creating Excel files.

For more information, visit the [GitHub repository](https://github.com/Levi-LMN/kcb-statement-etl-cli).

