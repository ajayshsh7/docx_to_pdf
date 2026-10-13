# Bulk DOCX to PDF Generator

Generate personalized PDF documents in bulk from a Microsoft Word
(`.docx`) template and an Excel (`.xlsx`) workbook. The app detects
template placeholders, lets you map them to spreadsheet columns, creates
one PDF for each valid data row, and packages the results into a ZIP
file.

## Features

-   Generate many personalized PDFs from one Word template.
-   Support multiple placeholders in the same template.
-   Map each placeholder to any column in the Excel worksheet.
-   Choose which columns provide the person's name and ID for filenames.
-   Preview spreadsheet data before generating files.
-   Show generation progress and report rows that fail.
-   Download all successfully generated PDFs as one ZIP file.
-   Preserve much of the Word document's formatting.
-   Replace placeholders in document paragraphs, tables, headers, and
    footers.
-   Sanitize output filenames and avoid overwriting files when duplicate
    names occur.

## How placeholders work

Add placeholders to your Word template using double curly braces. The
name inside the braces must use letters, numbers, and underscores, and
must start with a letter or underscore.

Examples:

``` text
Dear {{FIRST_NAME}},

Your employee ID is {{ID}}.
Your department is {{DEPARTMENT}}.
```

For every placeholder, the app asks you to select the Excel column that
supplies its value. You can use as many placeholders as needed, provided
each placeholder has a matching column mapping.

**Important:** The current app supports `{{PLACEHOLDER}}` syntax. Other
formats such as `«FIRST_NAME»` or `${FIRST_NAME}` are not detected
unless the code is changed.

## Excel workbook format

Create an Excel workbook (`.xlsx`) with one header row and one record
per person. The column names do not have to be identical to the
placeholder names because you can map them in the app.

Example:

  ID     FirstName   Department
  ------ ----------- -------------
  1001   Asha        Finance
  1002   Rahul       Engineering
  1003   Meera       Operations

For the example template above, map `FIRST_NAME` to `FirstName`, `ID` to
`ID`, and `DEPARTMENT` to `Department`.

-   The workbook must contain at least one non-empty data row.
-   Select the worksheet containing your data.
-   The app uses all non-empty rows in the selected worksheet.
-   The name and ID columns selected for filenames must both have values
    for a row to generate successfully.
-   Keep spreadsheet column headers unique.
-   Check that IDs such as `00123` are stored as text in Excel if
    leading zeros must be preserved.

## Requirements

-   Python 3.10 or newer is recommended.
-   LibreOffice must be installed. The app uses its headless
    command-line converter to turn generated Word files into PDFs.
-   Python packages listed in `requirements.txt`.

## Installation

### 1. Download or clone the project

Place `app.py`, `docx_engine.py`, and `requirements.txt` in the same
project folder.

### 2. Create a virtual environment (recommended)

**Windows (PowerShell):**

``` powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
```

**macOS / Linux:**

``` bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install Python dependencies

Run this command from the project folder:

``` bash
python -m pip install -r requirements.txt
```

### 4. Install LibreOffice

Install LibreOffice for your operating system. Make sure the `soffice`
command is available on your `PATH`.

On Windows, the app also checks these common locations:

-   `C:\Program Files\LibreOffice\program\soffice.exe`
-   `C:\Program Files (x86)\LibreOffice\program\soffice.exe`

If LibreOffice is installed elsewhere, add its program directory to
`PATH`.

## Run the application

From the project folder, run:

``` bash
streamlit run app.py
```

Streamlit will print a local URL, usually `http://localhost:8501`. Open
that address in your browser.

## Generate PDFs

1.  Upload your Word template (`.docx`).
2.  Upload your Excel workbook (`.xlsx`).
3.  Select the worksheet containing the records.
4.  Review the detected row count and spreadsheet preview.
5.  Map every detected template placeholder to the appropriate Excel
    column.
6.  Select the columns to use for output filenames.
7.  Click **Generate PDFs**.
8.  Download the generated PDFs as a ZIP file. If any rows fail, review
    the on-screen error report or download it as CSV.

Output filenames follow the pattern:

``` text
FirstName_ID.pdf
```

For example, a row with `FirstName = Asha` and `ID = 1001` produces
`Asha_1001.pdf`. Unsafe filename characters are replaced. If a filename
already exists in the same batch, a row number is appended to avoid
overwriting it.

## Troubleshooting

### No placeholders detected

Make sure the template uses the supported format, such as
`{{FIRST_NAME}}`, rather than `«FIRST_NAME»` or `${FIRST_NAME}`.

### A placeholder remains in the generated document

Every placeholder detected in the template must be mapped to a
spreadsheet column, and the template token must use the supported
syntax. The app reports rows where placeholders remain unresolved.

### LibreOffice was not found

Install LibreOffice and ensure `soffice` is on `PATH`. Restart the
terminal after changing environment variables, then run the app again.

### PDF conversion fails

Open the Word template in LibreOffice to check that it can be opened and
exported manually. Confirm LibreOffice is installed correctly and that
the template is not damaged.

### A row was skipped

Check the error report. Rows without a value in the selected name or ID
column are skipped. Other rows may fail if conversion fails or a
placeholder remains unreplaced.

### Formatting changes after replacement

The app preserves much of the original formatting, but a replaced token
that spans multiple Word formatting runs inherits the first run's style.
Review the generated PDFs before using them for official documents.

## Current limitations

-   Input files: Word `.docx` templates and Excel `.xlsx` workbooks.
-   Output: individual PDF files delivered together in a ZIP download.
-   Placeholder syntax: `{{PLACEHOLDER_NAME}}`.
-   The app processes the selected worksheet; it does not combine
    multiple worksheets.
-   Complex Word objects such as text boxes, shapes, or other content
    outside the document areas handled by `python-docx` may not be
    detected or replaced.
-   LibreOffice is required for PDF conversion.

## Project structure

``` text
docx-to-pdf/
├── app.py              # Streamlit user interface
├── docx_engine.py      # Placeholder replacement and PDF generation
├── requirements.txt    # Python dependencies
└── README.md           # Project documentation
```

## Dependencies

-   [Streamlit](https://streamlit.io/) --- web interface
-   [pandas](https://pandas.pydata.org/) --- spreadsheet data handling
-   [openpyxl](https://openpyxl.readthedocs.io/) --- Excel workbook
    reading
-   [python-docx](https://python-docx.readthedocs.io/) --- Word document
    editing
-   [LibreOffice](https://www.libreoffice.org/) --- DOCX-to-PDF
    conversion

## Privacy note

The app uses temporary working directories for uploaded files and
intermediate documents while processing. Run it locally when handling
sensitive documents, and avoid deploying it publicly without reviewing
your hosting environment, access controls, and data-handling
requirements.

## License

No license is specified in this repository. Add a `LICENSE` file before
distributing the project if you want to define how others may use,
modify, and redistribute it.
