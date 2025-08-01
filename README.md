# Document Filler Tool README

Welcome to the **Document Filler Tool**, a Python-based application designed to streamline the process of filling Word and PDF templates with data from Excel files, including macro-enabled workbooks (`.xlsm`). This tool uses a key-value pair system to replace placeholders in templates with corresponding values, making it easy to generate customized documents. Follow the step-by-step guide below to get started and make the most of this tool.

## Prerequisites
Before using the tool, ensure you have the following installed:
- **Python 3.7 or higher**
- Required Python packages: `pandas`, `openpyxl`, `docxtpl`, `pypdf`, `fdfgen`, `docx2pdf`, and `setuptools<81`
- Additional tools: `pdftk.exe` and `libiconv2.dll` (place them in the `tools/` directory)

### Installation
1. Clone or download the repository to your local machine.
2. Open a terminal or command prompt and navigate to the project directory (e.g., `cd "B:\00 XM NAS upload\0_Matthew\Document Filler Tool\document_filler_app"`).
3. Install the dependencies by running: pip install pandas openpyxl docxtpl pypdf fdfgen docx2pdf "setuptools<81"
4. Ensure `pdftk.exe` and `libiconv2.dll` are copied into the `tools/` folder.

## Step-by-Step Usage Guide

### 1. Copy an Excel Template
- **Purpose**: Start by creating a working copy of an Excel template to input your data.
- **Steps**:
1. Launch the application by running `python main.py` in the terminal.
2. In the main window, locate the "Excel File:" section.
3. Click the **"Copy Excel Template..."** button.
4. A file dialog will open. Navigate to the folder containing your Excel template (default is `assets/data` or `Downloads`) and select the template you want to use (e.g., an `.xlsx` or `.xlsm` file).
5. Choose a destination folder where the copy will be saved (default is `Downloads`).
6. The copied file (prefixed with "Copy of") will automatically open in your default spreadsheet application.
- **Next Action**: Use the macro in the opened file to clear non-formula fields, or manually enter your data. Save the file when finished.

### 2. Browse and Select the Excel File
- **Purpose**: Load the prepared Excel file into the tool to use its data.
- **Steps**:
1. Back in the main window, click the **"Browse..."** button next to "Excel File:".
2. A file dialog will appear, showing `.xlsx` and `.xlsm` files. Select the copied and edited Excel file you just saved.
3. The file path will appear in the "Excel File:" entry field.
4. The tool will load the data, and a progress bar will display during the process. Once complete, the "Loaded Data:" preview will update with the data from the Excel file.
- **Note**: If the data doesn’t load, check the status message or log file (`Documents/document_filler.log`) for errors.

### 3. Select Sheets and Rows to Read
- **Purpose**: Specify which sheets and rows of the Excel file to use.
- **Steps**:
1. If your Excel file has multiple sheets, the tool will prompt you to select which sheet(s) to read (via `read_excel_data` in `services/excel_parser.py`). Follow the on-screen instructions to choose.
2. In the "Rows to Read" field, enter a number to limit the rows (e.g., `10` for the first 10 rows), or leave it blank to read all rows.
3. The preview will reflect the selected data after loading.
- **Tip**: Use this to focus on specific data sets if your Excel file contains multiple entries.

### 4. Browse and Select Templates
- **Purpose**: Choose the Word (`.docx`) or PDF (`.pdf`) templates to fill with data.
- **Steps**:
1. In the "Templates:" section, click the **"Add Templates..."** button.
2. A file dialog will open (default is `assets/templates` or `Downloads`). Navigate to the folder corresponding to your task (e.g., a folder named after the task) and select one or more template files.
3. The selected template names will appear in the listbox.
4. To remove a template, select it in the listbox and click **"Remove Selected"**.
- **Note**: Ensure templates are organized by task for easy selection.

### 5. Review and Edit Loaded Data
- **Purpose**: Verify the data and make adjustments if needed.
- **Steps**:
1. Check the "Loaded Data:" preview text box for any missing or incorrect information.
2. To edit manually:
  - Click the **"Modify Data Manually"** button.
  - A dialog will open where you can add, edit, or remove key-value pairs.
  - Enter a key and value, then click **"Add/Edit Key-Value"**. Use the **"Refresh"** button to update the list.
  - To remove an entry, select it in the list and click **"Remove Selected"**.
  - Click **"Save Data"** to export the adjusted data as a `.csv`, `.xlsx`, or `.txt` file.
3. Alternatively, update the data in the Excel file and re-upload it using the "Browse..." button.
- **Tip**: The preview shows keys (placeholders) and values in a two-column format.

### 6. Select Output Folder
- **Purpose**: Choose where the filled documents will be saved.
- **Steps**:
1. In the "Output Folder:" section, click the **"Browse..."** button.
2. Select a folder (default is `Downloads`).
3. After generating documents, a subfolder named `task_clientnumber` (e.g., `Authorization_NY9999`) will be created.
4. This folder will contain all filled PDFs and, for Word templates, original `.docx` files in an "originals" subfolder.
- **Note**: The folder will open automatically after generation.

### 7. Generate Documents
- **Purpose**: Create the final filled documents.
- **Steps**:
1. Click the **"Generate Documents"** button.
2. A progress bar will show the generation process. Once complete, a success message will appear, and the output folder will open.
3. Check the generated files in the `task_clientnumber` folder.
- **Note**: If an error occurs, review the status message or log file.

## Key Features and Notes
- **Key-Value Pair System**: The tool uses a two-column format in Excel— the first column for keys (placeholders) and the second for values (substitutions). Ensure your data follows this structure.
- **Template Creation**: To make new templates, use the "All Inputs" Excel file as a guide. Copy its format and create new files accordingly.
- **Resolve Conflicts**: A conflict resolution feature assists with multi-sheet insertions. Follow on-screen prompts if conflicts arise.
- **Word Templates**:
- Use `{{key}}` syntax for placeholders (e.g., `{{Name}}`) for Word Documents.
- Word document placeholders are not friendly with any special characters other than _. Sanitize key is built in for excel reading but not for word templates.
- The tool preserves the original format and font. For all caps, use the font feature (not case) and the uppercase option in Word, avoiding errors.
- **PDF Templates**:
- Prepare forms to edit field names, ensuring they match the keys from the Excel data.
- Formatting follows the template field settings.
- **Logging**: Errors and actions are logged to `Documents/document_filler.log` for troubleshooting.
- **Extra Help**: If you run into problems, contact myd2011@stern.nyu.edu for assistance.

## Troubleshooting
- **Missing `.xlsm` Files**: Ensure `openpyxl` is installed. If issues persist, verify `services/excel_parser.py` uses `engine='openpyxl'`.
- **Errors**: Check the status message and log file for details. Common issues include missing dependencies or invalid file formats.
- **Performance**: For large files, test locally rather than on a NAS.
- **Known Potential Error**: Resolve conflicting keys (for multisheet insertion) may cause issues with missing potential replacements due to UI overlay issues. If no desired option shown, just take a note of it and manually replace value.