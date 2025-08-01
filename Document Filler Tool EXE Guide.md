# Document Filler Tool README (EXE Version)

Welcome to the **Document Filler Tool**, a standalone application that makes it simple to fill Word and PDF templates with data from Excel files, including macro-enabled workbooks (`.xlsm`). This version comes as an `.exe` file, so you don’t need to worry about installing Python or complex software—just follow this guide to get started! We’ve included detailed steps to ensure you feel confident using the tool, even if you’re new to it.

## Prerequisites
Before you begin, please prepare the following:
- Download the `DocumentFillerTool.exe` file from the provided source (ask your administrator or check the download link you received).
- Place the files `pdftk.exe` and `libiconv2.dll` in the same folder as the `.exe`. These are essential tools for handling PDFs, and you can get them from your project provider if they’re not included.

## Quick Start
1. Locate the `DocumentFillerTool.exe` file on your computer.
2. Double-click it to launch the program. A window will appear with labeled sections and buttons—don’t worry, we’ll walk you through each part step by step!

## Step-by-Step Usage Guide

### 1. Copy an Excel Template
- **What This Does**: This step creates a fresh copy of an Excel template where you can safely enter your data, like client names or dates, without changing the original file.
- **How to Do It**:
  1. In the main window, look for the "Excel File:" section near the top.
  2. Click the **"Copy Excel Template..."** button. A window will pop up asking you to find a template file.
  3. Browse to a folder (it might default to `assets/data` or your `Downloads` folder) and click on an Excel file (e.g., `.xlsx` or `.xlsm`) that fits your task. Then click "Open".
  4. Another window will appear—choose a folder to save the copy (like `Downloads` or a folder you prefer) and click "Select Folder" or "OK".
  5. The program will make a copy (named something like "Copy of YourTemplate.xlsx") and open it in your default spreadsheet program (e.g., Excel).
- **What to Do Next**: In the opened file, look for a macro option (a button like "Run Macro" in Excel) to clear non-formula fields, or type your data manually. Save the file and close it when you’re ready—we’ll use it next!

### 2. Browse and Select the Excel File
- **What This Does**: This loads the data from your edited Excel file into the tool so it can fill your templates with the right information.
- **How to Do It**:
  1. Back in the main window, find the **"Browse..."** button next to "Excel File:" and click it.
  2. A file explorer will open, showing `.xlsx` and `.xlsm` files. Navigate to where you saved your copied and edited Excel file, click it, and hit "Open".
  3. The file path will show up in the "Excel File:" box.
  4. The tool will start reading the data, and a progress bar will appear. When it’s done, the "Loaded Data:" preview area will fill with your keys (e.g., "ClientName") and values (e.g., "Acme Corp").
- **If Something Goes Wrong**: If the preview stays empty or an error message pops up, check the status bar at the bottom. You might need to ensure the file was saved correctly.

### 3. Select Sheets and Rows to Read
- **What This Does**: This lets you pick which parts of your Excel file to use, especially if it has multiple sheets or lots of rows.
- **How to Do It**:
  1. If your Excel file has more than one sheet (like "Sheet1" or "Data"), a prompt might appear asking you to choose a sheet. Click the one with your data and hit "OK".
  2. Find the "Rows to Read" box. Type a number (e.g., `5` for the first 5 rows) if you only want part of the data, or leave it blank to use everything in the sheet.
  3. The preview will update to show just the data you selected.
- **Helpful Hint**: Use this if your file has extra rows (like old data) that you don’t want to include.

### 4. Browse and Select Templates
- **What This Does**: This step lets you choose the Word or PDF templates you want to fill with your data.
- **How to Do It**:
  1. Scroll down to the "Templates:" section and click the **"Add Templates..."** button.
  2. A file explorer will open (it might start in `assets/templates` or `Downloads`). Go to the folder for your task (e.g., "ClientForms" or "Invoices") and click one or more `.docx` or `.pdf` files while holding Ctrl to select multiple.
  3. Click "Open" to add them. The file names will show up in the listbox below.
  4. If you picked the wrong file by mistake, click it in the listbox and click **"Remove Selected"** to take it out.
- **Organize Tip**: Keep your templates in task-specific folders to make this step faster next time.

### 5. Review and Edit Loaded Data
- **What This Does**: This is your chance to check the data and fix any mistakes before creating the final documents.
- **How to Do It**:
  1. Take a look at the "Loaded Data:" preview box. It lists your keys (placeholders like "Date") and values (like "07/14/2025") side by side. Make sure it all looks correct.
  2. To make changes:
     - Click the **"Modify Data Manually"** button. A new window will pop up.
     - In the "Key:" field, type a placeholder (e.g., "Address"), and in "Value:" type the data (e.g., "123 Main St"). Then click **"Add/Edit Key-Value"**.
     - Click **"Refresh"** to update the list. To remove an entry, click it and hit **"Remove Selected"**.
     - If you want to save your edits, click **"Save Data"**, pick a format (`.csv`, `.xlsx`, or `.txt`) in the save dialog, and choose a location.
  3. Or, go back to your Excel file, update the data, save it, and click "Browse..." again to reload.
- **Friendly Advice**: Spend a minute here to catch any typos or missing info—it’ll save you time later!

### 6. Select Output Folder
- **What This Does**: This tells the tool where to save your finished documents.
- **How to Do It**:
  1. Find the "Output Folder:" section and click the **"Browse..."** button.
  2. A folder selection window will open (it defaults to `Downloads`, but you can pick any folder like "My Documents"). Click a folder and hit "Select Folder" or "OK".
  3. After generating, the tool will create a subfolder named `task_clientnumber` (e.g., `Authorization_NY9999`) based on your data. It will contain all filled PDFs, and for Word files, the original `.docx` files will be in an "originals" subfolder.
- **Nice Feature**: The folder will open automatically after generation, so you can check your work right away.

### 7. Generate Documents
- **What This Does**: This creates your final filled documents using the data and templates you’ve prepared.
- **How to Do It**:
  1. Click the **"Generate Documents"** button at the bottom of the window.
  2. A progress bar will show as the tool works its magic. When it’s done, a message will pop up to confirm success, and the output folder will open.
  3. Open the files in the `task_clientnumber` folder to make sure your data is correctly placed in the templates.
- **If Something Goes Wrong**: Check the status bar for a message. If you’re unsure, the log file (`Documents/document_filler.log`) might have more details.

## Key Features and Helpful Notes
- **Data Format**: Your Excel file should have two columns—the first for keys (placeholders like "ClientName"), the second for values (like "Acme Corp"). This is how the tool knows what to replace.
- **Making New Templates**: Use the "All Inputs" Excel file as a guide. Copy its two-column layout to create new templates that work with this tool.
- **Resolve Conflicts**: If you’re using multiple sheets, a conflict resolution feature will help. Just follow the prompts if they appear.
- **Word Templates**:
  - Use `{{key}}` syntax (e.g., `{{ClientName}}`) where you want data to go.
  - Avoid special characters other than `_` in placeholders, as they may cause issues. The tool sanitizes Excel keys but not Word placeholders.
  - The tool keeps the original font and format. For all caps, adjust the font settings in Word (not case) to avoid errors.
- **PDF Templates**:
  - Use a PDF editor to set field names that match your keys (e.g., "ClientName" field = "ClientName" key).
  - The formatting will follow the PDF’s design.

## Troubleshooting
- **Files Missing**: Ensure `pdftk.exe` and `libiconv2.dll` are in the same folder as the `.exe`. If `.xlsm` files don’t work, check with your provider.
- **Errors**: Look at the status bar for hints. The log file (`Documents/document_filler.log`) can provide more information if needed.
- **Slow Performance**: For large files, try running the tool on your local computer instead of a network drive.