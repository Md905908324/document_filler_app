import pandas as pd
import datetime
import re
import tkinter as tk
from tkinter import ttk, messagebox
import logging
import os

# Log to a local drive
logging.basicConfig(
    filename=os.path.join(os.path.expanduser("~"), "Documents", "document_filler.log"),
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

datetime_pattern = re.compile(r"^\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}:\d{2}(?:\.\d+)?)?$")

def select_sheet(excel_file, parent):
    """Prompt user to select multiple sheets from a listbox if the Excel file has multiple sheets."""
    with pd.ExcelFile(excel_file) as xls:
        sheet_names = xls.sheet_names
        logging.debug(f"Found sheets: {sheet_names}")
        if len(sheet_names) == 1:
            return [sheet_names[0]]

    # Create a Toplevel dialog for sheet selection
    dialog = tk.Toplevel(parent)
    dialog.title("Select Sheets")
    dialog.geometry("300x200")
    dialog.resizable(False, False)
    dialog.transient(parent)
    dialog.grab_set()

    frame = ttk.Frame(dialog, padding=10)
    frame.pack(fill=tk.BOTH, expand=True)

    ttk.Label(frame, text="Select sheets (hold Ctrl for multiple):").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
    sheet_listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, height=min(len(sheet_names), 10))
    sheet_listbox.grid(row=1, column=0, columnspan=2, sticky=tk.NSEW, pady=(0, 5))

    for sheet in sheet_names:
        sheet_listbox.insert(tk.END, sheet)

    selected_sheets = []

    def on_ok():
        selection = sheet_listbox.curselection()
        if selection:
            selected_sheets.extend([sheet_names[i] for i in selection])
            dialog.destroy()
        else:
            messagebox.showerror("Error", "Please select at least one sheet", parent=dialog)
            logging.warning("No sheets selected in dialog")

    ttk.Button(frame, text="OK", command=on_ok).grid(row=2, column=0, pady=5)
    ttk.Button(frame, text="Cancel", command=dialog.destroy).grid(row=2, column=1, pady=5)

    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(1, weight=1)
    dialog.wait_window()

    if not selected_sheets:
        logging.error("Sheet selection cancelled")
        raise ValueError("No sheets selected")
    logging.debug(f"Selected sheets: {selected_sheets}")
    return selected_sheets

def resolve_duplicates(duplicates, all_data, parent):
    """Prompt user to select values for duplicate keys with non-identical non-empty values, with a Skip option."""
    resolved_data = {}
    dialog = tk.Toplevel(parent)
    dialog.title("Resolve Duplicate Keys")
    dialog.geometry("400x300")
    dialog.transient(parent)
    dialog.grab_set()

    frame = ttk.Frame(dialog, padding=10)
    frame.pack(fill=tk.BOTH, expand=True)

    canvas = tk.Canvas(frame)
    scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    selections = {}

    for i, (key, occurrences) in enumerate(duplicates.items()):
        current_value = all_data.get(key, "")
        ttk.Label(scrollable_frame, text=f"Current {key} value: '{current_value}'").grid(row=i*2, column=0, sticky=tk.W, pady=5)
        var = tk.StringVar(value="skip")  # Default to Skip
        selections[key] = var
        ttk.Radiobutton(scrollable_frame, text=f"Skip (Keep current value: '{current_value}')", value="skip", variable=var).grid(row=i*2 + 1, column=0, padx=20, sticky=tk.W, pady=2)
        for j, (sheet, value) in enumerate(occurrences):
            if value and value != current_value:  # Only show non-empty, different values
                ttk.Radiobutton(scrollable_frame, text=f"Replace with '{value}' from {sheet}", value=value, variable=var).grid(row=i*2 + 2 + j, column=0, padx=40, sticky=tk.W, pady=2)

        def validate_selection(k=key, v=var):
            selected = v.get()
            if selected == "skip" or selected in [occ[1] for occ in occurrences if occ[1]]:
                return True
            return False

        def on_ok():
            for key, var in selections.items():
                selected_value = var.get()
                if validate_selection(key, var):
                    if selected_value != "skip":
                        resolved_data[key] = selected_value
                    logging.info(f"Resolved {key}: {'kept current' if selected_value == 'skip' else f'replaced with {selected_value}'}")
                else:
                    messagebox.showerror("Error", f"Please select a valid option for key '{key}'", parent=dialog)
                    return
            dialog.destroy()

    ttk.Button(scrollable_frame, text="OK", command=on_ok).grid(row=i*2 + len(occurrences) + 2, column=0, pady=10)
    ttk.Button(scrollable_frame, text="Cancel", command=dialog.destroy).grid(row=i*2 + len(occurrences) + 2, column=1, pady=10)

    scrollable_frame.columnconfigure(0, weight=1)
    dialog.wait_window()

    logging.debug(f"Resolved duplicates: {resolved_data}")
    return resolved_data

def read_excel_data(filepath, nrows=None, parent=None):
    """Read key-value pairs from selected Excel sheets, handling duplicates interactively."""
    logging.debug(f"Reading Excel file: {filepath}, nrows={nrows}")
    
    # Check number of sheets
    with pd.ExcelFile(filepath) as xls:
        sheet_names = xls.sheet_names
    
    # Select sheets if multiple
    sheets_to_read = select_sheet(filepath, parent) if len(sheet_names) > 1 else [sheet_names[0]]
    
    # Initialize data and duplicates tracking
    all_data = {}
    duplicates = {}
    
    for sheet in sheets_to_read:
        # Read the sheet with openpyxl to get evaluated values
        df = pd.read_excel(filepath, sheet_name=sheet, header=None, nrows=nrows, dtype=str, engine='openpyxl')
        
        if df.empty:
            logging.warning(f"Sheet {sheet} is empty")
            continue
        if df.shape[1] < 2:
            logging.error(f"Sheet {sheet} has fewer than two columns")
            raise ValueError(f"Sheet {sheet} must have at least two columns")
        
        # Use only the first two columns
        df = df.iloc[:, :2]
        df.columns = ['Key', 'Value']

        def sanitize_key_name(k):
            if pd.isna(k):
                return ""
            return k.replace(" ", "_").replace("(", "").replace(")", "").replace(",", "").replace("'", "")

        def format_value(val):
            if pd.isna(val):
                return ""
            try:
                dt = pd.to_datetime(val, errors='coerce')
                if pd.notna(dt):
                    return dt.strftime("%m/%d/%Y")
            except:
                pass
            return str(val)

        df['Key'] = df['Key'].apply(sanitize_key_name)
        df['Value'] = df['Value'].apply(format_value)

        # Merge data and track duplicates within selected sheets
        for key, value in zip(df['Key'], df['Value']):
            if key:  # Skip empty keys
                if key in all_data:
                    current_value = all_data[key]
                    if current_value and value and current_value != value:  # Conflict only if both non-empty and different
                        if (sheet, value) not in duplicates.get(key, []):  # Avoid duplicates in occurrences
                            duplicates.setdefault(key, []).append((sheet, value))
                    elif value and not current_value:  # Prefer non-empty value if current is empty
                        all_data[key] = value
                else:
                    all_data[key] = value

    # Resolve duplicates if any, skipping if any value is empty
    if duplicates:
        effective_duplicates = {k: v for k, v in duplicates.items() if all(occ[1] for occ in v) and all_data.get(k)}
        if effective_duplicates:
            resolved_data = resolve_duplicates(effective_duplicates, all_data, parent)
            all_data.update(resolved_data)

    logging.debug(f"Excel data parsed: {all_data}")
    return all_data