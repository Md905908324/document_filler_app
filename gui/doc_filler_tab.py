import os
import tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import shutil
import platform
import subprocess
import pandas as pd
from services.excel_parser import read_excel_data
from utils.formatter import sanitize_key, format_value
from services.word_filler import fill_word_template
from services.pdf_filler import fill_pdf_template, get_pdftk_path  # Import get_pdftk_path
from services.docx_to_pdf import convert_docx_to_pdf
import logging
from fdfgen import forge_fdf
import time

# Log to a local drive
logging.basicConfig(
    filename=os.path.join(os.path.expanduser("~"), "Documents", "document_filler.log"),
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def create_document_filler_tab(notebook):
    frame = ttk.Frame(notebook, padding=10)

    # Variables to store user input/state
    excel_path_var = tk.StringVar()
    output_dir_var = tk.StringVar()
    nrows_var = tk.StringVar(value="")
    status_text_var = tk.StringVar(value="Ready")
    template_paths = []
    loaded_data = {}  # Single dictionary for all data

    # Excel File input
    ttk.Label(frame, text="Excel File:").grid(row=0, column=0, sticky=tk.W)
    excel_entry = ttk.Entry(frame, textvariable=excel_path_var, width=50)
    excel_entry.grid(row=1, column=0, columnspan=2, sticky=tk.EW, padx=(0, 5))
    ttk.Button(frame, text="Browse...",
               command=lambda: browse_excel(excel_path_var, preview_text, nrows_var, status_text_var, frame, loaded_data, notebook)).grid(row=1, column=2, padx=(0, 5))
    ttk.Button(frame, text="Copy Excel Template...",
               command=lambda: copy_excel_template(status_text_var)).grid(row=1, column=3)

    # Template selection listbox
    ttk.Label(frame, text="Templates:").grid(row=2, column=0, sticky=tk.W, pady=(10, 0))
    template_listbox = tk.Listbox(frame, height=5, selectmode=tk.MULTIPLE)
    template_listbox.grid(row=3, column=0, columnspan=4, sticky=tk.EW)

    ttk.Button(frame, text="Add Templates...",
               command=lambda: add_templates(template_paths, template_listbox)).grid(row=4, column=0, pady=5)
    ttk.Button(frame, text="Remove Selected",
               command=lambda: remove_templates(template_paths, template_listbox)).grid(row=4, column=1)

    # Rows to read input
    ttk.Label(frame, text="Rows to Read (blank for all rows):").grid(row=5, column=0, sticky=tk.W, pady=(10, 0))
    ttk.Entry(frame, textvariable=nrows_var, width=10).grid(row=5, column=1, sticky=tk.W)

    # Output folder input
    ttk.Label(frame, text="Output Folder:").grid(row=6, column=0, sticky=tk.W, pady=(10, 0))
    output_entry = ttk.Entry(frame, textvariable=output_dir_var, width=50)
    output_entry.grid(row=7, column=0, columnspan=2, sticky=tk.EW, padx=(0, 5))
    ttk.Button(frame, text="Browse...", command=lambda: browse_output(output_dir_var)).grid(row=7, column=2)

    # Loaded Data and Modify Data Manually
    ttk.Label(frame, text="Loaded Data:").grid(row=8, column=0, sticky=tk.W, pady=(10, 0))
    ttk.Button(frame, text="Modify Data Manually",
               command=lambda: open_manual_data_dialog(loaded_data, preview_text, status_text_var, frame)).grid(row=8, column=3, columnspan=4, sticky=tk.EW, pady=(10, 0))

    preview_text = tk.Text(frame, height=10, state='disabled', wrap='none', width=50)
    preview_text.grid(row=9, column=0, rowspan=2, columnspan=3, sticky=tk.NSEW)
    scrollbar = ttk.Scrollbar(frame, orient='vertical', command=preview_text.yview)
    scrollbar.grid(row=9, column=3, rowspan=2, sticky=tk.NS)
    preview_text.config(yscrollcommand=scrollbar.set)
    preview_text.tag_configure("blue", foreground="blue")

    ttk.Button(frame, text="Generate Documents",
            command=lambda: process_documents(
                excel_path_var.get(), template_paths, output_dir_var.get(), status_text_var, frame, loaded_data
            )).grid(row=11, column=0, columnspan=1, sticky='ew', padx=5, pady=10)

    # Status label
    ttk.Label(frame, textvariable=status_text_var).grid(row=11, column=1, columnspan=5, sticky='ew', padx=5, pady=10)

    # Configure grid weights for resizing
    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(9, weight=1)

    return frame

def browse_excel(excel_path_var, preview_text, nrows_var, status_text_var, frame, loaded_data, parent):
    default_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    path = filedialog.askopenfilename(initialdir=default_dir, filetypes=[("Excel Files", "*.xlsx *.xlsm")])
    if path:
        excel_path_var.set(path)
        progress = ttk.Progressbar(frame, mode='indeterminate')
        progress.grid(row=13, column=0, columnspan=4, pady=5)
        progress.start()
        try:
            nrows_str = nrows_var.get().strip()
            limit = int(nrows_str) if nrows_str else None
            loaded_data.clear()  # Clear existing data
            loaded_data.update(read_excel_data(path, limit, parent=parent))

            # Update preview with loaded data
            update_preview(preview_text, loaded_data)
            status_text_var.set("Excel data loaded and preview updated")
        except Exception as e:
            status_text_var.set(f"Error reading Excel: {str(e)}")
            preview_text.config(state='normal')
            preview_text.delete(1.0, tk.END)
            preview_text.insert(tk.END, f"Error: {e}")
            preview_text.config(state='disabled')
            logging.error(f"Excel read error: {str(e)}")
        finally:
            progress.stop()
            progress.destroy()

def open_manual_data_dialog(loaded_data, preview_text, status_text_var, parent):
    """Open a dialog for adding, editing, and removing key-value pairs."""
    dialog = tk.Toplevel(parent)
    dialog.title("Modify Data Manually")
    dialog.geometry("400x400")
    dialog.transient(parent)
    dialog.grab_set()

    frame = ttk.Frame(dialog, padding=10)
    frame.pack(fill=tk.BOTH, expand=True)

    # Key-value entry
    ttk.Label(frame, text="Key:").grid(row=0, column=0, sticky=tk.W)
    key_entry = tk.StringVar()
    ttk.Entry(frame, textvariable=key_entry, width=30).grid(row=1, column=0, columnspan=2, sticky=tk.EW, padx=(0, 5))

    ttk.Label(frame, text="Value:").grid(row=2, column=0, sticky=tk.W, pady=(5, 0))
    value_entry = tk.StringVar()
    ttk.Entry(frame, textvariable=value_entry, width=30).grid(row=3, column=0, columnspan=2, sticky=tk.EW, padx=(0, 5))

    ttk.Button(frame, text="Add/Edit Key-Value",
               command=lambda: modify_manual_entry(key_entry, value_entry, loaded_data, preview_text, status_text_var)).grid(row=4, column=0, pady=5)
    ttk.Button(frame, text="Refresh",
               command=lambda: refresh_manual_listbox(manual_listbox, loaded_data)).grid(row=4, column=2, pady=5)

    # Manual data listbox
    manual_listbox = tk.Listbox(frame, height=10, width=40, selectmode=tk.SINGLE)
    manual_listbox.grid(row=5, column=0, columnspan=2, sticky=tk.NSEW)
    manual_scrollbar = ttk.Scrollbar(frame, orient='vertical', command=manual_listbox.yview)
    manual_scrollbar.grid(row=5, column=2, sticky=tk.NS)
    manual_listbox.config(yscrollcommand=manual_scrollbar.set)

    # Populate listbox with current loaded data
    for key, value in loaded_data.items():
        manual_listbox.insert(tk.END, f"{key}: {value}")

    # Populate entry fields on selection
    def on_select(event):
        if not manual_listbox.curselection():
            return
        index = manual_listbox.curselection()[0]
        key_value = manual_listbox.get(index).split(": ")
        key_entry.set(key_value[0])
        value_entry.set(key_value[1] if len(key_value) > 1 else "")

    manual_listbox.bind("<<ListboxSelect>>", on_select)

    ttk.Button(frame, text="Remove Selected",
               command=lambda: remove_manual_entry(loaded_data, manual_listbox, preview_text, status_text_var)).grid(row=6, column=0, pady=5)

    # Save data button
    ttk.Button(frame, text="Save Data",
               command=lambda: save_data(loaded_data, status_text_var)).grid(row=6, column=1, pady=5)

    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(5, weight=1)
    dialog.wait_window()

def refresh_manual_listbox(listbox, loaded_data):
    """Refresh the manual listbox with the current loaded_data."""
    listbox.delete(0, tk.END)
    for key, value in loaded_data.items():
        listbox.insert(tk.END, f"{key}: {value}")
    logging.debug("Manual listbox refreshed with loaded data")

def update_preview(preview_text, loaded_data):
    """Update the preview with loaded data."""
    preview_text.config(state='normal')
    preview_text.delete(1.0, tk.END)
    preview_text.insert(tk.END, "Key".ljust(30) + "Value\n" + "-"*60 + "\n")
    for key, value in loaded_data.items():
        preview_text.insert(tk.END, f"{str(key).ljust(30)} ", "blue" if key.startswith("New_") else None)
        preview_text.insert(tk.END, f"{value}\n", "blue" if value.startswith("New_") else None)
    preview_text.config(state='disabled')
    logging.debug("Preview updated with loaded data")

def modify_manual_entry(key_entry, value_entry, loaded_data, preview_text, status_text_var):
    """Add or update a key-value pair in loaded_data."""
    key = key_entry.get().strip()
    value = value_entry.get().strip()
    if not key:
        status_text_var.set("Key cannot be empty")
        messagebox.showerror("Error", "Please enter a key")
        return
    sanitized_key = sanitize_key(key)
    formatted_value = format_value(value)
    if sanitized_key in loaded_data and loaded_data[sanitized_key] == formatted_value:
        status_text_var.set("No change detected")
        update_preview(preview_text, loaded_data)  # Ensure preview updates even if no change
        return
    if sanitized_key in loaded_data:
        del loaded_data[sanitized_key]
    loaded_data[sanitized_key] = formatted_value
    key_entry.set("")
    value_entry.set("")
    update_preview(preview_text, loaded_data)  # Update preview after modification
    status_text_var.set("Key-value pair added/updated")
    logging.debug(f"Modified key-value: {sanitized_key}={formatted_value}")

def remove_manual_entry(loaded_data, manual_listbox, preview_text, status_text_var):
    """Remove selected key-value pair from loaded_data."""
    if not manual_listbox.curselection():
        status_text_var.set("No key-value pair selected")
        messagebox.showerror("Error", "Please select a key-value pair to remove")
        return
    key_value = manual_listbox.get(manual_listbox.curselection()[0]).split(": ")
    key = key_value[0]
    del loaded_data[key]
    manual_listbox.delete(manual_listbox.curselection()[0])
    update_preview(preview_text, loaded_data)
    status_text_var.set("Key-value pair removed")
    logging.debug(f"Removed key-value: {key}")

def save_data(loaded_data, status_text_var):
    """Save the loaded data as CSV, Excel, or TXT."""
    if not loaded_data:
        messagebox.showerror("No Data", "No data to save")
        return

    # Convert to DataFrame
    df = pd.DataFrame(list(loaded_data.items()), columns=['Key', 'Value'])

    # File dialog for saving
    default_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    file_types = [
        ("CSV Files", "*.csv"),
        ("Excel Files", "*.xlsx"),
        ("Text Files", "*.txt")
    ]
    file_path = filedialog.asksaveasfilename(
        initialdir=default_dir,
        filetypes=file_types,
        defaultextension=".csv"
    )
    if not file_path:
        return

    try:
        if file_path.endswith(".csv"):
            df.to_csv(file_path, index=False)
        elif file_path.endswith(".xlsx"):
            df.to_excel(file_path, index=False, engine='openpyxl')
        elif file_path.endswith(".txt"):
            df.to_csv(file_path, sep='\t', index=False)
        status_text_var.set(f"Data saved to: {file_path}")
        logging.debug(f"Data saved to: {file_path}")
    except Exception as e:
        status_text_var.set("Error saving data")
        messagebox.showerror("Error", f"Failed to save data: {str(e)}")
        logging.error(f"Save data error: {str(e)}")

def copy_excel_template(status_text_var):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_dir = os.path.join(script_dir, "..", "assets", "data")
    default_dir = os.path.abspath(default_dir)

    if not os.path.isdir(default_dir):
        default_dir = os.path.join(os.path.expanduser("~"), "Downloads")

    # Select Excel template
    template_path = filedialog.askopenfilename(
        title="Select Excel Template",
        initialdir=default_dir,
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not template_path:
        return

    # Select destination folder
    dest_folder = filedialog.askdirectory(
        title="Select Destination Folder",
        initialdir=os.path.join(os.path.expanduser("~"), "Downloads")
    )

    if not dest_folder:
        return

    try:
        # Validate template is a readable Excel file
        import pandas as pd
        pd.read_excel(template_path, header=None, nrows=1)

        # Get the base name and prepend "Copy of"
        template_name = os.path.basename(template_path)
        dest_name = f"Copy of {template_name}"
        dest_path = os.path.join(dest_folder, dest_name)

        # Check if destination file exists
        if os.path.exists(dest_path):
            response = messagebox.askyesno("File Exists", f"{dest_name} already exists. Overwrite?")
            if not response:
                status_text_var.set("Template copy cancelled")
                return

        # Copy the file
        shutil.copy2(template_path, dest_path)

        # Open the copied file
        if platform.system() == "Windows":
            os.startfile(dest_path)
        elif platform.system() == "Darwin":  # macOS
            subprocess.run(["open", dest_path], check=True)
        else:  # Linux
            subprocess.run(["xdg-open", dest_path], check=True)

        messagebox.showinfo("Success", f"Template copied to:\n{dest_path}\nFile opened for editing")
        status_text_var.set("Excel template copied and opened")
        logging.debug(f"Excel template copied to: {dest_path}")
    except Exception as e:
        status_text_var.set("Error copying template")
        messagebox.showerror("Error", f"Failed to copy or open template: {str(e)}")
        logging.error(f"Template copy error: {str(e)}")

def add_templates(template_paths, listbox):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_dir = os.path.join(script_dir, "..", "assets", "templates")
    default_dir = os.path.abspath(default_dir)

    if not os.path.isdir(default_dir):
        default_dir = os.path.join(os.path.expanduser("~"), "Downloads")

    paths = filedialog.askopenfilenames(
        title="Select Template Files",
        initialdir=default_dir,
        filetypes=[("Templates", "*.docx *.pdf")]
    )

    for path in paths:
        if path not in template_paths:
            template_paths.append(path)
            listbox.insert(tk.END, os.path.basename(path))
            logging.debug(f"Added template: {path}")

def remove_templates(template_paths, listbox):
    indices = listbox.curselection()
    for index in reversed(indices):
        del template_paths[index]
        listbox.delete(index)
    logging.debug("Removed selected templates")

def browse_output(output_dir_var):
    default_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    folder = filedialog.askdirectory(initialdir=default_dir)
    if folder:
        output_dir_var.set(folder)
        logging.debug(f"Selected output directory: {folder}")

def process_documents(excel_path, template_paths, output_dir, status_text_var, frame, loaded_data):
    if not excel_path or not template_paths or not output_dir:
        messagebox.showerror("Missing Input", "Please select all inputs")
        return

    if not loaded_data:
        messagebox.showerror("No Data", "Please load an Excel file or add manual data first")
        return

    progress = ttk.Progressbar(frame, maximum=len(template_paths), mode='determinate')
    progress.grid(row=13, column=0, columnspan=4, pady=5)

    try:
        data = loaded_data.copy()  # Use loaded_data directly

        # Get client number and task for folder naming
        client_number = data.get("Client_Number", "output")
        task = data.get("Task", "output")
        task = task.replace(" ", "_")
        folder_name = f"{task}_{client_number}"
        client_folder = os.path.join(output_dir, folder_name)
        originals_folder = os.path.join(client_folder, "originals")
        os.makedirs(client_folder, exist_ok=True)
        os.makedirs(originals_folder, exist_ok=True)

        for i, template in enumerate(template_paths):
            name = os.path.basename(template)
            base_name = os.path.splitext(name.replace("_template", ""))[0]
            output_name = base_name
            for key, value in data.items():
                sanitized_value = value.replace(" ", "_")
                output_name = output_name.replace(key, sanitized_value)
            
            if template.lower().endswith(".docx"):
                docx_path = os.path.join(originals_folder, f"{output_name}.docx")
                pdf_path = os.path.join(client_folder, f"{output_name}.pdf")
                fill_word_template(template, data, docx_path)
                convert_docx_to_pdf(docx_path, pdf_path)
            elif template.lower().endswith(".pdf"):
                pdf_path = os.path.join(client_folder, f"{output_name}.pdf")
                unflattened_pdf_path = os.path.join(originals_folder, f"{output_name}_unflattened.pdf")
                
                # Prepare FDF data locally
                adjusted_data = {}
                for key, value in data.items():
                    val_str = str(value)
                    if key.lower().startswith("check_box") or "checkbox" in key.lower():
                        adjusted_data[key] = "Yes" if val_str.lower() in ["yes", "true", "on", "1"] else "Off"
                    else:
                        adjusted_data[key] = val_str
                fields = [(key, val) for key, val in adjusted_data.items()]
                fdf_data = forge_fdf("", fields, [], [], [])
                
                with tempfile.NamedTemporaryFile(delete=False, suffix=".fdf") as fdf_file:
                    fdf_file.write(fdf_data)
                    fdf_file_path = fdf_file.name

                try:
                    # Create flattened PDF
                    cmd_flatten = [get_pdftk_path(), template, "fill_form", fdf_file_path, "output", pdf_path, "flatten"]
                    logging.debug(f"Running pdftk command for flattened PDF: {' '.join(cmd_flatten)}")
                    subprocess.run(cmd_flatten, check=True, capture_output=True, text=True)
                    
                    # Create unflattened PDF
                    cmd_unflatten = [get_pdftk_path(), template, "fill_form", fdf_file_path, "output", unflattened_pdf_path]
                    logging.debug(f"Running pdftk command for unflattened PDF: {' '.join(cmd_unflatten)}")
                    subprocess.run(cmd_unflatten, check=True, capture_output=True, text=True)
                    logging.debug(f"Unflattened PDF created at {unflattened_pdf_path}")
                except Exception as e:
                    logging.error(f"PDF processing failed: {str(e)}")
                    raise
                finally:
                    for _ in range(3):
                        try:
                            os.unlink(fdf_file_path)
                            break
                        except OSError:
                            time.sleep(0.1)
            
            progress['value'] = i + 1
            frame.update()

        # Open the output folder
        if platform.system() == "Windows":
            os.startfile(client_folder)
        elif platform.system() == "Darwin":
            subprocess.run(["open", client_folder], check=True)
        else:
            subprocess.run(["xdg-open", client_folder], check=True)

        messagebox.showinfo("Done", f"Files saved to:\n{client_folder}\nWord and PDF originals in: {originals_folder}\nFolder opened")
        status_text_var.set("Documents created and folder opened")
        logging.debug(f"Documents generated in: {client_folder}")
    except Exception as e:
        status_text_var.set("Error during generation")
        messagebox.showerror("Error", str(e))
        logging.error(f"Document generation error: {str(e)}")
    finally:
        progress.destroy()