import tkinter as tk
from tkinter import ttk, filedialog, messagebox
# Ensure these imports match your actual services structure
from services.pdf_field_reader import extract_fields_pdftk, extract_fields_pypdf
from services.pdf_field_renamer import rename_pdf_fields # New import

def create_pdf_fields_tab(parent):
    frame = ttk.Frame(parent, padding=10)

    # PDF file selection
    ttk.Label(frame, text="PDF File:").grid(row=0, column=0, sticky=tk.W, pady=(0,5))

    pdf_path_var = tk.StringVar()
    pdf_entry = ttk.Entry(frame, textvariable=pdf_path_var, width=50)
    pdf_entry.grid(row=1, column=0, sticky=tk.EW, padx=(0,5))

    def browse_pdf():
        path = filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF Files", "*.pdf")])
        if path:
            pdf_path_var.set(path)

    browse_btn = ttk.Button(frame, text="Browse...", command=browse_pdf)
    browse_btn.grid(row=1, column=1)

    # Status bar
    status_var = tk.StringVar()
    status_label = ttk.Label(frame, textvariable=status_var, relief=tk.SUNKEN, anchor=tk.W)
    status_label.grid(row=5, column=0, columnspan=2, sticky=tk.EW, pady=(10,0))

    # Notebook to select extraction method
    method_notebook = ttk.Notebook(frame)
    method_notebook.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(10, 5))

    # PDFtk Tab
    pdftk_tab = ttk.Frame(method_notebook)
    method_notebook.add(pdftk_tab, text="PDFtk Fields")
    pdftk_text = tk.Text(pdftk_tab, height=15, state='disabled', wrap=tk.WORD)
    pdftk_text.pack(expand=True, fill=tk.BOTH)

    # PyPDF Tab (will use Treeview)
    pypdf_tab = ttk.Frame(method_notebook)
    method_notebook.add(pypdf_tab, text="PyPDF Fields")

    # Treeview for PyPDF fields
    columns = ("Original Name", "New Name")
    pypdf_tree = ttk.Treeview(pypdf_tab, columns=columns, show="headings", height=15)
    pypdf_tree.heading("Original Name", text="Original Field Name")
    pypdf_tree.heading("New Name", text="New Field Name (Editable)")
    pypdf_tree.column("Original Name", width=200, anchor=tk.W)
    pypdf_tree.column("New Name", width=200, anchor=tk.W)
    pypdf_tree.pack(expand=True, fill=tk.BOTH)

    # Scrollbar for Treeview
    tree_scrollbar = ttk.Scrollbar(pypdf_tree, orient="vertical", command=pypdf_tree.yview)
    pypdf_tree.configure(yscrollcommand=tree_scrollbar.set)
    tree_scrollbar.pack(side="right", fill="y")

    # Dictionary to store field renames for PyPDF Treeview
    # Key: original_name, Value: new_name
    pypdf_field_map = {}

    def set_cell_value(event):
        """Allows in-place editing of Treeview cells."""
        region = pypdf_tree.identify_region(event.x, event.y)
        if region == "cell":
            column = pypdf_tree.identify_column(event.x)
            # Only allow editing the 'New Name' column (column #2 corresponds to index 1)
            if column == '#2':
                item_id = pypdf_tree.identify_row(event.y)
                if not item_id:
                    return

                # Get current value from the treeview item
                current_values = pypdf_tree.item(item_id, 'values')
                original_name = current_values[0]
                current_new_name = current_values[1] if len(current_values) > 1 else ""

                # Get bounding box of the cell
                x, y, width, height = pypdf_tree.bbox(item_id, column)

                # Create an Entry widget for editing
                editor = ttk.Entry(pypdf_tree, justify='left')
                editor.insert(0, current_new_name)
                editor.selection_range(0, tk.END) # Select all text
                editor.focus_force() # Set focus to the entry widget

                def on_edit_confirm(event):
                    new_value = editor.get()
                    pypdf_tree.item(item_id, values=(original_name, new_value))
                    pypdf_field_map[original_name] = new_value # Update the map
                    editor.destroy()

                editor.bind("<Return>", on_edit_confirm) # Save on Enter key
                editor.bind("<FocusOut>", lambda e: editor.destroy()) # Destroy on focus loss
                editor.place(x=x, y=y, width=width, height=height)

    pypdf_tree.bind("<Double-1>", set_cell_value) # Bind double-click to editing

    def extract_fields_and_populate():
        pdf_file = pdf_path_var.get()
        if not pdf_file:
            messagebox.showerror("Error", "Please select a PDF file first.")
            return

        status_var.set("Extracting fields...")

        # Clear previous results
        pdftk_text.config(state='normal')
        pdftk_text.delete(1.0, tk.END)
        for item in pypdf_tree.get_children():
            pypdf_tree.delete(item)
        pypdf_field_map.clear() # Clear the renaming map

        # Extract with PDFtk
        try:
            # Assuming extract_fields_pdftk returns a list of (name, info_dict) tuples
            # fields_pdftk = extract_fields_pdftk(pdf_file)
            fields_pdftk = sorted(extract_fields_pdftk(pdf_file), key=lambda x: x[0])  # Sort by field name            
            if fields_pdftk:
                for name, info in fields_pdftk:
                    line = f"{name} (Type: {info.get('type', 'Unknown')}, Values: {info.get('values', 'N/A')})\n"
                    pdftk_text.insert(tk.END, line)
            else:
                pdftk_text.insert(tk.END, "No form fields found with pdftk.")
        except Exception as e:
            pdftk_text.insert(tk.END, f"Error with pdftk: {str(e)}")
        pdftk_text.config(state='disabled')

        # Extract with PyPDF and populate Treeview
        try:
            # Assuming extract_fields_pypdf returns a list of field names
            fields_pypdf = extract_fields_pypdf(pdf_file)
            if fields_pypdf:
                fields_pypdf.sort(key=str.lower)
                for name in fields_pypdf:
                    pypdf_tree.insert("", tk.END, values=(name, name)) # Initially, new name is same as original
                    pypdf_field_map[name] = name # Initialize map with original name as new name
            else:
                pypdf_tree.insert("", tk.END, values=("No form fields found with pypdf.", ""))
        except Exception as e:
            pypdf_tree.insert("", tk.END, values=(f"Error with pypdf extraction: {str(e)}", ""))

        status_var.set("Extraction complete.")

    def rename_and_save_pdf():
        pdf_file = pdf_path_var.get()
        if not pdf_file:
            messagebox.showerror("Error", "Please select a PDF file first.")
            return

        if not pypdf_field_map:
            messagebox.showinfo("Info", "No fields to rename or fields not extracted.")
            return

        # Filter out fields where the new name is the same as the original,
        # as we only need to pass actual renames to the function.
        actual_renames = {
            old_name: new_name
            for old_name, new_name in pypdf_field_map.items()
            if old_name != new_name
        }

        if not actual_renames:
            messagebox.showinfo("Info", "No field names were changed.")
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            title="Save Renamed PDF As"
        )
        if not output_path:
            return # User cancelled save dialog

        status_var.set("Renaming fields and saving PDF...")
        # if rename_pdf_fields(pdf_file, output_path, actual_renames):
        #     messagebox.showinfo("Success", f"PDF fields renamed and saved to:\n{output_path}")
        #     status_var.set(f"PDF saved to {output_path}")
        # else:
        #     messagebox.showerror("Error", "Failed to rename and save PDF fields.")
        #     status_var.set("Failed to rename PDF fields.")

        if rename_pdf_fields(pdf_file, output_path, actual_renames):
            messagebox.showinfo("Success", f"PDF fields renamed and saved to:\n{output_path}")
            status_var.set(f"PDF saved to {output_path}")
        else:
            messagebox.showerror("Error", "Failed to rename and save PDF fields. Check the input PDF or field names.")
            status_var.set("Failed to rename PDF fields. Check console for details.")            

    # Buttons for actions
    action_frame = ttk.Frame(frame)
    action_frame.grid(row=3, column=0, columnspan=2, pady=(10,0), sticky=tk.EW)

    extract_btn = ttk.Button(action_frame, text="Extract Fields", command=extract_fields_and_populate)
    extract_btn.pack(side=tk.LEFT, padx=(0, 5))

    rename_btn = ttk.Button(action_frame, text="Rename & Save PDF", command=rename_and_save_pdf)
    rename_btn.pack(side=tk.LEFT, padx=(5, 0))

    # Configure grid weights for resizing
    frame.grid_columnconfigure(0, weight=1)
    frame.grid_rowconfigure(2, weight=1) # Allow notebook to expand vertically

    return frame