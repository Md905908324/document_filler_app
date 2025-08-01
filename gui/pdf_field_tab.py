import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from services.pdf_field_reader import extract_fields_pdftk, extract_fields_pypdf

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

    # Notebook to select extraction method
    method_notebook = ttk.Notebook(frame)
    method_notebook.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(10, 5))

    # Text widgets to show fields
    pdftk_text = tk.Text(method_notebook, height=15, state='disabled', wrap='word')
    pypdf_text = tk.Text(method_notebook, height=15, state='disabled', wrap='word')

    method_notebook.add(pdftk_text, text="pdftk Extraction")
    method_notebook.add(pypdf_text, text="pypdf Extraction")

    # Status label
    status_var = tk.StringVar(value="Ready")
    status_label = ttk.Label(frame, textvariable=status_var)
    status_label.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(5,0))



    # Extraction function
    def extract_fields():
        pdf_file = pdf_path_var.get()
        if not pdf_file:
            messagebox.showerror("Error", "Please select a PDF file first.")
            return

        # Clear text fields
        pdftk_text.config(state='normal')
        pdftk_text.delete(1.0, tk.END)
        pypdf_text.config(state='normal')
        pypdf_text.delete(1.0, tk.END)

        status_var.set("Extracting fields...")

        try:
            fields_pdftk = extract_fields_pdftk(pdf_file)
            if fields_pdftk:
                for name, info in fields_pdftk:
                    line = f"{name} (Type: {info.get('type', 'Unknown')}, Values: {info.get('values', 'N/A')})\n"
                    pdftk_text.insert(tk.END, line)
            else:
                pdftk_text.insert(tk.END, "No form fields found with pdftk.")

        except Exception as e:
            pdftk_text.insert(tk.END, f"Error: {str(e)}")

        pdftk_text.config(state='disabled')

        try:
            fields_pypdf = extract_fields_pypdf(pdf_file)
            if fields_pypdf:
                for name in fields_pypdf:
                    pypdf_text.insert(tk.END, f"{name}\n")
            else:
                pypdf_text.insert(tk.END, "No form fields found with pypdf.")

        except Exception as e:
            pypdf_text.insert(tk.END, f"Error: {str(e)}")

        pypdf_text.config(state='disabled')
        status_var.set("Extraction complete.")

    extract_btn = ttk.Button(frame, text="Extract Fields", command=extract_fields)
    extract_btn.grid(row=4, column=0, columnspan=2, pady=(10,0))

    # Configure grid weights for resizing nicely
    frame.columnconfigure(0, weight=1)
    frame.columnconfigure(1, weight=0)
    frame.rowconfigure(2, weight=1)

    return frame
