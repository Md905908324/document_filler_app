import tkinter as tk
from tkinter import ttk
import logging
import os
from gui.doc_filler_tab import create_document_filler_tab
from gui.pdf_field_tab import create_pdf_fields_tab

# Log to a local drive
logging.basicConfig(
    filename=os.path.join(os.path.expanduser("~"), "Documents", "document_filler.log"),
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def main():
    logging.debug("Starting Document Filler Tool")
    root = tk.Tk()
    root.title("Document Filler and PDF Tool")
    root.geometry("800x600")

    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True)

    # Add tabs to notebook
    doc_filler_tab = create_document_filler_tab(notebook)
    notebook.add(doc_filler_tab, text="Document Filler")

    pdf_fields_tab = create_pdf_fields_tab(notebook)
    notebook.add(pdf_fields_tab, text="PDF Fields")

    logging.debug("GUI initialized")
    root.mainloop()

if __name__ == "__main__":
    main()