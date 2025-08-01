import tkinter as tk
from tkinter import ttk

class DocumentFillerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Filler Tool")
        self.root.geometry("700x700")

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)

