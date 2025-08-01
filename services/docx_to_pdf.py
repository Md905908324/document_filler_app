# In services/docx_to_pdf.py
import logging
import time
from docx2pdf import convert
import os

logging.basicConfig(
    filename=os.path.join(os.path.expanduser("~"), "Documents", "document_filler.log"),
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def convert_docx_to_pdf(docx_path, pdf_path):
    start_time = time.time()
    logging.debug(f"Converting {docx_path} to {pdf_path}")
    try:
        convert(docx_path, pdf_path)
        logging.debug(f"Converted {docx_path} to {pdf_path} in {time.time() - start_time:.2f} seconds")
    except Exception as e:
        logging.error(f"Word-to-PDF conversion failed: {str(e)}")
        raise