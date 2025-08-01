import os
import subprocess
import tempfile
import time
from fdfgen import forge_fdf
import logging

# Log to a local drive to avoid NAS latency
logging.basicConfig(
    filename=os.path.join(os.path.expanduser("~"), "Documents", "document_filler.log"),
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def get_pdftk_path():
    """Return the path to pdftk.exe relative to the project directory."""
    base_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'tools', 'pdftk.exe')
    logging.debug(f"Using pdftk path: {base_path}")
    if not os.path.exists(base_path):
        logging.error(f"pdftk.exe not found at: {base_path}")
        raise FileNotFoundError(f"pdftk.exe not found at: {base_path}")
    return base_path

def fill_pdf_template(input_pdf_path, output_pdf_path, data_dict):
    """Fill a PDF template with data and save to output path."""
    start_time = time.time()
    logging.debug(f"Filling PDF: {input_pdf_path} -> {output_pdf_path}")
    adjusted_data = {}
    for key, value in data_dict.items():
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
        cmd = [get_pdftk_path(), input_pdf_path, "fill_form", fdf_file_path, "output", output_pdf_path, "flatten"]
        logging.debug(f"Running pdftk command: {' '.join(cmd)}")
        subprocess.run(cmd, check=True, capture_output=True, text=True)
        logging.debug(f"PDF filled in {time.time() - start_time:.2f} seconds")
        return fdf_file_path  # Return the FDF file path for reuse
    except Exception as e:
        logging.error(f"PDF filling failed: {str(e)}")
        raise
    finally:
        for _ in range(3):
            try:
                os.unlink(fdf_file_path)
                break
            except OSError:
                time.sleep(0.1)