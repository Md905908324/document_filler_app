import subprocess
import tempfile
import os
import time
from pypdf import PdfReader

def extract_fields_pdftk(pdf_path):
    # Path to pdftk executable
    pdftk_path = r"./tools/pdftk.exe" #external pdftk_path instead of directly installing.

    # Create a temporary file for the output
    with tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode='w', encoding='utf-8') as temp_file:
        output_file = temp_file.name

    try:
        # Run the pdftk command with the full path to pdftk.exe
        result = subprocess.run(
            [
                pdftk_path,  # Full path to pdftk.exe
                pdf_path,
                "dump_data_fields",
                "output",
                output_file
            ],
            check=True, capture_output=True, text=True
        )
        
        # Print the output and error (for debugging)
        print(result.stdout)
        print(result.stderr)
        
        # Parse the fields from the output file
        fields = []
        current_field = None
        with open(output_file, "r", encoding="utf-8") as f:
            lines = f.read().splitlines()

        for line in lines:
            line = line.strip()
            if not line:
                continue
            if line.startswith("FieldName:"):
                if current_field:
                    fields.append((current_field["name"], current_field))
                current_field = {"name": line.split(":", 1)[1].strip()}
            elif current_field:
                if line.startswith("FieldType:"):
                    current_field["type"] = line.split(":", 1)[1].strip()
                elif line.startswith("FieldValue:"):
                    current_field["value"] = line.split(":", 1)[1].strip()
                elif line.startswith("FieldStateOption:"):
                    current_field.setdefault("values", []).append(line.split(":", 1)[1].strip())
        if current_field:
            fields.append((current_field["name"], current_field))

        return fields
    except subprocess.CalledProcessError as e:
        print(f"Error calling pdftk: {e}")
        print(e.stderr)
        return []
    finally:
        # Ensure the temporary file is deleted
        for _ in range(3):
            try:
                os.unlink(output_file)
                break
            except OSError:
                time.sleep(0.1)

def extract_fields_pypdf(pdf_path):
    reader = PdfReader(pdf_path)
    fields = []

    for page in reader.pages:
        if "/Annots" in page:
            for annot in page["/Annots"]:
                obj = annot.get_object()
                if "/T" in obj and "/FT" in obj:
                    fields.append(obj["/T"])

    return fields
