# services/pdf_field_reader.py (Assumed existing content for context)
import subprocess, os
from pypdf import PdfReader

def extract_fields_pdftk(pdf_path):
    """
    Extracts field information using pdftk.
    Returns a list of (field_name, field_info_dict) tuples.
    """
    # Path to pdftk executable
    pdftk_path = r"./tools/pdftk.exe" #external pdftk_path instead of directly installing.
    if not os.path.exists(pdftk_path):
        print(f"Error: pdftk executable not found at {pdftk_path}")
        return []
        
    try:
        cmd = [pdftk_path, pdf_path, "dump_data_fields"]
        result = subprocess.run(cmd, capture_output=True, text=True, check=True, encoding='utf-8', errors='ignore')
        output_lines = result.stdout.splitlines()

        fields_data = {}
        current_field_name = None
        for line in output_lines:
            if line.startswith("FieldName:"):
                current_field_name = line.split(":", 1)[1].strip()
                fields_data[current_field_name] = {}
            elif current_field_name and ":" in line:
                key, value = line.split(":", 1)
                fields_data[current_field_name][key.strip()] = value.strip()
        
        return [(name, info) for name, info in fields_data.items()]
    except Exception as e:
        print(f"Error with pdftk extraction: {e}")
        return []

def extract_fields_pypdf(pdf_path):
    """
    Extracts field names using pypdf.
    Returns a list of unique field names.
    """
    try:
        reader = PdfReader(pdf_path)
        field_names = []
        for page in reader.pages:
            if "/Annots" in page:
                for annot_ref in page["/Annots"]:
                    field_object = annot_ref.get_object()
                    if "/T" in field_object and field_object["/T"]: # Ensure /T exists and is not empty
                        field_names.append(field_object["/T"])
        return list(set(field_names)) # Return unique names
    except Exception as e:
        print(f"Error with pypdf field extraction: {e}")
        return []


# def extract_fields_pdftk(pdf_path):
#     # Path to pdftk executable
#     pdftk_path = r"./tools/pdftk.exe" #external pdftk_path instead of directly installing.

#     # Create a temporary file for the output
#     with tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode='w', encoding='utf-8') as temp_file:
#         output_file = temp_file.name

#     try:
#         # Run the pdftk command with the full path to pdftk.exe
#         result = subprocess.run(
#             [
#                 pdftk_path,  # Full path to pdftk.exe
#                 pdf_path,
#                 "dump_data_fields",
#                 "output",
#                 output_file
#             ],
#             check=True, capture_output=True, text=True
#         )
        
#         # Print the output and error (for debugging)
#         print(result.stdout)
#         print(result.stderr)
        
#         # Parse the fields from the output file
#         fields = []
#         current_field = None
#         with open(output_file, "r", encoding="utf-8") as f:
#             lines = f.read().splitlines()

#         for line in lines:
#             line = line.strip()
#             if not line:
#                 continue
#             if line.startswith("FieldName:"):
#                 if current_field:
#                     fields.append((current_field["name"], current_field))
#                 current_field = {"name": line.split(":", 1)[1].strip()}
#             elif current_field:
#                 if line.startswith("FieldType:"):
#                     current_field["type"] = line.split(":", 1)[1].strip()
#                 elif line.startswith("FieldValue:"):
#                     current_field["value"] = line.split(":", 1)[1].strip()
#                 elif line.startswith("FieldStateOption:"):
#                     current_field.setdefault("values", []).append(line.split(":", 1)[1].strip())
#         if current_field:
#             fields.append((current_field["name"], current_field))

#         return fields
#     except subprocess.CalledProcessError as e:
#         print(f"Error calling pdftk: {e}")
#         print(e.stderr)
#         return []
#     finally:
#         # Ensure the temporary file is deleted
#         for _ in range(3):
#             try:
#                 os.unlink(output_file)
#                 break
#             except OSError:
#                 time.sleep(0.1)

# def extract_fields_pypdf(pdf_path):
#     reader = PdfReader(pdf_path)
#     fields = []

#     for page in reader.pages:
#         if "/Annots" in page:
#             for annot in page["/Annots"]:
#                 obj = annot.get_object()
#                 if "/T" in obj and "/FT" in obj:
#                     fields.append(obj["/T"])

#     return fields
