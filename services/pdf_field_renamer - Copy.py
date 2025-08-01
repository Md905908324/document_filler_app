from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, TextStringObject

def rename_pdf_fields(input_pdf_path, output_pdf_path, field_renames):
    """
    Renames fields in a PDF form.

    Args:
        input_pdf_path (str): Path to the input PDF file.
        output_pdf_path (str): Path to save the output PDF file with renamed fields.
        field_renames (dict): A dictionary where keys are original field names
                              and values are new field names.
    Returns:
        bool: True if fields were renamed and saved successfully, False otherwise.
    """
    try:
        reader = PdfReader(input_pdf_path)
        writer = PdfWriter()

        # Copy pages from the reader to the writer
        for page in reader.pages:
            writer.add_page(page)

        # Get the AcroForm dictionary from the reader and attach it to the writer
        # This is crucial for form fields to be recognized in the output PDF
        if "/AcroForm" in reader.trailer["/Root"]:
            writer._root_object.update({NameObject("/AcroForm"): reader.trailer["/Root"]["/AcroForm"]})

        # Iterate through annotations to find and rename fields
        for page in writer.pages:
            if "/Annots" in page:
                for annot_ref in page["/Annots"]:
                    field_object = annot_ref.get_object() # Get the actual dictionary object
                    if "/T" in field_object:
                        original_name = field_object["/T"]
                        if original_name in field_renames:
                            new_name = field_renames[original_name]
                            # Update the /T (Text/Field Name) attribute
                            field_object.update({NameObject("/T"): TextStringObject(new_name)})
        
        # Save the new PDF
        with open(output_pdf_path, "wb") as output_stream:
            writer.write(output_stream)
        return True
    except Exception as e:
        print(f"Error renaming PDF fields: {e}")
        return False