from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, TextStringObject, ArrayObject, DictionaryObject

def rename_pdf_fields(input_pdf_path, output_pdf_path, field_renames):
    """
    Renames fields in a PDF form and ensures the output PDF remains editable.

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

        # Get the AcroForm dictionary from the reader
        acro_form = reader.trailer["/Root"].get("/AcroForm")
        if acro_form:
            # Create a new AcroForm dictionary for the writer
            writer._root_object.update({
                NameObject("/AcroForm"): DictionaryObject({
                    NameObject("/Fields"): ArrayObject(),
                    NameObject("/NeedAppearances"): NameObject("/True")  # Ensure fields are interactive
                })
            })
            writer_acro_form = writer._root_object["/AcroForm"]

            # Copy other AcroForm attributes (e.g., /DR, /DA) to preserve form behavior
            for key, value in acro_form.items():
                if key != "/Fields":  # We'll handle /Fields separately
                    writer_acro_form.update({NameObject(key): value})
        else:
            # If no AcroForm exists, create an empty one with /NeedAppearances
            writer._root_object.update({
                NameObject("/AcroForm"): DictionaryObject({
                    NameObject("/Fields"): ArrayObject(),
                    NameObject("/NeedAppearances"): NameObject("/True")
                })
            })
            writer_acro_form = writer._root_object["/AcroForm"]

        # Map to store field objects for updating /AcroForm/Fields
        field_object_map = {}

        # Iterate through pages to find and rename fields in annotations
        for page in writer.pages:
            if "/Annots" in page:
                for annot_ref in page["/Annots"]:
                    field_object = annot_ref.get_object()
                    if "/T" in field_object:
                        original_name = field_object["/T"]
                        field_object_map[original_name] = field_object
                        if original_name in field_renames:
                            new_name = field_renames[original_name]
                            field_object.update({
                                NameObject("/T"): TextStringObject(new_name)
                            })

        # Update the /AcroForm/Fields array
        if acro_form and "/Fields" in acro_form:
            for field_ref in acro_form["/Fields"]:
                field = field_ref.get_object()
                if "/T" in field:
                    original_name = field["/T"]
                    if original_name in field_object_map:
                        # Ensure the field object in /Fields matches the updated annotation
                        writer_acro_form["/Fields"].append(field_ref)
                        if original_name in field_renames:
                            field.update({
                                NameObject("/T"): TextStringObject(field_renames[original_name])
                            })

        # Save the new PDF
        with open(output_pdf_path, "wb") as output_stream:
            writer.write(output_stream)
        return True
    except Exception as e:
        print(f"Error renaming PDF fields: {e}")
        return False