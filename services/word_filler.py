from docxtpl import DocxTemplate

def fill_word_template(template_path, context, output_path):
    doc = DocxTemplate(template_path)
    doc.render(context)  # This line often triggers error if context is malformed
    doc.save(output_path)
