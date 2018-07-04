from docx import Document
import docx
import os

from .utils.handle_floats import handle_floats
from .utils.handle_hyperlinks import handle_hyperlinks
from .utils.handle_styles import handle_styles
from .utils.handle_inlines import handle_inlines
from .utils.handle_sections import handle_sections
from .utils.handle_headers_footers import handle_headers_footers

# Merge 'file' into 'template'
def merge_docx(template, file, destination):
    # Save with a .docx file extension
    if destination[-5:] != '.docx':
        destination += '.docx'

    template_no_elem = template + 'no_elem.docx'
    file_no_elem = file + 'no_elem.docx'

    # Remove floating shapes and hyperlinks
    handle_floats(template, template_no_elem)
    handle_floats(file, file_no_elem)
    handle_hyperlinks(template_no_elem, template_no_elem)
    handle_hyperlinks(file_no_elem, file_no_elem)

    # Use the first document as a template file. All formatting will be 
    # inherited from this file. All media and relationships from this file are
    # preserved.
    merged_document = Document(template_no_elem)

    # The document to be merged in exists as the second file in the list. Note
    # that this script only merges 2 files at a time.
    sub_doc = Document(file_no_elem)

    # Add each style in sub_styles to merged_styles.
    handle_styles(merged_document, sub_doc)

    # Remove the temporary files generated without floating elements. If a file
    # is being merged into itself, there will only be one document generated 
    # without floats/hyperlinks. Trying to delete both documents will throw an exception,
    # but I don't care as long as the file gets deleted. 
    try:
        os.remove(template_no_elem)
        os.remove(file_no_elem)
    except OSError:
        pass

    # Add the inline images from sub_doc into the merged document
    handle_inlines(merged_document, sub_doc)

    # Merge the document bodies   
    for element in sub_doc.element.body:
        merged_document.element.body.append(element)

    handle_sections(merged_document)

    handle_headers_footers(merged_document)

    # Save the altered file
    merged_document.save(destination)
    