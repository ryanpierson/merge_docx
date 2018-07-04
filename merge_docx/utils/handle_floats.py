from docx import Document
import docx
from docx.oxml.shape import CT_Inline

# Remove all of the floating layer drawings from the document. This allows 
# the generation of uncorrupted docx files from documents that have floating 
# images.
def handle_floats(file, dest):

    doc = Document(file)

    drawings = doc.element.xpath('//w:drawing')

    for shape in drawings:
        img = shape.getchildren()[0]
        if not isinstance(img, CT_Inline):
            # then it must be a floater
            # delete img
            parent = shape.getparent()
            parent.remove(shape)

    doc.save(dest)