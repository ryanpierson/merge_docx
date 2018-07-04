from docx import Document
import docx

# Remove all hyperlinks from the document. This allows the generation of 
# uncorrupted docx files from documents that have floating images.
def handle_hyperlinks(file, dest):
    doc = Document(file)

    hyperlinks = doc.element.xpath('//w:hyperlink')

    for link in hyperlinks:
        if 'rId' in link.values()[0]:
            parent = link.getparent()
            parent.remove(link)

    doc.save(dest)