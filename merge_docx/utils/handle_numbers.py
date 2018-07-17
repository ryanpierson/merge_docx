from docx import Document
import docx

# Remove all hyperlinks from the document. This allows the generation of 
# uncorrupted docx files from documents that have floating images.
def handle_numbers(template, sub):
    print(template)
    print(sub)