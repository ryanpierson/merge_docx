import docx

# Remove all references to headers and footers.
def handle_footnotes(doc):
    footnote_ref_tag = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnoteReference'

    # Remove all footnoteReference elements in sectPr
    for para in doc.paragraphs:
        for run in para.runs:
            for elem in run.element.getchildren():
                if elem.tag == footnote_ref_tag:
                    run.element.remove(elem)


