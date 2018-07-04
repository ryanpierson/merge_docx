import docx

# Remove all references to headers and footers.
def handle_headers_footers(doc):
    header_ref_tag = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}headerReference'
    footer_ref_tag = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footerReference'

    # Remove all headerReference & footerReference elements in sectPr
    for sect in doc.element.sectPr_lst:
        for elem in sect.getchildren():
            if elem.tag == header_ref_tag or elem.tag == footer_ref_tag:
                sect.remove(elem)