import docx
# Orientations
from docx.oxml.text.paragraph import CT_P
from docx.oxml.document import CT_Body

# Preserve orientation. Iterate over each sectPr in the merged document, 
# excluding the last one. The last sectPr element is excluded because it 
# should exist outside of any paragraph element. 
# Refer to http://officeopenxml.com/WPsection.php for explanation.
def handle_sections(doc):
    for sect in doc.element.sectPr_lst[:-1]:
        # If sectPr is not wrapped in a paragraph, wrap it in the previous
        # paragraph element
        if isinstance(sect.getparent(), CT_Body):
            # Loop through the sectPr's previous elements until a paragraph
            # is found
            cp_elem = sect
            while cp_elem.getprevious() is not None:
                cp_elem = cp_elem.getprevious()
                # Once the preceeding paragraph is found, wrap the sectPr
                # within it
                if isinstance(cp_elem, CT_P):
                    para = cp_elem
                    para.set_sectPr(sect)
                    break

    # Remove any sectPr elements that have been copied into a paragraph
    for sect in doc.element.sectPr_lst[:-1]:
        if isinstance(sect.getparent(), CT_Body):
            sect.getparent().remove(sect)