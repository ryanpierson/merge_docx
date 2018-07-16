import docx
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree

def iter_footnote_parts(doc):
    for rel in doc.part.rels.values():
        if rel.reltype == RT.FOOTNOTES:
            yield rel.target_part


def get_highest_footnote_id(doc):
    highest_id = 0

    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

    for footnote_part in iter_footnote_parts(doc):
        footnote_xml = footnote_part._blob

        root = etree.fromstring(footnote_xml)

        children = list(root)
        for child in children:
            x = int(child.get(ns + 'id'))
            if x > highest_id:
                highest_id = x

    return highest_id


def get_sub_footnotes(doc):
    highest_id = get_highest_footnote_id(template)

    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

    for footnote_part in iter_footnote_parts(sub_doc):
        footnote_xml = footnote_part._blob

        root = etree.fromstring(footnote_xml)

        children = list(root)
        for child in children:
            if int(child.get(ns + 'id')) > 0:
                # Append it to templates footnotes.xml
                print(child)





# Remove all references to headers and footers.
def handle_footnotes(template, sub_doc):
    highest_id = get_highest_footnote_id(template)


    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

    for footnote_part in iter_footnote_parts(sub_doc):
        footnote_xml = footnote_part._blob

        root = etree.fromstring(footnote_xml)

        children = list(root)
        for child in children:
            if int(child.get(ns + 'id')) > 0:
                # Append it to templates footnotes.xml
                print(child)

        new_footnote_xml = footnote_xml

        footnote_part._blob = new_footnote_xml


