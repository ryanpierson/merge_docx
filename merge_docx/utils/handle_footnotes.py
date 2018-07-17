import docx
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree

NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# Return the content corresponding to a documents footnotes.xml file
def get_footnote_part(doc):
    for rel in doc.part.rels.values():
        if rel.reltype == RT.FOOTNOTES:
            return rel.target_part


# Get the highest footnote id present in a given document.
def get_highest_footnote_id(doc):
    highest_id = 0

    # Get docs's footnotes.xml file
    footnote_part = get_footnote_part(doc)
    footnote_xml = footnote_part._blob

    root = etree.fromstring(footnote_xml)

    children = list(root)
    for child in children:
        id_num = int(child.get(NS + 'id'))
        if id_num > highest_id:
            highest_id = id_num

    return highest_id


# Return a list of footnote elements in a given document.
def get_footnotes(doc):
    footnotes = []

    # Get docs's footnotes.xml file
    footnote_part = get_footnote_part(doc)
    if footnote_part:
        footnote_xml = footnote_part._blob

        root = etree.fromstring(footnote_xml)

        children = list(root)
        for child in children:
            if int(child.get(NS + 'id')) > 0:
                footnotes.append(child)

        return footnotes
    else:
        return None


# Update each footnote id to avoid a conflict with footnotes already 
# existing in template's footnotes.xml file.
def update_footnote_ids(footnotes, highest_id):
    # For each footnote starting at 1, re-number starting at highest_id
    for footnote in footnotes:
        id_num = int(footnote.get(NS + 'id'))
        id_num += highest_id
        footnote.set(NS + 'id', str(id_num))

    return footnotes


# Update all of the documents footnote references to match the updated ids 
# generated in update_footnote_ids(footnotes, highest_id).
def update_document_footnotes(doc, highest_id):
    for para in doc.paragraphs:
        for run in para.runs:
            for elem in run.element.getchildren():
                if elem.tag == NS + 'footnoteReference':
                    id_num = int(elem.get(NS + 'id'))
                    id_num += highest_id
                    elem.set(NS + 'id', str(id_num))


# Remove all references to headers and footers.
def handle_footnotes(template, sub_doc):
    # Get template's footnotes.xml file
    footnote_part = get_footnote_part(template)

    # This list will remain empty if sub_doc has no footnotes
    updated_footnotes = []

    # Get the highest footnote id in the template document. This will be used
    # to ensure that any new footnotes inserted into the document do not share
    # an id with an already existing footnote.
    highest_id = get_highest_footnote_id(template)

    # Return a list of footnote elements in sub_doc
    sub_footnotes = get_footnotes(sub_doc)

    if sub_footnotes:

        # Update sub_docs footnote ids
        updated_footnotes = update_footnote_ids(sub_footnotes, highest_id)

        # Update sub_docs footnote references
        update_document_footnotes(sub_doc, highest_id)


    footnote_xml = footnote_part._blob

    # Convert footnote_xml into etree elements for processing
    root = etree.fromstring(footnote_xml)
        
    # Append sub_docs updated footnotes to root's footnotes element
    for footnote in updated_footnotes:
        root.append(footnote)

    # Serialize the xml to update the document
    new_footnote_xml = etree.tostring(root)

    # Overwrite the old footnotes.xml file
    footnote_part._blob = new_footnote_xml


    


