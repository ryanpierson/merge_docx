from docx import Document
import docx

NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


# Remove all hyperlinks from the document. This allows the generation of 
# uncorrupted docx files from documents that have floating images.
def handle_numbers(template, sub):

    # Get template's numbering.xml file
    try:
        template_numbering_part = template.part.numbering_part.numbering_definitions._numbering
    except:
        return
    # Get a list of template's abstractNum elements
    template_abstract_list = template_numbering_part.xpath('//w:abstractNum')
    # Get a list of temlate's num elements
    template_num_list = template_numbering_part.xpath('//w:num')

    # Do the same thing with sub
    try:
        sub_numbering_part = sub.part.numbering_part.numbering_definitions._numbering
    except:
        return
    sub_abstract_list = sub_numbering_part.xpath('//w:abstractNum')
    sub_num_list = sub_numbering_part.xpath('//w:num')


    # Find template's highest abstractNumId
    template_highest_abstract_id = 0
    for elem in template_abstract_list:
        abstract_id = int(elem.get(NS + 'abstractNumId'))
        if abstract_id > template_highest_abstract_id:
            template_highest_abstract_id = abstract_id

    # Increment template's highest abstractNumId. This allows abstractNum 
    # elements from sub to be merged in without having conflicting ids
    template_highest_abstract_id += 1

    # Loop over each abstractNum element in sub's numbering part
    for elem in sub_abstract_list:
        # Get the abstractNumId of the current element
        abstract_id = int(elem.get(NS + 'abstractNumId'))
        # Increment the current abstractNumId by the value of template's
        # highest abstractNumId. This ensures each abstractNum element being
        # merged in won't have a conflicting id
        new_id = abstract_id + template_highest_abstract_id
        # Set the abstractNum element's new id
        elem.set(NS + 'abstractNumId', str(new_id))
        # Add the abstractNum element to template's numbering part.
        template_numbering_part.append(elem)


    # Find template's highest numId
    template_highest_num_id = 0
    for elem in template_num_list:
        num_id = int(elem.get(NS + 'numId'))
        if num_id > template_highest_num_id:
            template_highest_num_id = num_id


    # Loop over each num element in sub's numbering part
    for elem in sub_num_list:
        # Get children (particularly abstractNumId) from the num element
        children = elem.getchildren()
        for child in children:
            # Get the abstractNumId element's id value from the num element
            if child.tag == NS + 'abstractNumId':

                abstract_id_val = int(child.get(NS + 'val'))
                if abstract_id_val is not None:
                    # Increment the abstractNumId reference value by an amount equal to 
                    # template's highest abstract abstractNum id value. This is done to 
                    # match the offset created by incrementing sub's abstractNum's id 
                    # values earlier to maintain sub's numbering references.
                    new_abstract_id = abstract_id_val + template_highest_abstract_id
                    child.set(NS + 'val', str(new_abstract_id))


        # Get the current element's numId
        num_id = int(elem.get(NS + 'numId'))

        # Increment the current numId by an amount equal to template's highest
        # numId. This offset will have to be corrected in sub's document part.
        new_id = num_id + template_highest_num_id

        # Set the current num element's new id
        elem.set(NS + 'numId', str(new_id))
        # Add the num element to template's numbering part.
        template_numbering_part.append(elem)


    # Get each numId element in sub's document.xml file
    sub_doc_num_id_list = sub.part.element.xpath('//w:numId')

    # Loop over each numId element in sub's document part
    for elem in sub_doc_num_id_list:
        # Get the current element's id value
        num_id = int(elem.get(NS + 'val'))
        # Increment the current numId by an amount equal to template's highest
        # numId to match the offset created earlier by incrementing sub's 
        # num elements' ids
        new_id = num_id + template_highest_num_id
        # Set the current numId element's new id value
        elem.set(NS + 'val', str(new_id))



    # All abstractNum elements must appear before any num element
    # Get post-merged template's numbering.xml file
    final_numbering_part = template.part.numbering_part.numbering_definitions._numbering
    # Get a list of post-merged template's abstractNum elements
    final_abstract_list = template_numbering_part.xpath('//w:abstractNum')
    # Get a list of post-temlate's num elements
    final_num_list = template_numbering_part.xpath('//w:num')

    # Keep a list of all abstract elements
    abstract_list = []
    # Remove each abstract element from the merged numbering part
    for elem in final_abstract_list:
        abstract_list.append(elem)
        parent = elem.getparent()
        parent.remove(elem)

    # Keep a list of all num elements
    num_list = []
    # Remove each num element from the merged numbering part
    for elem in final_num_list:
        num_list.append(elem)
        parent = elem.getparent()
        parent.remove(elem)

    #Re-add each abstract element
    for elem in abstract_list:
        final_numbering_part.append(elem)
    #Re-add each num element
    for elem in num_list:
        final_numbering_part.append(elem)

