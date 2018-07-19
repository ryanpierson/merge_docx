import docx
from docx.oxml.shape import CT_Inline

# Inline
import os
# Unique filepath
import random


def handle_floats(template, sub):

    template_drawings = template.element.xpath('//w:drawing')
    sub_drawings = sub.element.xpath('//w:drawing')

    rels = template.part.rels
    #print(rels)

    #x_path = '//a:blip'
    # blip_list holds a list of all of the a:blip elements in the sub_doc
    #blip_list = sub.element.xpath(x_path)
    #print(blip_list)


    #NS = '{http://schemas.openxmlformats.org/drawingml/2006/main}'

    for shape in sub_drawings:
        img = shape.getchildren()[0]
        if not isinstance(img, CT_Inline):
            # then it must be a floater
            blip_list = img.xpath('//a:blip', namespaces={
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
            })
            print(blip_list)



# Add all of the inline images from sub into template.
def handle_inlines(template, sub):
    # rels holds the list of relationships present in the template file.
    rels = template.part.rels
    
    x_path = '//a:blip'
    # blip_list holds a list of all of the a:blip elements in the sub_doc
    blip_list = sub.element.xpath(x_path)

    # shapes holds a list of all the InlineShape objects holding the media
    # of interest
    shapes = sub.inline_shapes

    # Save each image with a unique file path
    for i in range(len(shapes)):

        num_id = random.randint(10000, 100000) 
        
        # Get the rId
        shape = shapes[i]
        rId = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
        
        # ImagePart object
        image_part = sub.part.related_parts[rId]

        # Overwrite .emf images
        if image_part.filename[-4:] == '.emf':
            image_blob = EMF_PH
            unique_path = 'image' + str(num_id) + '.' + 'png'

        else:
            # Image object
            actual_image = image_part.image

            # Binary image
            image_blob = actual_image.blob

            # Compose unique filepath str
            unique_path = 'image' + str(num_id) + '.' + actual_image.ext

        # Save blob as image file
        fh = open(unique_path, "wb")
        fh.write(image_blob)
        fh.close()

        # Create the new relationship in the template file for sub_doc's media 
        rId, img = template.part.get_or_add_image(unique_path)
        # rId holds the new rId for the added image
        # img holds the Image object (unused)

        elem = blip_list[i]

        # Sync sub_doc's relationship ID's with the relationships created in
        # the template file
        elem.set('{http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships}embed', rId)

        # Remove the image file because it's been copied into the template 
        # file's media folder already.
        os.remove(unique_path)