# Merge_Docx
Merge_Docx uses the python-docx library to merge files in .docx format. 

Learn more about the .docx file format here:
[OpenOfficeXML](http://officeopenxml.com/WPcontentOverview.php)

## Notes
This script preserves inline images, page orientations, styles, and numbering in the merge.<br />
Floating shapes, header & footer references, external hyperlinks are removed.<br />
Enhanced MetaFile '.emf' images are replaced with a placeholder .png image.

## Usage
Download the source code
```
pip install git+https://github.com/ryanpierson/merge_docx.git
```

Sample usage
```python
from merge_docx import merge_docx
merge_docx('template.docx', 'sample.docx', 'destination.docx')
```
A merged docx file will be created with the file name 'destination.docx'

## Explanation
A .docx file is a zip folder containing a number of "parts"<br />
   - Typically UTF encoded XML files, though the ackage may also contain a media folder for images/video<br />
Relationships between parts, i.e. images, footnotes, numbering, styles must be preserved when merging .docx files together<br />

#### Example for inline images
  * This is a Microsoft Word document consisting of a .png image surrounded by two runs of text. Saving this document stores it in the .docx file format.<br />
IMAGE PLACEHOLDER - example docx<br />
  * Unzipping the .docx file and viewing the underlying "word/document.xml" file reveals this (simplified) xml representation:<br />
```xml
<w:document>
    <w:body>
        <w:p>
            <w:r>
                <w:t>Run of text 1</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:drawing>
                    <simplified>
                        <a:blip r:embed=“rId4”>
                            <simplified>
                        </a:blip>
                    <simplified>
                </w:drawing>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>Run of text 1</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>
```

  * The text is stored directly within the "word/document.xml" file, but the image is not. Instead, Microsoft Word uses "relationships" to specify the connection between a source part and a target resource.<br />
  * Instead of embedding the image directly in a binary format, when word displays this file, it looks into the `<a:blip>` element for an `r:embed` attribute. This tells Word that something is supposed to be embedded in the document here, but it has to look up what it is.<br />
      - This is achieved by `r:embed`’s "relationship ID" value, which is `rId4` in the example document.<br />
  * Once the relationship ID is located, Word then looks up the relationship ID’s target in the "word/\_rels/document.xml.rels" file. Below is a simplified version of that relationship file. We can see that rId4 corresponds to a file located at "media/image1.png,"" which is in fact the Proteus logo seen in the example Microsoft Word document.
IMAGE PLACEHOLDER - rels file<br />
  * The use of relationships in Open Office XML allows developers to easily locate document resources without having the parse the document’s content directly. It presents a problem, however, when programmatically merging .docx files.<br />
  * This Python code adds each element in the body of "sub_doc" (the file being merged in) to the body of "merged_doc" (the file being merged into).<br />
IMAGE PLACEHOLDER - Python code<br />
  * Considering the simplified xml representation of the example Microsoft Word document, it is clear that the body of that document has three child elements: a paragraph of text, a paragraph containing an image, and another paragraph of text.<br />
  * If sub_doc was the example .docx file, this script would take these three elements and add them to the end of merged_doc’s “word/document.xml” file. The runs of text would show up no problem because they already exist as a part of the element being copied over. The image, however, would be lost.<br />
  * Why? Consider merged_doc’s "word/\_rels/document.xml.rels" file:<br />
IMAGE PLACEHOLDER - rels2 file<br />
  * When somebody views the merged document in Microsoft Word, it attempts to display the image from sub_doc by looking up its relationship ID followed by displaying the relationship’s target in the specified location, same process as before.<br />
  * The problem is that the target of “rId4” in the merged document is a file called "fontTable.xml."" Surely Microsoft Word can’t display a font table as an image, so the image fails to render.<br />
      - Additionally, the media folder containing the actual .png image file was never copied into the merged document’s media folder, so even if the relationship ID was correct it still couldn’t work.<br />

How to correct for this?
1. Get a list of all the inline images with an XPath query.
   - 
   ```python
   sub_doc.element.xpath(‘//w:p/w:r/w:drawing/wp:inline’)
   ```
2. From here, the rId for each image can be located.
   - 
   ```python
   rId = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
   ```
3. The rId allows us to look into the relationships file and retrieve the ”target” file.
   - 
   ```python
   image_blob = sub_doc.part.related_parts[rId].image.blob
   ```
4. Now that we have the binary content for the image, we can add it into merged_doc’s media folder.
   - 
   ```python
   new_rId = merged_doc.part.get_or_add_image(image_blob)
   ```
   - The image gets added into merged_doc’s media folder, and a relationship is created with the next available rId. This next available rId is returned and stored in new_rId.
5. The final step before we can merge the documents is to update sub_doc’s ”word/document.xml” file with the new image rId. This is done with another XPath query to get a list of all <a:blip> elements.
   - 
   ```python
   blip_list = sub_doc.element.xpath(‘//a:blip’)
   ```
6. Then set the new rId value for each element in blip_list.
   - 
   ```python
   blip_list[i].set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', new_rId)
   ```
7. At this point, the image and its relationship have been added into the merged document, so the images will display properly.



## Built With
python-docx

## License
This project is licensed under the MIT License - see the LICENSE.md file for details