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
```
from merge_docx import merge_docx
merge_docx('template.docx', 'sample.docx', 'destination.docx')
```
A merged docx file will be created with the file name 'destination.docx'

## Explanation
A .docx file is a zip folder containing a number of "parts"<br />
   - Typically UTF encoded XML files, though the ackage may also contain a media folder for images/video<br />
Relationships between parts, i.e. images, footnotes, numbering, styles must be preserved when merging .docx files together<br />

#### Example for inline images
This is a Microsoft Word document consisting of a .png image surrounded by two runs of text. Saving this document stores it in the .docx file format.<br />
IMAGE PLACEHOLDER<br />
Unzipping the .docx file and viewing the underlying “word/document.xml” file reveals this (simplified) xml representation:<br />
IMAGE PLACEHOLDER<br />
The text is stored directly within the “word/document.xml” file, but the image is not. Instead, Microsoft Word uses “relationships” to specify the connection between a source part and a target resource.<br />
Instead of embedding the image directly in a binary format, when word displays this file, it looks into the `<a:blip>` element for an `r:embed` attribute. This tells Word that something is supposed to be embedded in the document here, but it has to look up what it is.<br />
   - This is achieved by `r:embed`’s “relationship ID” value, which is `rId4` in the example document.<br />





## Built With
python-docx

## License
This project is licensed under the MIT License - see the LICENSE.md file for details