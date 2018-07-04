# Merge_Docx
This script uses the python-docx library to merge files in .docx format. 

Learn more about the .docx file format here:
[OpenOfficeXML](http://officeopenxml.com/WPcontentOverview.php)
and
[WordprocessingML](https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-wordprocessingml-document)

## Notes
This script preserves inline images, page orientations, styles in the merge.<br />
Floating shapes, header & footer references, external hyperlinks are removed.<br />
Enhanced MetaFile '.emf' images are replaced with a placeholder .png image

## Built With
python-docx

## Usage
Make sure you have python-docx installed
```
pip install python-docx
```

Download the source code
```
pip install git+https://github.com/ryanpierson/merge_docx.git
```

Use the code
```
from merge_docx import merge_docx
merge_docx('template.docx', 'sample.docx', 'destination.docx')
```
A merged docx file will be created with the file name 'destination.docx'


## License
This project is licensed under the MIT License - see the LICENSE.md file for details