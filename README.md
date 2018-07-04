# Merge_Docx
Merge_Docx uses the python-docx library to merge files in .docx format. 

Learn more about the .docx file format here:
[OpenOfficeXML](http://officeopenxml.com/WPcontentOverview.php)

## Notes
This script preserves inline images, page orientations, styles in the merge.<br />
Floating shapes, header & footer references, external hyperlinks are removed.<br />
Enhanced MetaFile '.emf' images are replaced with a placeholder .png image

## Usage
Make sure you have python-docx installed
```
pip install python-docx
```

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

## Built With
python-docx

## License
This project is licensed under the MIT License - see the LICENSE.md file for details