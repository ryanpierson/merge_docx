import docx
from docx.oxml.styles import CT_Style

# Add all of the styles from sub into templates 'styles.xml' file. Overwrite
# the style if it already exists.
def handle_styles(template, sub):
    template_styles = template.styles.element
    sub_styles = sub.styles.element

    for s in sub_styles:
        if isinstance(s, CT_Style):
            template_styles.append(s)