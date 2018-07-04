import docx

# Add all of the styles from sub into templates 'styles.xml' file. Overwrite
# the style if it already exists.
def handle_styles(template, sub):
	template_styles = template.styles.element
	sub_styles = sub.styles.element
	for s in sub_styles:
	    template_styles.append(s)