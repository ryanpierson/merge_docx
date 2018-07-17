from setuptools import setup, find_packages

setup(name='merge_docx',
      version='0.1',
      description='merge docx files',
      long_description='m e r g e  d o c x  f i l e s',
      packages=find_packages(),
      install_requires=[
          'python-docx',
      ],
      include_package_data=True,
      package_data = {
      '' : ['*.docx'],
      })