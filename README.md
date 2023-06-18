# Easy docs from excel
Create many words documents using data from an excel file. Optionally insert images.
## Basic usage
### Prerequisites
Make sure that you have installed all necessary packages in your environment. If needed use: `pip install -r requirements.txt`
### Configuration
See the contents of `config_to_fill.yml` file:
```yml
excel_source_name: Name of your Excel file # Mandatory
sheet_name: Name of the sheet # Mandatory
template_name: Name of the template # Mandatory
dynamic_values: # Optional - original column names can be used
  column_name_in_excel: name_on_template
  some_column: number # <- Number is important if you want to use dynamic images in your template
static_values: # Optional - useful if the same values are needed on many templates
  name: value
images_config: # Optional
  folder_one: # Name of a folder inside 'images' directory
    number_reference: folder # One to many - one folder with many images per one record from Excel
    prefix_on_template: image_ # You can access the images on template with names image_1, image_2 and so on.
    max_number_of_images: 10
  folder_two: # Name of an another folder inside 'images' directory
    number_reference: file # One to one - one image per one record from Excel
    filename_pattern: <<number>>_200m.png # Change everything except of '<<number>>' to match with your filename pattern.
    name_on_template: my_single_image 
```
### Start
Run code from `main.py`
```python
from easy_docs_from_excel.TemplateFiller import TemplateFiller

tf = TemplateFiller.from_yaml('config_to_fill.yml') # Remember to change the name of the file if needed
tf.make_all_docs()
```
## TODO
- Create a better documentation
- Add tests
- Add linting
- Prepare samples
- Publish to PyPI
