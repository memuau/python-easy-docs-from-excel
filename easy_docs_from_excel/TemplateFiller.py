import os
import yaml
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
from docx2pdf import convert
import pandas as pd
from PIL import Image, ImageOps

class TemplateFiller:
    def __init__(self, 
                 excel_source_name: str, 
                 sheet_name: str, 
                 template_name: str, 
                 dynamic_values: dict|None = None, 
                 static_values: dict|None = None, 
                 images_config: dict|None = None,
                 fillna: str = "-"
                 ):
        self.contexts = self._get_contexts_from_excel(excel_source_name, sheet_name, fillna, dynamic_values, static_values)
        self.template = DocxTemplate(f"templates/{template_name}.docx")
        self.template_name = template_name
        self.images_config = images_config
        if images_config:
            os.makedirs("images", exist_ok=True)

    @classmethod
    def from_yaml(cls, path_to_config: str):
        with open(path_to_config, "r", encoding="utf-8") as file:
            cfg = yaml.full_load(file)
        return cls(**cfg)

    def make_all_docs(self, create_pdfs: bool = False):
        for context in self.contexts:
            self.make_one_doc(context, create_pdfs)

    def make_one_doc(self, context: dict, create_pdfs: bool = False):
        if self.images_config:
            context = self._add_images(context)
        self.template.render(context)
        output_path = f"output\docx\{self.template_name}_{context['number']}.docx"
        self.template.save(output_path)
        if create_pdfs:
            pdf_path = output_path.replace('docx', 'pdf')
            convert(output_path, pdf_path) 

    def _add_images(self, context: dict) -> dict:
        """
        Returns a modified version of the given context
        """
        output = dict() # New elements will be added here 
        for folder_name, params in self.images_config.items():
            if params['number_reference'] == 'file': # One to one relationship
                image = InlineImage(
                    self.template, 
                    f"images\{folder_name}\{params['filename_pattern'].replace('<<number>>', str(context['number']))}",
                    height=Cm(10))
                output[params['name_on_template']] = image
            elif params['number_reference'] == 'folder': # One to many relationship
                folder_path = f"images\{folder_name}/{context['number']}"
                images = os.listdir(folder_path)
                output = {f"{params['prefix_on_template']}{x}": "" for x in range(params["max_number_of_images"])}
                for index, filename in enumerate(images):
                    image_path = f"{folder_path}/{filename}"
                    img = Image.open(image_path)
                    if img.getexif().get(274, 1) != 1: # Rotation
                        print(f"Detected an image that needs to be rotated: {image_path}")
                        new_img = ImageOps.exif_transpose(img)
                        new_img.save(image_path)
                    image = InlineImage(self.template, image_path, height=Cm(10))
                    output[f"{params['prefix_on_template']}{index}"] = image
        return context | output

    def _get_contexts_from_excel(self, 
                             excel_source_name: str, 
                             sheet_name: str, 
                             fillna: str, 
                             dynamic_values: dict|None = None, 
                             static_values: dict|None = None
                             ) -> list[dict]:
        df = pd.read_excel(f"{excel_source_name}.xlsx", sheet_name)
        df = df.fillna(fillna)
        if dynamic_values:
            df = df.rename(columns= dynamic_values)
        if static_values:
            for key, value in static_values.items():
                df[key] = value
        self.contexts = df.to_dict('records')
        