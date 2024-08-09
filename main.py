from pptx import Presentation
from dotenv import load_dotenv
import os

load_dotenv()

path = "C:/Prueba"

def copy_slide(new_pptx, ex_slide):
    slide_layout = ex_slide.slide_layout
    new_slide = new_pptx.slides.add_slide(slide_layout)
    for shape in ex_slide.shapes:
        if shape.shape_type == 14:  # Text Boxes
                new_shape = new_slide.shapes.add_textbox(
                    shape.left, shape.top, shape.width, shape.height
                )
                if shape.has_text_frame:
                    new_shape.text = shape.text_frame.text

        else:
            new_shape = new_slide.shapes.add_shape(shape.auto_shape_type, shape.left, shape.top, shape.width, shape.height)
            if shape.has_text_frame:
                new_shape.text = shape.text_frame.text

def copy_pptx():
    new_pptx = Presentation()
    for root, _, files in os.walk(path):
        for file in files:
            act_pptx = Presentation(f'{root}/{file}')
            for slide in act_pptx.slides:
                copy_slide(new_pptx, slide)

    new_pptx.save('output/output.pptx')

copy_pptx()