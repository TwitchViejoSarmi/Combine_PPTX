from pptx import Presentation
from dotenv import load_dotenv
import os
from io import BytesIO

load_dotenv()

path = "Prueba"

def copy_slide(new_pptx, ex_slide):
    new_slide = new_pptx.slides.add_slide(ex_slide.slide_layout)
    for shape in ex_slide.shapes:
        print(f"{shape.shape_type}")
        if shape.has_text_frame:
            print(shape.text)
        if shape.is_placeholder:
            phf = new_slide.placeholders[shape.placeholder_format.idx]
            if shape.has_text_frame:
                phf.text = shape.text_frame.text
            elif shape.has_chart:
                chart = shape.chart
                phf.insert_chart(
                    chart.chart_type,
                    chart_data=chart.chart_data
                )
            elif shape.has_table:
                table = shape.table
                for i, row in enumerate(table.rows):
                    for j, cell in enumerate(row.cells):
                        phf.table.cell(i, j).text = cell.text
            elif shape.has_picture:
                image = shape.image
                phf.insert_picture(image)
        elif shape.shape_type == 17:  # 17 es el tipo de shape para cuadros de texto
            new_shape = new_slide.shapes.add_textbox(
                shape.left, shape.top, shape.width, shape.height
            )
            new_shape.text_frame.text = shape.text_frame.text
        elif shape.shape_type == 13:  # 13 es el tipo de shape para imágenes
            image = shape.image
            new_slide.shapes.add_picture(BytesIO(image.blob), shape.left, shape.top, shape.width, shape.height)
        else:
            if shape.has_chart:
                chart = shape.chart
                new_chart = new_slide.shapes.add_chart(
                    chart.chart_type, 
                    shape.left, shape.top, shape.width, shape.height, 
                    chart_data=chart.chart_data
                ).chart
                
                new_chart.has_legend = chart.has_legend
                new_chart.legend.position = chart.legend.position

            elif shape.has_table:
                table = shape.table
                new_table = new_slide.shapes.add_table(
                    rows=len(table.rows),
                    cols=len(table.columns),
                    left=shape.left,
                    top=shape.top,
                    width=shape.width,
                    height=shape.height
                ).table

                for i, row in enumerate(table.rows):
                    for j, cell in enumerate(row.cells):
                        new_table.cell(i, j).text = cell.text

            else:
                new_shape = new_slide.shapes.add_shape(
                    shape.shape_type,
                    shape.left, shape.top, shape.width, shape.height
                )
        
            if shape.has_text_frame:
                new_shape.text = shape.text_frame.text

    return new_slide

def copy_pptx():
    new_pptx = Presentation()
    for root, _, files in os.walk(path):
        for file in files:
            act_pptx = Presentation(f'{root}/{file}')
            for slide in act_pptx.slides:
                copy_slide(new_pptx, slide)

    try:
        os.makedirs('output/')
    except FileExistsError:
        pass
    
    new_pptx.save('output/output.pptx')
copy_pptx()