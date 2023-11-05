import os
from pptx import Presentation
from pptx.util import Pt
from function import *
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Inches
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.dml.color import RGBColor


def creer(idBoky, toko, andininydeb, andininyfin):
    pres = Presentation("template.pptx")
    slide0 = pres.slides[0]
    title_shapes = [shape for shape in slide0.shapes if shape.has_text_frame]
    title = [
        shape for shape in title_shapes if shape.has_text_frame and shape.text == "boky"
    ]

    conn = create_connection()
    with conn:
        title[0].text = getBoky(conn, idBoky, toko, andininydeb, andininyfin)
        format_title(slide0)
        rows = getVerser(conn, idBoky, toko, andininydeb, andininyfin)

    for row in rows:
        create_verset_slides(row, pres)

    fichier = "test.pptx"
    pres.save(fichier)
    return fichier


def format_title(slide0):
    title_para = slide0.shapes.title.text_frame.paragraphs[0]
    title_para2 = slide0.shapes.title.text_frame.paragraphs[1]
    title_para.font.name = "Alégre Sans NC"
    size = (
        Pt(100)
        if title_para.text == "Tonon-kiran'i Solomona "
        else Pt(134)
        if title_para.text == "Asan'ny Apostoly "
        else Pt(170)
        if title_para.text == "Deotoronomia "
        else Pt(175)
        if title_para.text == "I Tesaloniana "
        else Pt(175)
        if title_para.text == "II Tesaloniana "
        else Pt(185)
    )
    title_para.font.size = size
    title_para2.font.name = "Alégre Sans NC"
    title_para2.font.size = Pt(185)
    title_para.line_spacing = Pt(12)
    title_para.font.color.rgb = RGBColor(255, 255, 0)
    title_para2.font.color.rgb = RGBColor(255, 255, 0)


def create_verset_slides(row, pres):
    verset = row[1]
    first_slide_layout = pres.slide_layouts[5]

    if len(verset.split()) > 13:
        ponctuations = [",", ";", "!", "?"]
        zaraina = 0

        for ponctuation in ponctuations:
            if ponctuation in verset:
                zaraina = verset.count(ponctuation)
                break

        versetoff = verset.split(ponctuation, zaraina)
        base_text = str(row[0])

        for i, x in enumerate(versetoff):
            if x:
                slide = pres.slides.add_slide(first_slide_layout)
                text = f"{base_text} {x}" if i == 0 else x
                set_slide_properties(slide, text, size=Pont_size(x))
    else:
        slide = pres.slides.add_slide(first_slide_layout)
        text = f"{row[0]} {row[1]}"
        set_slide_properties(slide, text, size=Pont_size(row[1]))

def set_slide_properties(slide, text, size):
    slide.shapes.title.text = text
    slide.shapes.title.width = Inches(10)
    slide.shapes.title.height = Inches(2.28)
    slide.shapes.title.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    slide.shapes.title.text_frame.margin_bottom = Inches(-5.1)
    slide.shapes.title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    slide.shapes.title.text_frame.paragraphs[0].font.name = "Helvetica Inserat LT Std"
    slide.shapes.title.text_frame.paragraphs[0].font.size = size

def Pont_size(text):
    return Pt(72) if len(text.split()) == 12 else Pt(80)

def openFile(pres):
    fileName = pres
    os.system("start " + fileName)
