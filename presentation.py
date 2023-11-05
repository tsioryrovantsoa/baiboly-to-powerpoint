import os
from pptx import Presentation
from pptx.util import Pt
from function import *
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Inches
from pptx.enum.text import MSO_AUTO_SIZE

def creer(idBoky,toko,andininydeb,andininyfin):
    pres  = Presentation('template.pptx')
    slides = [slide for slide in pres.slides]
    slide0 = slides[0]
    title_shapes = [shape for shape in slide0.shapes if shape.has_text_frame]
    title = [shape for shape in title_shapes if shape.has_text_frame and shape.text == 'boky']

    conn = create_connection()
    with conn:
        title[0].text = getBoky(conn, idBoky,toko,andininydeb,andininyfin)
        title_para = slide0.shapes.title.text_frame.paragraphs[0]
        title_para2 = slide0.shapes.title.text_frame.paragraphs[1]
        title_para.font.name = "Alégre Sans NC"
        if (title_para.text .__eq__ ("Tonon-kiran'i Solomona ")):
            title_para.font.size = Pt(96)
        elif(title_para.text =="Asan'ny Apostoly "):
            title_para.font.size = Pt(134)
        else:
            title_para.font.size = Pt(161)
        title_para2.font.name = "Alégre Sans NC"
        title_para2.font.size = Pt(161)
        rows = getVerser(conn, idBoky,toko,andininydeb,andininyfin)
        for row in rows:
            verset = row[1]
            if(len(verset.split())>13):
                ponctuation = ','

                zaraina = verset.count(ponctuation)
                if(zaraina==0):
                    ponctuation =';'
                    zaraina = verset.count(ponctuation)
                    if(zaraina==0):
                        ponctuation='!'
                        zaraina = verset.count(ponctuation)
                        if (zaraina == 0):
                            ponctuation = '?'
                            zaraina = verset.count(ponctuation)
                versetoff = verset.split(ponctuation,zaraina)
                for x in versetoff:
                    if(x==''):
                        continue
                    first_slide_layout = pres.slide_layouts[5]
                    slide = pres.slides.add_slide(first_slide_layout)
                    if(x == versetoff[0]):
                        slide.shapes.title.text = "" + str(row[0]) + " " + str(x)
                    else:
                        slide.shapes.title.text = "" + str(x)
                    slide.shapes.title.width = Inches(10)
                    slide.shapes.title.height = Inches(2.28)
                    slide.shapes.title.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
                    slide.shapes.title.text_frame.margin_bottom = Inches(-5.1)
                    slide.shapes.title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                    slide.shapes.title.text_frame.paragraphs[0].font.name = "Helvetica Inserat LT Std"
                    if (12 == len(x.split())):
                        slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(72)
                    else:
                        slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(80)
            else:
                first_slide_layout = pres.slide_layouts[5]
                slide = pres.slides.add_slide(first_slide_layout)
                slide.shapes.title.text = ""+str(row[0])+" "+row[1]
                slide.shapes.title.width= Inches(10)
                slide.shapes.title.height = Inches(2.28)
                slide.shapes.title.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
                slide.shapes.title.text_frame.margin_bottom = Inches(-5.1)
                slide.shapes.title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                slide.shapes.title.text_frame.paragraphs[0].font.name = "Helvetica Inserat LT Std"
                versetoff = verset.split(", ", 2)
                for x in versetoff:
                    if (12 == len(x.split())):
                        slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(72)
                    else:
                        slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(80)

    fichier = 'test.pptx'
    pres.save(fichier)
    return fichier

def openFile(pres):
    fileName = pres
    os.system("start " + fileName)