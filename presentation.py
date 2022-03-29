import os
from pptx import Presentation
from pptx.util import Pt
from function import *
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Inches
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.shapes.base import BaseShape

def creer(idBoky,toko,andininydeb,andininyfin):
    current_dir = os.path.dirname(os.path.realpath(__file__))

    pres  = Presentation('template.pptx')
    slides = [slide for slide in pres.slides]
    slide0 = slides[0]
    slide0_shapes = [shape for shape in slide0.shapes]
    title_shapes = [shape for shape in slide0.shapes if shape.has_text_frame]
    placeholder_text = [shape.text for shape in title_shapes if shape.has_text_frame]
    title = [shape for shape in title_shapes if shape.has_text_frame and shape.text == 'boky']


    # slide1 = slides[1]
    # slide1_shapes = [shape for shape in slide1.shapes]
    # title1_shapes = [shape for shape in slide1.shapes if shape.has_text_frame]
    # placeholder1_text = [shape.text for shape in title1_shapes if shape.has_text_frame]
    # title1 = [shape for shape in title1_shapes]
    # title1[0]. text = "Eto ilay verse"
    # new_placeholder_text1 = [shape.text for shape in title1_shapes if shape.has_text_frame]
    #
    # title_para1 = slide1.shapes.title.text_frame.paragraphs[0]
    # title_para1.font.name = "Helvetica Inserat LT Std"
    # title_para1.font.size = Pt(80)
    conn = create_connection()
    with conn:
        title[0].text = getBoky(conn, idBoky,toko,andininydeb,andininyfin)
        new_placeholder_text = [shape.text for shape in title_shapes if shape.has_text_frame]
        print('ity ilay titre eh'+title[0].text)
        title_para = slide0.shapes.title.text_frame.paragraphs[0]
        print('ity ilay titre 1'+str(title_para.text))
        title_para2 = slide0.shapes.title.text_frame.paragraphs[1]
        print('ity ilay titre 2' + str(title_para))
        title_para.font.name = "Alégre Sans NC"
        print(type(title_para.text))
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
            # delimiter = ','
            # verset = delimiter.join([str(value) for value in row])
            # slide.shapes.title.text = ""+str(row[0])+" "+row[1]
            verset = row[1]
            print(len(row[1].split()))
            if(len(verset.split())>13):
                print('-----------------etooo')
                ponctuation = ','
                print(verset.count(','))

                zaraina = verset.count(ponctuation)
                if(zaraina==0):
                    ponctuation =';'
                    zaraina = verset.count(ponctuation)
                    print('tafitra atooo')
                    print('ity ny isany'+str(zaraina))
                    if(zaraina==0):
                        ponctuation='!'
                        zaraina = verset.count(ponctuation)
                        if (zaraina == 0):
                            ponctuation = '?'
                            zaraina = verset.count(ponctuation)
                print("ity : ...." + str(verset.split(ponctuation)))
                # versetoff = [e + ponctuation for e in verset.split(ponctuation, zaraina) if e] (version misy virgule)
                versetoff = verset.split(ponctuation,zaraina)
                for x in versetoff:
                    if(x==''):
                        continue
                    first_slide_layout = pres.slide_layouts[5]
                    slide = pres.slides.add_slide(first_slide_layout)
                    # versetspliter1= verset.split(", ")[0:13]
                    print("ito ilay 0:" +versetoff[0])
                    if(x == versetoff[0]):
                        slide.shapes.title.text = "" + str(row[0]) + " " + str(x)
                    else:
                        slide.shapes.title.text = "" + str(x)
                    # first1_slide_layout = pres.slide_layouts[5]
                    # slide1 = pres.slides.add_slide(first1_slide_layout)
                    # slide1.shapes.title.text = "" + str(verset.split(", ")[13::])
                    print("boucle " +str(x))
                    slide.shapes.title.width = Inches(10)
                    slide.shapes.title.height = Inches(2.28)
                    slide.shapes.title.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
                    slide.shapes.title.text_frame.margin_bottom = Inches(-5.1)
                    slide.shapes.title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                    slide.shapes.title.text_frame.paragraphs[0].font.name = "Helvetica Inserat LT Std"
                    print("ity ilay isany : ... " +str(len(x.split())))
                    if (12 == len(x.split())):
                        slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(72)
                    else:
                        slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(80)
                    # slide1.shapes.title.width = Inches(10)
                    # slide1.shapes.title.height = Inches(2.28)
                    # slide1.shapes.title.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
                    # slide1.shapes.title.text_frame.margin_bottom = Inches(-5.1)
                    # slide1.shapes.title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                    # slide1.shapes.title.text_frame.paragraphs[0].font.name = "Helvetica Inserat LT Std"
                    # slide1.shapes.title.text_frame.paragraphs[0].font.size = Pt(80)
                    # slide = pres.slides.add_slide(first_slide_layout)
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
                    # placeholder = slide.placeholders[0]
                    # assert placeholder.type == PP_PLACEHOLDER.CENTER_TITLE
                    print(row)
                    # conn.close()


    print(len(slide0_shapes))
    print(len(slides))
    print(placeholder_text)
    print(new_placeholder_text)

    fichier = 'test.pptx'
    pres.save(fichier)
    return fichier

def openFile(pres):
    fileName = pres
    os.system("start " + fileName)