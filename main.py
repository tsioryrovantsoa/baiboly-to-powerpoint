import tkinter as tk
from tkinter import ttk
from tkinter import *
from presentation import *
from tkinter.ttk import *
from PIL import ImageTk, Image
from tkinter import messagebox

root = tk.Tk();
root.geometry('350x250')
root.title("Baiboly to PowerPoint")
root.iconbitmap(r'baiboly.ico')
# imagy = Image.open("fimpiz.PNG")
# resized_image= imagy.resize((80,75))
# img = ImageTk.PhotoImage(resized_image)
# imagyr = Image.open("fimpiz.PNG")
# resized_imager= imagyr.resize((80,75))
# imgr = ImageTk.PhotoImage(resized_imager)
labelChoix = tk.Label(root,text="FJKM Andavamamba Kristy Velona")
labelChoix.pack()
titre = tk.Label(root,text="BAIBOLY",font=('Averta-ExtraBoldItalic',25))
titre.pack()



# sary = Label(root,image=img)
# sary.place(x=0,y=10)
# saryr = Label(root,image=imgr)
# saryr.place(x=260,y=20)

root.resizable(False, False)

conn = create_connection()
with conn:

    rows = allBoky(conn)
    Boky = []
    idBoky = []
    for row in rows:

        Boky.append(row[1])
        idBoky.append(row[0])
    #     print(row[1])
    #     id = row[0]
    #     print(row[0])



toko = Label(root,text="Toko")
toko.place(x=100,y=120)
entre_1= Entry(root,width=10)
entre_1.place(x=140,y=120)
andininydeb = Label(root,text="Andininy faha")
andininydeb.place(x=100,y=150)
entre_2= Entry(root,width=10)
entre_2.place(x=180,y=150)
andininyfin = Label(root,text="hatramin'ny")
andininyfin.place(x=100,y=180)
entre_3= Entry(root,width=10)
entre_3.place(x=180,y=180)




def sokafy():
    select = listeBoky.get()
    print(select)
    ato = Boky.index(select)
    id = idBoky[ato]
    print(id)
    toko_tena = entre_1.get()
    andininydeb_tena = entre_2.get()
    andininyfin_tena = entre_3.get()
    # if(type(andininydeb_tena)&type(andininyfin_tena)!='int'):
    if(toko_tena=='' or andininydeb_tena=='' or andininyfin_tena==''):
        messagebox.showerror("Baiboly to PowerPoint", "Hamarino tsara ny zavatra ampidirinao.\nMisaotra!")
    else:
        print('faranyy-----')
        print(type(toko))
        print(type(andininyfin_tena))
        print(type(andininydeb_tena))
        print((id))
        print(type(toko_tena))
        print(type(andininydeb_tena))
        print(type(andininyfin_tena))
        fichier = creer(str(id),toko_tena,andininydeb_tena,andininyfin_tena)
        openFile(fichier)

style = Style()
button=Button(root, text="Sokafy",style="W.TButton",width=15,command=sokafy)
button.place(x=120,y=215)

# listeBoky.bind("<<ComboboxSelected>>",sokafy)
listeBoky = ttk.Combobox(root,state="readonly", values=Boky)
listeBoky.current(18)
listeBoky.place(x=100,y=90)
conn.close()






root.mainloop()