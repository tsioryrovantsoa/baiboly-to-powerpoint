import tkinter as tk
from tkinter import ttk
from tkinter import *
from presentation import *
from tkinter.ttk import *
from PIL import ImageTk, Image
from tkinter import messagebox

def main():
    root = tk.Tk()
    root.geometry('350x250')
    root.title("Baiboly to PowerPoint")
    root.iconbitmap(r'baiboly.ico')

    labelChoix = tk.Label(root, text="FJKM Andavamamba Kristy Velona")
    labelChoix.pack()
    titre = tk.Label(root, text="BAIBOLY", font=('Averta-ExtraBoldItalic', 25))
    titre.pack()
    root.resizable(False, False)

    conn = create_connection()
    with conn:
        rows = allBoky(conn)
        Boky = [row[1] for row in rows]
        idBoky = [row[0] for row in rows]

    toko = Label(root, text="Toko")
    toko.place(x=100, y=120)
    entre_1 = Entry(root, width=10)
    entre_1.place(x=140, y=120)
    andininydeb = Label(root, text="Andininy faha")
    andininydeb.place(x=100, y=150)
    entre_2 = Entry(root, width=10)
    entre_2.place(x=180, y=150)
    andininyfin = Label(root, text="hatramin'ny")
    andininyfin.place(x=100, y=180)
    entre_3 = Entry(root, width=10)
    entre_3.place(x=180, y=180)

    def sokafy():
        select = listeBoky.get()
        ato = Boky.index(select)
        id = idBoky[ato]
        toko_tena = entre_1.get()
        andininydeb_tena = entre_2.get()
        andininyfin_tena = entre_3.get()
        if toko_tena == '' or andininydeb_tena == '' or andininyfin_tena == '':
            messagebox.showerror("Baiboly to PowerPoint", "Hamarino tsara ny zavatra ampidirinao.\nMisaotra!")
        else:
            fichier = creer(str(id), toko_tena, andininydeb_tena, andininyfin_tena)
            openFile(fichier)

    button = Button(root, text="Sokafy", style="W.TButton", width=15, command=sokafy)
    button.place(x=120, y=215)

    listeBoky = ttk.Combobox(root, state="readonly", values=Boky)
    listeBoky.current(18)
    listeBoky.place(x=100, y=90)
    conn.close()

    root.mainloop()

if __name__ == '__main__':
    main()
