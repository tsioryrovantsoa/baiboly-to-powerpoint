import tkinter as tk
from tkinter import ttk
from tkinter import *
from presentation import *
from tkinter.ttk import *
from PIL import ImageTk, Image
from tkinter import messagebox
from function import allBoky


def on_key_release(event):
    """
    Fonction pour capturer les frappes clavier lors de la sélection des options dans la Combobox
    et sélectionner automatiquement l'élément correspondant dans la Combobox.
    """
    current_value = event.char.lower()
    if current_value.isalpha() or current_value in ['1', '2', '3']:
        current_index = listeBoky.current()
        next_index = (
            current_index + 1 if current_index < len(listeBoky["values"]) - 1 else 0
        )
        for i, value in enumerate(
            listeBoky["values"][next_index:] + listeBoky["values"][:next_index]
        ):
            if value.lower().startswith(current_value):
                listeBoky.current((i + next_index) % len(listeBoky["values"]))
                break

def main():
    root = tk.Tk()
    root.geometry("350x250")
    root.title("Baiboly to PowerPoint")
    root.iconbitmap(r"baiboly.ico")

    labelChoix = tk.Label(root, text="FJKM Andavamamba Kristy Velona")
    labelChoix.pack()
    titre = tk.Label(root, text="BAIBOLY", font=("Averta-ExtraBoldItalic", 25))
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
        try:
            select = listeBoky.get()
            ato = Boky.index(select)
            id = idBoky[ato]
            toko_tena = entre_1.get()
            andininydeb_tena = entre_2.get()
            andininyfin_tena = entre_3.get()
            if toko_tena == "" or andininydeb_tena == "" or andininyfin_tena == "":
                raise Exception("Hamarino tsara ny zavatra ampidirinao.\nMisaotra!")
            fichier = creer(str(id), toko_tena, andininydeb_tena, andininyfin_tena)
            openFile(fichier)
        except Exception as e:
            messagebox.showerror("Baiboly to PowerPoint", str(e))

    button = Button(root, text="Sokafy", style="W.TButton", width=15, command=sokafy)
    button.place(x=120, y=215)

    global listeBoky
    listeBoky = ttk.Combobox(root, state="readonly", values=Boky)
    listeBoky.current(61)
    listeBoky.place(x=100, y=90)
    listeBoky.bind("<KeyRelease>", on_key_release)
    listeBoky.focus_set()
    conn.close()
    
    version_label = tk.Label(root, text="v1.1-nov2023") 
    version_label.place(relx=1, rely=1, anchor='se')

    root.mainloop()


if __name__ == "__main__":
    main()
