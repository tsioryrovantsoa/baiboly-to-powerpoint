import sqlite3
from sqlite3 import Error
from tkinter import messagebox
import os

def create_connection():
    try:
        current_directory = os.path.dirname(os.path.abspath(__file__))
        conn = sqlite3.connect(os.path.join(current_directory, "bible_baiboly.sqlite"))
        return conn
    except Error as e:
        print(e)
        return None

def allBoky(conn):
    cur = conn.cursor()
    cur.execute("SELECT _id, title FROM chapters ORDER BY title")
    return cur.fetchall()

def getVerser(conn, idBoky, toko, andininydeb, andininyfin):
    cur = conn.cursor()
    cur.execute("SELECT position, text FROM texts WHERE chapter_id=? AND chapter_num=? AND position BETWEEN ? AND ?", (idBoky, toko, andininydeb, andininyfin))
    rows = cur.fetchall()
    if not rows:
        raise Exception("Hamarino tsara ny zavatra ampidirinao.\nMisaotra!")
    return rows

def getBoky(conn, idBoky, toko, andininydeb, andininyfin):
    cur = conn.cursor()
    cur.execute("SELECT title FROM chapters WHERE _id=?", (idBoky,))
    rows = cur.fetchone()
    tokohafa = rows[0]
    if '1' in tokohafa:
        tokohafa = tokohafa.replace('1', 'I')
    elif '2' in tokohafa:
        tokohafa = tokohafa.replace('2', 'II')
    elif '3' in tokohafa:
        tokohafa = tokohafa.replace('3', 'III')
    if andininydeb == andininyfin:
        valiny = f"{tokohafa} \n{toko} : {andininydeb}"
    else:
        valiny = f"{tokohafa} \n{toko} : {andininydeb} - {andininyfin}"
    return valiny

# def main():
#     conn = create_connection()
#     with conn:
#         # rows = getVerser(conn, 8265, 15, 10, 10)
#         rows = getBoky(conn, 8265, 15, 10, 10)
#         print(rows)
#         # conn.close()
#         # allBoky(conn)
#
# if __name__ == '__main__':
#     main()