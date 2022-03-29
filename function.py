import sqlite3
from sqlite3 import Error


import sqlite3
from sqlite3 import Error
from tkinter import messagebox

def create_connection():
    """ create a database connection to the SQLite database
        specified by the db_file
    :param db_file: database file
    :return: Connection object or None
    """
    conn = None
    try:
        conn = sqlite3.connect(r"bible_baiboly.sqlite")
    except Error as e:
        print(e)

    return conn

def allBoky(conn):
    """
    Query all rows in the tasks table
    :param conn: the Connection object
    :return:
    """
    cur = conn.cursor()
    cur.execute("select _id,title from chapters")

    rows = cur.fetchall()
    # conn.close()
    return rows


def getVerser(conn, idBoky,toko,andininydeb,andininyfin):
    """
    Query tasks by priority
    :param conn: the Connection object
    :param priority:
    :return:
    """
    cur = conn.cursor()
    cur.execute("select position,text from texts where chapter_id=? and chapter_num=? and  position between ? and ?", (idBoky,toko,andininydeb,andininyfin))

    rows = cur.fetchall()
    print('-----------****************--------------')
    print('ity : '+str(len(rows)))
    if(len(rows) == 0):
        messagebox.showerror("Baiboly to PowerPoint", "Hamarino tsara ny zavatra ampidirinao.\nMisaotra!")
    else:
    #conn.close()
        return rows

def getBoky(conn,idBoky,toko,andininydeb,andininyfin):
    cur = conn.cursor()
    cur.execute("select title from chapters where _id=?",(idBoky,))
    print(idBoky)
    rows = cur.fetchone()
    print('eto hoee...')
    print(rows[0])
    print(toko)
    print(andininyfin)
    print(andininydeb)
    tokohafa = rows[0]
    if '1' in tokohafa:
        tokohafa = rows[0].replace('1','I')
        if (andininydeb == andininyfin):
            valiny = "" + str(tokohafa) + " \n" + str(toko) + " : " + str(andininydeb)
        else:
            valiny = "" + str(tokohafa) + " \n" + str(toko) + " : " + str(andininydeb) + " - " + str(andininyfin)
    elif '2' in tokohafa:
        tokohafa = rows[0].replace('2','II')
        if (andininydeb == andininyfin):
            valiny = "" + str(tokohafa) + " \n" + str(toko) + " : " + str(andininydeb)
        else:
            valiny = "" + str(tokohafa) + " \n" + str(toko) + " : " + str(andininydeb) + " - " + str(andininyfin)
    if '3' in tokohafa:
        tokohafa = rows[0].replace('3','III')
        if (andininydeb == andininyfin):
            valiny = "" + str(tokohafa) + " \n" + str(toko) + " : " + str(andininydeb)
        else:
            valiny = "" + str(tokohafa) + " \n" + str(toko) + " : " + str(andininydeb) + " - " + str(andininyfin)
    else:
        if(andininydeb==andininyfin):
            valiny = "" + str(tokohafa) + " \n" + str(toko) + " : " + str(andininydeb)
        else:
            valiny = "" +str(tokohafa)+ " \n" +str(toko)+ " : " +str(andininydeb)+ " - " +str(andininyfin)
    print(valiny)
    # conn.close()
    return valiny

# def main():
#
#     # create a database connection
#     conn = create_connection()
#     with conn:
#
#         # rows = getVerser(conn, 8265,15,10,10)
#         rows = getBoky(conn, 8265,15,10,10)
#         print(rows)
#         # conn.close()
#         # allBoky(conn)
#
#
# if __name__ == '__main__':
#     main()