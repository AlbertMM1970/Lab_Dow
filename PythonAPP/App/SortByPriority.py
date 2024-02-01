
import datetime
import stat
from tkinter import filedialog
import tkinter as tk
# import customtkinter as cTk
from tkinter import *
import tkinter.ttk as ttk
import os.path
from tkinter import messagebox
import xlsxwriter
import xlrd
import copy
import shutil
from functools import partial
from tkinter.simpledialog import askfloat
from tkinter.simpledialog import askinteger
from tkinter.messagebox import showinfo
from decimal import *
import tkinter.font as font
from collections import defaultdict
import webbrowser
from difflib import SequenceMatcher
import clipboard as copia
import sqlite3
import sys
import traceback
import os


# BackUP = []
# data = ""
mat = []
mats = []
SamplesN = []
a = 0
Path_image = r'C:\Python\Images\Dow_Chemical_logo.png'
url = 'https://irssrs.intranet.dow.com/Reports/report/LIMS_Reports/Performance%20Plastics/StudySummary_BC'
tagn = False
SampleN = False
wordToFind = ""
SortResin = ""
Layer = ""
# rep = ""
ListSort = []
Sorting = []
f = []
j = ""
l = 0
h = []
matM = []
matYesNo = []
id2 = ""
Base = False
resin = False
input1 = False
input2 = False
LogOpen = False
text2 = ""


class Sql_Data():

    def Connection():
        try:
            sqliteConnection = sqlite3.connect('SQLite_Lab.db')
            cursor = sqliteConnection.cursor()
            print("Database created and Successfully Connected to SQLite")

            sqlite_select_Query = "select sqlite_version();"
            cursor.execute(sqlite_select_Query)
            record = cursor.fetchall()
            print("SQLite Database Version is: ", record)
            cursor.close()

        except sqlite3.Error as error:
            print("Error while connecting to sqlite", error)
        finally:
            if sqliteConnection:
                sqliteConnection.close()
                print("The SQLite connection is closed")

    def Insert_data(name, date, mfi, density):

        try:
            sqliteConnection = sqlite3.connect('SQLite_Lab.db')
            cursor = sqliteConnection.cursor()
            print("Successfully Connected to SQLite")

            sqlite_insert_query = f"""INSERT INTO db_resin(
                                  id,                                  
                                  date,
                                  resin, 
                                  mfi, 
                                  density) 
                                  VALUES (?,{name},{date},{mfi},{density})"""

            count = cursor.execute(sqlite_insert_query)
            sqliteConnection.commit()
            print("Record inserted successfully into db_resin table ",
                  cursor.rowcount)
            cursor.close()

        except sqlite3.Error as error:
            print("Failed to insert data into sqlite table", error)
        finally:
            if sqliteConnection:
                sqliteConnection.close()
                print("The SQLite connection is closed")

    def Create_db():

        try:
            sqliteConnection = sqlite3.connect('SQLite_Lab.db')
            sqlite_create_table_query = '''CREATE TABLE IF NOT EXISTS db_resin (
                                id INTEGER PRIMARY KEY,                                
                                date datetime NOT NULL, 
                                resin VARCHAR(100) NOT NULL UNIQUE,                               
                                mfi DECIMAL NOT NULL,
                                density DECIMAL NOT NULL);'''

            cursor = sqliteConnection.cursor()
            print("Successfully Connected to SQLite")
            cn = cursor.execute(sqlite_create_table_query)
            sqliteConnection.commit()
            if cn == True:
                print("SQLite table created")
            else:
                print("SQLite table already exists")

            cursor.close()

        except sqlite3.Error as error:
            print("Error while creating a sqlite table", error)
        finally:
            if sqliteConnection:
                sqliteConnection.close()
                print("sqlite connection is closed")

#######  -------- GUI - Application SortByPriority -------- ######


class Window:

    def __init__(self, master):

        # Main screen configuration
        self.master = master
        self.master.resizable(1, 1)
        self.master.title("Sort By Priority by amanonellas@dow.com")
        # Obtenemos el largo y  ancho de la pantalla
        wtotal = self.master.winfo_screenwidth()
        htotal = self.master.winfo_screenheight()
        # Guardamos el largo y alto de la ventana
        wventana = 1000
        hventana = 600
        # Aplicamos la siguiente formula para calcular donde debería posicionarse
        pwidth = round(wtotal/2-wventana/2)
        pheight = round(htotal/2-hventana/2)
        # Se lo aplicamos a la geometría de la ventana
        self.master.geometry(str(wventana)+"x"+str(hventana) +
                             "+"+str(pwidth)+"+"+str(pheight))
        self.master.columnconfigure(0, weight=1)  # column with treeview
        self.master.rowconfigure(4, weight=1)  # row with treeview
        self.Label_Engineer = tk.StringVar()

        # Button Choose file
        self.btn = tk.Button(
            text="Choose File ...", command=self.Choose_File)
        self.btn.grid(row=0, column=0, padx=5, pady=10,
                      ipady=0, ipadx=5, columnspan=1, sticky="NSE")

        # Button Open file
        self.btn1 = tk.Button(text="Open File ...", command=self.open_file)
        self.btn1.grid(row=0, column=1, padx=5, pady=10,
                       ipady=0, ipadx=5, sticky="EW")

        # Button search Study
        self.Sts = tk.Button(text="Study", command=self.Open_web)
        self.Sts.grid(row=0, column=2, padx=5, pady=0,
                      ipady=0, ipadx=5, sticky="E")

        # Text & Label Number of Samples
        self.labl1 = tk.Label(text="Number of Samples :").grid(
            row=0, column=3, ipady=5)
        self.entry3 = tk.Entry(justify='center', width=4, font="Verdana 10")
        self.entry3.grid(row=0, column=4, padx=0, pady=0,
                         ipady=3, ipadx=0, sticky="w")

        # Combo Samples
        self.SamplesCO = ttk.Combobox(
            justify='center', state='readonly', width=15, font="Verdana 10")
        self.SamplesCO.grid(row=0, column=5, padx=0, pady=0,
                            ipady=3, ipadx=0, columnspan=3, sticky="w")

        # Button go Sample
        self.btnGo = tk.Button(text="Go...", command=self.createModal)
        self.btnGo.grid(row=0, column=7, padx=50, pady=0,
                        ipady=0, ipadx=0, sticky="W")

        # Labels Owner of the Study
        self.LOwner = tk.Label(text="Engineer :").grid(
            row=0, column=8)
        self.LOwnerN = tk.Label(textvariable=self.Label_Engineer, font=(
            "Verdana bold", 14)).grid(
            row=1, column=6, columnspan=5, padx=0, pady=0, ipady=0, ipadx=60)

        # Button Console
        self.btnConsola = tk.Button(text="Console", command=self.Consola)
        self.btnConsola.grid(row=0, column=10, columnspan=10,
                             padx=0, pady=0, ipady=0, ipadx=30, sticky="W")

        # Text path of the file
        self.entry = tk.Entry(width=65, font="Verdana 10")
        self.entry.grid(row=1, column=0, columnspan=7,
                        padx=10, pady=0, ipady=5, ipadx=0, sticky="w")

        # Button Export Excel file
        self.btn3 = tk.Button(text="Export to EXCEL file",
                              command=self.SaveToExcel)
        self.btn3.grid(row=2, column=0, columnspan=2, padx=10,
                       pady=5, ipady=0, ipadx=0, sticky="w")

        # Button MDO tools
        self.btn4 = tk.Button(text="MDO tools",
                              command=self.mdotools)
        self.btn4.grid(row=2, column=1, columnspan=3, padx=40,
                       pady=5, ipady=0, ipadx=0, sticky="w")
        # Button Datasheet
        self.btn4 = tk.Button(text="DataSheet",
                              command=self.open_DataSheet)
        self.btn4.grid(row=2, column=2, columnspan=5, padx=20,
                       pady=5, ipady=0, ipadx=0, sticky="w")
        # Button RawData
        self.btn5 = tk.Button(text="RawData tools",
                              command=self.rawData)
        self.btn5.grid(row=2, column=3, columnspan=6, padx=30,
                       pady=5, ipady=0, ipadx=0, sticky="w")

        # Text to entry EXCEL name
        self.entry1 = tk.Entry(font="Verdana 10", width=15)
        self.entry1.grid(row=3, column=0, columnspan=2, padx=10,
                         pady=0, ipady=5, ipadx=0, sticky="w")

        # Button Sort by Density
        self.btn_Density = tk.Button(text="Sort by Density",
                                     command=partial(self.SortBy, "Density"))
        self.btn_Density.grid(row=3, column=1, columnspan=2, padx=10,
                              pady=0, ipady=0, ipadx=0, sticky="E")

        # Button Sort by Melt Index
        self.btn2 = tk.Button(text="Sort by Melt Index",
                              command=partial(self.SortBy, "Layer"))
        self.btn2.grid(row=3, column=3, sticky="w")

        # Combo Melt Index option
        self.EntryMelt = ttk.Combobox(
            justify='center', state='readonly', width=5)
        self.EntryMelt.grid(row=3, column=4, padx=0,
                            pady=0, ipady=3, ipadx=0, sticky="w")
        intervals1 = ["High", "Low"]
        self.EntryMelt['values'] = intervals1
        self.EntryMelt.set("High")

        # Label & Text Layer number
        self.labl = tk.Label(text="Layer :").grid(
            row=3, column=5, padx=5,
            pady=0, ipady=0, ipadx=4, sticky="w")
        self.InterEntry = ttk.Combobox(
            justify='center', state='readonly', width=3)
        self.InterEntry.grid(row=3, column=6,  padx=0,
                             pady=0, ipady=3, ipadx=0, sticky="w")
        intervals = ["1", "2", "3", "4", "5", "6", "7", "8", "9"]
        self.InterEntry['values'] = intervals
        self.InterEntry.set("1")

        # Label & Text Materials
        self.res = tk.Label(text="Resin & Additive :").grid(
            row=3, column=7,  padx=25,
            pady=5, ipady=0, ipadx=0, sticky="w")
        self.Materials = ttk.Combobox(justify='center', state='readonly')
        self.Materials.grid(row=3, column=8, padx=0,
                            pady=0, ipady=3, ipadx=40)
        self.btn4 = tk.Button(text="Sort by Resin ...",
                              command=partial(self.SortBy, "Resin"))
        self.btn4.grid(row=3, column=10, padx=3,
                       pady=0, ipady=0, ipadx=0)

        # Treeview Samples
        self.tree = ttk.Treeview(
            self.master, selectmode='browse')
        self.tree.grid(row=4, column=0, columnspan=13,
                       sticky="nsew", padx=15, pady=10, ipadx=0)

        self.CreateTable()

        self.firstOPen = False

    def rawData(self):

        today = datetime.date.today()
        year = today.year
        gipnpath = "//tgnt01/g-tg-plastic/Laboratory/LIMS/LIMS Upload"

        msg = tk.simpledialog.askinteger(
            "RAW DATA TOOLS", "-               Please, enter the STUDY NUMBER                   -")
        if msg == "" or msg == None:
            return

        rawdatapath_main = f"//tnnt02/ptc2/Nautilus/GLIMS_Raw_Data/{str(year)}"

        rawdatapath = "//tnnt02/ptc2/Nautilus/GLIMS_Raw_Data/" + \
            str(year) + "/Study_" + str(msg) + "/Tarragona TSD Fabrication Lab"

        if os.path.isdir(rawdatapath):
            with os.scandir(rawdatapath) as ficheros:
                ficheros = [
                    fichero.name for fichero in ficheros if fichero.is_file()]
            if len(ficheros) > 0:
                msgb = tk.messagebox.askquestion(
                    "Study found", f"This study {msg} is on the RawData.\nAnd there are {len(ficheros)} file/s \nDo you want to pass all of them to GIPN?")
                if msgb == "yes":

                    try:
                        for fl in ficheros:
                            shutil.copy2(rawdatapath + "/" + str(fl), gipnpath)
                        tk.messagebox.showinfo(
                            "Passing files", "File/s passed correctly")
                    except:
                        tk.messagebox.showerror(
                            "Passing files", "File/s not passed correctly", icon="warning")
                else:
                    msgb = tk.messagebox.askquestion(
                        "Study found", f"Do you want to select the files to pass to GIPN?", icon="info")
                if msgb == "yes":
                    newtk = tk.Toplevel()  # 416673
                    newtk.resizable(0, 0)
                    newtk.title("Files on the Raw Data Study Folder")
                    wtotal = newtk.winfo_screenwidth()
                    htotal = newtk.winfo_screenheight()
                    # Guardamos el largo y alto de la ventana
                    wventana = 800
                    hventana = 500
                    # Aplicamos la siguiente formula para calcular donde debería posicionarse
                    pwidth = round(wtotal/2-wventana/2)
                    pheight = round(htotal/2-hventana/2)
                    # Se lo aplicamos a la geometría de la ventana
                    newtk.geometry(str(wventana)+"x"+str(hventana) +
                                   "+"+str(pwidth)+"+"+str(pheight))
                    newtk.focus()
                    newtk.grab_set()
                    variables = []
                    c = 0
                    self.data = ""
                    Label(newtk, text=f"Files of the STUDY {msg} ", font=("Verdana bold", 16)).place(
                        x=230, y=10)
                    Button(newtk, text=f"Copy to GIPN", font=("Verdana bold", 12), command=partial(self.save_gipn, newtk, ficheros, variables, rawdatapath, gipnpath)).place(
                        x=330, y=450)

                    for fl in ficheros:
                        variables.append(tk.IntVar(value=0))
                        tk.Checkbutton(newtk, text=str(
                            fl), variable=variables[-1], onvalue=1, offvalue=0, font="Verdana 12").place(x=40, y=60+c)
                        c = c + 22
                        # If row is out of range
                else:
                    msgb = tk.messagebox.askquestion(
                        "Study found", f"Do you want to open the RaWData folder?", icon="info")
                    if msgb == "yes":
                        webbrowser.open(str("file:") + rawdatapath)

        else:
            tk.messagebox.showerror(
                "Checking Study", f"This Study - {msg} - isn't in the RawData folder", icon="warning")
            msgb = tk.messagebox.askquestion(
                "Study not found", f"Do you want to open the RawData folder?", icon="info")
            if msgb == "yes":
                webbrowser.open(str("file:") + rawdatapath_main)

    def save_gipn(self, wn, data, data1, rawdatapath, gipnpath):

        result = [ing for ing, cb in zip(data, data1) if cb.get() > 0]
        try:
            for fl in result:
                shutil.copy2(rawdatapath + "/" + str(fl), gipnpath)
            tk.messagebox.showinfo(
                "Passing files", "File/s passed correctly", parent=wn)
            wn.destroy()
        except:
            tk.messagebox.showerror(
                "Passing files", "File/s not passed correctly", icon="warning", parent=wn)
            wn.destroy()

    def Open_web(self):

        pathfile = r'C:/Python/Studies/'
        pathDown = f'C:/Users/{os.getlogin()}/Downloads/'
        pathDownFile = f'C:/Users/{os.getlogin()}/Downloads/StudySummary_BC.xlsx'
        pathDownFile1 = f'C:/Users/{os.getlogin()}/Downloads/StudySummary_BC .xlsx'

        if os.path.isfile(pathDownFile):
            os.remove(pathDownFile)
        if os.path.isfile(pathDownFile1):
            os.remove(pathDownFile1)

        # print(pathDown)
        self.studyN = askinteger(
            "Search Study", "-                    Enter Study Number                    -")

        if self.studyN == "" or self.studyN == None:
            return

        ex = False
        check_file = os.path.isfile(
            pathfile + '/' + str(self.studyN) + '.xlsx')

        if not check_file:

            msg = tk.messagebox.askquestion(
                "Studies saved", f"This study {self.studyN} isn't saved\nDo you want to save it?")
            if msg == "yes":
                webbrowser.open(f"{url} ?StudyID={self.studyN}")
                while ex == False:
                    msg = tk.messagebox.askquestion(
                        "Saving Study", f"Are you sure that the excel file\nof this study {self.studyN} is downloaded?")
                    if msg == "yes":
                        try:
                            if os.path.isfile(pathDownFile):
                                Downfile = pathDownFile
                            if os.path.isfile(pathDownFile1):
                                Downfile = pathDownFile1

                            shutil.copy2(Downfile, pathfile +
                                         str(self.studyN) + '.xlsx')
                            os.remove(Downfile)
                            ex = True
                            messagebox.showinfo(message="Job DONE!!",
                                                title="SSaving Study")
                            msg = tk.messagebox.askquestion(
                                "Studies saved", f"Do you want to open it?")
                            if msg == "yes":
                                nameFile = str(pathfile) + \
                                    str(self.studyN) + '.xlsx'
                                self.entry.insert(0, nameFile)
                                self.filename = nameFile
                                self.open_file()
                        except:
                            ex = False
                    else:
                        ex = True
            else:

                webbrowser.open(f"{url} ?StudyID={self.studyN}")

        else:
            msg = tk.messagebox.askquestion(
                "Studies stored", f"This study {self.studyN} is saved\nDo you want to open it?")
            if msg == "yes":
                nameFile = str(pathfile) + str(self.studyN) + '.xlsx'
                self.entry.insert(0, nameFile)
                self.filename = nameFile
                self.open_file()
            else:
                if self.studyN != None:
                    webbrowser.open(f"{url} ?StudyID={self.studyN}")

    def CleanTXT(self):
        self.entry.delete('0', END)
        self.entry1.delete('0', END)

    def Choose_File(self):
        self.CleanTXT()
        try:
            self.filename = filedialog.askopenfilename(title="Open file")
            self._path = self.filename.split(".")
            if self._path[1] == "xlsx":
                self.entry.insert(0, str(self.filename))
            else:
                messagebox.showerror(
                    message="This file is not EXCEL", title="Error extension file")
        except:
            return

    def FillCombos(self, id):
        global mat
        ret = []
        i = 0
        if id == 1:
            for i in range(0, 24):
                ret.append(str(i))
                return ret
        if id == 2:
            for i in range(0, 60):
                ret.append(str(i))
        if id == 3:
            for m in mat:
                mm = m.split(",")
                ret.append(str(mm[0]))
        return ret

    def updateCombo():
        return

    def GetData(self, spl1):

        global data
        global SampleN
        dat = []

        if "Sample" in str(spl1):

            dat = spl1[1].split("\n")
            self.IDsample = (str(dat[0]).replace("Sample ID: ", ""))
            if self.IDsample != "":
                self.IDsample = (str(dat[0]).replace("Sample ID: ", ""))
                SampleN = True
                return
            else:
                SampleN = False
                return

        if "Layer 01" in spl1:
            data = "/01/" + str(spl1[5]) + "/" + str(spl1[6]) + "/" + str(
                spl1[8]) + "/" + str(spl1[10]) + "/" + str(spl1[13]) + "/" + str(spl1[12])
            return data
        if "Layer 02" in spl1:
            data = "/02/" + str(spl1[5]) + "/" + str(spl1[6]) + "/" + str(
                spl1[8]) + "/" + str(spl1[10]) + "/" + str(spl1[13]) + "/" + str(spl1[12])
            return data
        if "Layer 03" in spl1:
            data = "/03/" + str(spl1[5]) + "/" + str(spl1[6]) + "/" + str(
                spl1[8]) + "/" + str(spl1[10]) + "/" + str(spl1[13]) + "/" + str(spl1[12])
            return data
        if "Layer 04" in spl1:
            data = "/04/" + str(spl1[5]) + "/" + str(spl1[6]) + "/" + str(
                spl1[8]) + "/" + str(spl1[10]) + "/" + str(spl1[13]) + "/" + str(spl1[12])
            return data
        if "Layer 05" in spl1:
            data = "/05/" + str(spl1[5]) + "/" + str(spl1[6]) + "/" + str(
                spl1[8]) + "/" + str(spl1[10]) + "/" + str(spl1[13]) + "/" + str(spl1[12])
            return data
        if "Layer 06" in spl1:
            data = "/06/" + str(spl1[5]) + "/" + str(spl1[6]) + "/" + str(
                spl1[8]) + "/" + str(spl1[10]) + "/" + str(spl1[13]) + "/" + str(spl1[12])
            return data
        if "Layer 07" in spl1:
            data = "/07/" + str(spl1[5]) + "/" + str(spl1[6]) + "/" + str(
                spl1[8]) + "/" + str(spl1[10]) + "/" + str(spl1[13]) + "/" + str(spl1[12])
            return data
        if "Layer 08" in spl1:
            data = "/08/" + str(spl1[5]) + "/" + str(spl1[6]) + "/" + str(
                spl1[8]) + "/" + str(spl1[10]) + "/" + str(spl1[13]) + "/" + str(spl1[12])
            return data
        if "Layer 09" in spl1:
            data = "/09/" + str(spl1[5]) + "/" + str(spl1[6]) + "/" + str(
                spl1[8]) + "/" + str(spl1[10]) + "/" + str(spl1[13]) + "/" + str(spl1[12])
            return data

    def Get_StudyNumber(self, rows):
        # print("----- > " + str(rows))
        row1 = []
        # row1 = rows.split("#")
        # entry1.insert(0,str(row1[1]) + "_Exported" )

    def Get_Engineer(self, rows):
        global Owner
        if Owner != "":
            return
        for row1 in rows:
            if "Owner" in row1:
                own = row1.split(" ")
                Owner = str(own[2]) + str(own[3])
                self.Label_Engineer.set(Owner)
                LogText = "Owner of the Study : " + str(Owner)
                print(LogText)
                self.set_Log_Text(LogText)

    def open_file(self):

        global data
        global Base
        global tagn
        global wordToFind
        global input1
        global input2
        global Owner

        self.DatoStudy = []
        rep = ""
        mat.clear()
        mats.clear()
        Owner = ""
        SamplesN.clear()
        self.entry3.delete('0', END)
        pas = self.CheckBoxes(False)
        if pas == "":
            return
        try:
            workbook = xlrd.open_workbook(self.filename)
            worksheet = workbook.sheet_by_name('Sheet2')
            self.master.state("zoomed")

        except:
            messagebox.showerror(
                message="Problems opening the Excel file.\nPlease check the file.", title="Error opening file")
            return

        print("Opening file ......")
        self.set_Log_Text("Opening file ......")
        num_rows = worksheet.nrows - 1
        if num_rows < 9:
            messagebox.showerror(
                message="Please check the Excel file,\ncould be not data inside.", title="Error opening file")
            return
        curr_row = -1
        row = []
        curr_row = 0
        Base = True
        if self.tree.exists:
            self.tree.destroy()
            print("Creating Table......")
            self.set_Log_Text("Creating Table......")
            self.CreateTable()
        self.DatoStudy.clear()
        nsamples = 0
        msg_box = ""
        mfi_boolean = False
        den_boolean = False
        mfi_den_boolean = False

        while curr_row < num_rows:
            curr_row += 1
            row = worksheet.row_values(curr_row)
            if curr_row > 2 and curr_row < 6:
                self.Get_Engineer(row)

            value_missing = 0
            pasa = False
            pasa1 = False

            if "Sample ID:" in str(row[1]):  # and len(str(row[1])) == 10:
                dat1 = self.GetData(row)

            if "Layer" in str(row[1]) and len(str(row[1])) == 8:

                rowM = row[10]
                if rowM == '' or rowM == None:
                    value_missing = 1
                rowM = row[8]
                if rowM == '' or rowM == None:
                    value_missing = value_missing + 2

                # check Density value
                if self.Check_Valor(row[8], "Density") == False:

                    for n, d, t, b in mats:
                        if row[6] == n and b == True and t == "Density":
                            row[8] = self.InputMFI_DENSITY(
                                row[6], "Density", True)
                            dat1 = self.GetData(row)
                            pasa = True

                    if pasa == False:

                        msg_box = tk.messagebox.askquestion('Wrong Value', f'This Density value {rowM} for this material {row[6]} is wrong.\nDo you want to change it?',
                                                            icon='warning', parent=self.master)
                        if msg_box == "yes":
                            row[8] = self.InputMFI_DENSITY(
                                row[6], "Density", True)

                # Check MFI value
                if self.Check_Valor(row[10], "MFI") == False:

                    for n, d, t, b in mats:
                        if row[6] == n and b == True and t == "MFI":
                            row[10] = self.InputMFI_DENSITY(
                                row[6], "MFI", True)
                            dat1 = self.GetData(row)
                            pasa1 = True

                    if pasa1 == False:

                        msg_box = tk.messagebox.askquestion('Wrong Value', f'This MFI value {rowM} for this material {row[6]} is wrong.\nDo you want to change it?',
                                                            icon='warning', parent=self.master)
                        if msg_box == "yes":
                            row[10] = self.InputMFI_DENSITY(
                                row[6], "MFI", True)

                if mfi_boolean == False and value_missing == 1:
                    msg_box = tk.messagebox.askquestion('Mfi missing', 'Some Mfi values are missing.\nDo you want to fill all of them?',
                                                        icon='warning', parent=self.master)
                    if msg_box == 'yes':
                        mfi_boolean = True
                        input1 = True
                    else:
                        mfi_boolean = True
                        input1 = False

                elif den_boolean == False and value_missing == 2:
                    msg_box = tk.messagebox.askquestion('Density missing', 'Some Density values are missing.\nDo you want to fill all of them?',
                                                        icon='warning', parent=self.master)
                    if msg_box == 'yes':
                        den_boolean = True
                        input1 = True
                    else:
                        den_boolean = True
                        input1 = False

                elif mfi_den_boolean == False and value_missing == 3:
                    msg_box = tk.messagebox.askquestion('Mfi & Density missing', 'Some Mfi & Density values are missing.\nDo you want to fill all of them?',
                                                        icon='warning', parent=self.master)
                    if msg_box == 'yes':
                        mfi_den_boolean = True
                        den_boolean = True
                        mfi_boolean = True
                        input1 = True
                    else:
                        mfi_den_boolean = True
                        input1 = False

                if input1 == True:
                    if value_missing == 1:
                        row[10] = str(self.InputMFI_DENSITY(row[6], "Mfi"))
                    elif value_missing == 2:
                        row[8] = str(self.InputMFI_DENSITY(row[6], "Density"))
                    elif value_missing == 3:
                        row[10] = str(self.InputMFI_DENSITY(row[6], "Mfi"))
                        row[8] = str(self.InputMFI_DENSITY(row[6], "Density"))
                else:
                    if row[10] == "" or row[10] == None:
                        row[10] = 0
                    if row[8] == "" or row[8] == None:
                        row[8] = 0

                dat1 = self.GetData(row)

            # print("dat1 ----" + str(dat1))
                if str(dat1) != "None":
                    if self.IDsample != "":
                        if self.IDsample == rep:
                            data = str("-----") + str(dat1)
                            data1 = str(self.IDsample) + str(dat1)
                            self.AddData(str(data), str(data1))

                        else:
                            data = str(self.IDsample) + str(dat1)
                            self.AddData(str(data), str(data))
                            rep = self.IDsample
                            nsamples = nsamples + 1
        self.entry3.insert(0, str(nsamples))
        self.FillCombos(3)
        self.Materials['values'] = mat
        self.Materials.current(0)
        self.SamplesCO['values'] = SamplesN
        self.SamplesCO.current(0)
        print("Número de Rows ----- " + str(len(self.DatoStudy)))
        self.set_Log_Text("Número de Rows ----- " + str(len(self.DatoStudy)))
        print("Número de Samples ----- " + str(nsamples))
        self.set_Log_Text("Número de Samples ----- " + str(nsamples))

    def InputMFI_DENSITY(self, name, tip, bo=False):

        global input1
        global input2

        pas = False

        for n, t in matM:
            if n == name and t == tip:
                pas = True
                break
        if pas == False:
            ad = name, tip
            matM.append(ad)
            valor = self.CheckMFI_DEN(name, tip)

            if valor != "":
                ins = name, valor, tip, bo
                mats.append(ins)

                return valor
            valor = askfloat(
                f'{tip} missing', f'This component - {name} has not Value.\nPlease entry the new value.', parent=self.master)
            if valor == None:
                valor = 0
            ins = name, valor, tip, bo
            mats.append(ins)
            self.Save_MFI_Density(name, valor, tip)
            return valor
        else:
            for txt in mats:
                if txt[0] == name and txt[2] == tip:
                    return txt[1]

    def Check_Valor(self, valor, types):

        if valor == None:
            return True

        if types == "Density":
            if (not "." in str(valor) and str(valor) != ""):
                if str(valor).isdigit:
                    if int(valor) in range(1, 50):
                        return True
                return False
            else:
                return True

        if types == "MFI":
            if (not "." in str(valor) and str(valor) != ""):
                if str(valor).isdigit:
                    if int(valor) in range(1, 50):
                        return True
                return False
            else:
                return True

    def CheckMFI_DEN(self, name, tip):
        valor = ""
        mat1 = f'C:\Python\{tip}/{name}.txt'
        check_file = os.path.isfile(mat1)
        if check_file == True:

            text_file = open(mat1, "r")
            valor = text_file.read()
            text_file.close
            msg_box = tk.messagebox.askquestion(f'{tip} missing', f'There is {tip} value stored for this Material.\n {name} value = {valor}\nDo you want to use it?',
                                                icon='warning', parent=self.master)
            if msg_box == "yes":
                return valor
        return ""

    def Save_MFI_Density(self, mat, valor, tip):

        mat1 = f'C:\Python\{tip}/{mat}.txt'
        text_file = open(mat1, "w")
        text_file.write(str(valor))
        text_file.close

    def CheckBoxes(self, condition):

        # XLSX File to open
        if str(self.entry.get()) == "":
            messagebox.showerror(
                message="Please, choose file to Process!!", title="Error choosing file")
            return ""
        else:
            self._path = str(self.entry.get())
            # Name of the XLSX
        self.NameFile = str(self.entry1.get())
        if self.NameFile == "" and condition == True:
            messagebox.showerror(
                message="Please, fill the name file box!!", title="Error file name")
            return "1"

    def SaveToExcel(self):

        pas = self.CheckBoxes(True)
        if pas == "" or pas == "1":
            return
        else:
            if len(self.tree.get_children()) == 0:
                messagebox.showerror(
                    message="Please, open the file to Process!!", title="Error file exporting")
                return

            folder = os.path.dirname(self.filename)
            Export_filename = folder + "/" + self.NameFile + "_Exported.xlsx"
            LogText = "File ---- " + str(Export_filename)
            print(LogText)
            self.set_Log_Text(LogText)
            workbook = xlsxwriter.Workbook(Export_filename)
            worksheet = workbook.add_worksheet(self.NameFile)

            bold = workbook.add_format(
                {"bold": True, "font_size": 12, "align": CENTER, 'border': 1})
            headings = ['Sample ID', 'Layer', 'Percentage',
                        'Amount %', 'Density', 'Melt Index', 'Resin']
            worksheet.write_row("A1", headings, bold)
            bold = workbook.add_format(
                {"bold": False, "font_size": 10, "align": CENTER, 'border': 1})
            for row_id in self.tree.get_children():
                row = self.tree.item(row_id)['values']
                worksheet.write_row(int(row_id)+1, 0, row, bold)
                worksheet.autofit()
            workbook.close()
            messagebox.showinfo(message="Job DONE!!",
                                title="Status Conversion")
            print("Excel created ...")
            self.set_Log_Text("Excel created ...")

    def CreateTable(self):

        self.tree = ttk.Treeview(
            self.master, selectmode='browse')
        self.tree.grid(row=4, column=0, columnspan=13,
                       sticky="nsew", padx=15, pady=10, ipadx=0)
        vsb = ttk.Scrollbar(self.master, orient="vertical",
                            command=self.tree.yview)
        vsb.grid(row=4, column=12, sticky="ns", pady=10)

        self.tree.configure(yscrollcommand=vsb.set)

        self.tree['columns'] = ('Sample ID', 'Layer', 'Percentage %',
                                'Amount %', 'Density', 'Melt Index', 'Resin')

        # format our column
        self.tree.column("#0", width=0,  stretch=NO)
        self.tree.column("Sample ID", anchor=CENTER, width=100)
        self.tree.column("Layer", anchor=CENTER, width=80)
        self.tree.column("Percentage %", anchor=CENTER, width=80)
        self.tree.column("Amount %", anchor=CENTER, width=80)
        self.tree.column("Density", anchor=CENTER, width=90)
        self.tree.column("Melt Index", anchor=CENTER, width=90)
        self.tree.column("Resin", anchor=CENTER, width=360)

    # Create Headings
        self.tree.heading("#0", text="", anchor=CENTER)
        self.tree.heading("Sample ID", text="Sample ID", anchor=CENTER)
        self.tree.heading("Layer", text="Layer", anchor=CENTER)
        self.tree.heading("Percentage %", text="Percentage %", anchor=CENTER)
        self.tree.heading("Amount %", text="Amount %", anchor=CENTER)
        self.tree.heading("Density", text="Density g/10 m", anchor=CENTER)
        self.tree.heading(
            "Melt Index", text="Melt Index g/cm³", anchor=CENTER)
        self.tree.heading("Resin", text="Resin/Additives", anchor=CENTER)

    def AddData(self, dataSample, dataSample2):
        # add data
        global a
        global tagn

        global mat
        global SampleN

        dato = []
        dato1 = []
        dato = dataSample.split("/")
        dato1 = dataSample2.split("/")
        if Base == True:
            self.DatoStudy.append(dato1)
            if dato1[3] not in mat:
                mat.append(dato1[3])
            if dato[0] not in SamplesN:
                if dato[0] != "-----":
                    SamplesN.append(dato[0])
        if tagn == False:
            tagn = True
            self.tree.insert(parent="", index='end', iid=a, text='', values=(
                dato[0], dato[1], dato[2], dato[6], dato[4], dato[5], dato[3]), tag='gray')
        else:
            tagn = False
            self.tree.insert(parent="", index='end', iid=a, text='', values=(
                dato[0], dato[1], dato[2], dato[6], dato[4], dato[5], dato[3]))
            self.tree.tag_configure('gray', background='#cccccc')
        a = a + 1

    def SortBy(self, words):

        global IDsample1
        global BackUP
        global f
        global ListSort
        global Sorting

        global a
        global Base
        global IDsamples
        global tagn
        global wordToFind
        global SamplesN

        if len(self.SamplesCO["values"]) <= 0:

            messagebox.showerror(
                message="Please, open the file to Process!!", title="Not Samples available")
            return

        tagn = False
        Base = False
        SamplesN.clear()
        self.entry3.delete('0', END)

        a = 0
        self.tree.delete(*self.tree.get_children())
        SumMI = 0.0
        LogText = "Sorting ---- " + str(words)
        print(LogText)
        self.set_Log_Text(LogText)

        # self._log.Set_Text(self, "Sorting ---- " + str(words))

        SumTotal = 0
        IDsamples = ""
        l = 0
        kk = 0
        if words == "Layer":
            wordToFind = "0" + str(self.InterEntry.get())
            l = 1
            kk = 5
        if words == "Resin":
            wordToFind = str(self.Materials.get().strip())
            l = 3
        if words == "Density":
            wordToFind = "0" + str(self.InterEntry.get())
            l = 1
            kk = 4

        nrowss = len(self.DatoStudy)
        for i in range(0, nrowss):
            txt = self.DatoStudy[i]
            # break
            if txt[l] == wordToFind:

                if IDsamples == "":
                    IDsamples = str(txt[0])
                if str(txt[0]) == IDsamples:
                    SumMI = self.SumDifferentMI(str(txt[kk]), str(txt[6]))
                    SumTotal = SumTotal + SumMI
                    Idn = IDsamples, SumTotal

                else:
                    Idn = IDsamples, SumTotal
                    SumTotal = self.SumDifferentMI(str(txt[kk]), str(txt[6]))
                    f.append(Idn)
                    IDsamples = str(txt[0])

        Idn = IDsamples, SumTotal
        SumTotal = self.SumDifferentMI(str(txt[kk]), str(txt[6]))
        f.append(Idn)
        if len(f) == 0:
            f.append(Idn)

        # print("Sense Filtro ---- " + str(f))

        if str(self.EntryMelt.get()) == "High":
            rev = True
        else:
            rev = False
        Sorting = sorted(f, key=lambda x: x[1], reverse=rev)
        self.entry3.delete('0', END)
        self.entry3.insert(0, str(len(f)))
        f.clear()

        for item in Sorting:
            for item1 in self.DatoStudy:
                if item[0] == item1[0]:
                    ListSort.append(item1)

        TotalList = ""
        rep1 = ""
        IDsample1 = ""
        c = False

        for list1 in ListSort:
            c = True
            IDsample1 = list1[0]
            if IDsample1 not in SamplesN:
                SamplesN.append(IDsample1)
            for list2 in list1:
                if IDsample1 == rep1 and c == True:
                    list2 = str("-----")
                    c = False
                if TotalList == "":
                    TotalList = list2
                else:
                    TotalList = TotalList + "/" + str(list2)
            rep1 = IDsample1
            self.AddData(str(TotalList), str(TotalList))
            # print(TotalList)
            TotalList = ""
        self.SamplesCO["values"] = SamplesN
        self.SamplesCO.current(0)
        BackUP = self.DatoStudy
        ListSort.clear()
        Sorting.clear()

    def SumDifferentMI(self, mi, per):

        mresult = 0
        if mi == "" or mi == "None":
            mi = 0
        mresult = Decimal(mi)*Decimal(per)/100
        return mresult

    def createModal(self):

        if len(self.SamplesCO["values"]) > 0:
            Sample_Window(self.master, self.SamplesCO,
                          self.DatoStudy, self)
        else:
            messagebox.showerror(
                message="Please, open the file to Process!!", title="Not Sample available")
            return

    def Consola(self):
        global LogOpen

        if LogOpen == True:
            self.new_win.destroy()
            LogOpen = False
        else:
            LogOpen = True
            # Create a TopLevel window
            self.new_win = tk.Toplevel
            self.new_win = self.new_win(self.master)
            self.new_win.resizable(0, 0)
            self.new_win.geometry("265x600+0+130")
            self.new_win.title("Debug window")
            self.button_Clear = Button(
                self.new_win, text="Clear Log", command=self.ClearText).pack(side=BOTTOM)
            self.my_text_box = Text(
                self.new_win, height=27, width=80, font="verdana 10")
            self.my_text_box.pack(fill='both', expand=1)
            self.my_text_box.insert(tk.INSERT, text)

        # if text != "" :
            # self.set_Log_Text(text)

    def ClearText(self):
        global text
        self.my_text_box.delete(1.0, tk.END)
        text = ""

    def set_Log_Text(self, txt):
        global text

        if txt != "":
            self.firstOPen = True
            text = text + txt + "\n"

        try:
            # value = Toplevel.winfo_exists(self.new_win)
            if LogOpen == True:
                self.my_text_box.config(state="normal")
                self.my_text_box.delete(1.0, tk.END)
                self.my_text_box.insert(tk.INSERT, text)
                self.my_text_box.config(state="disabled")
        except:
            return

    def mdotools(self):
        self.master.state("zoomed")
        Mdotools(self.master)

    def open_DataSheet(self):

        msg = tk.simpledialog.askstring(
            "DataSheet finder", "-               Please, enter the resine Name                   -")
        if msg == "" or msg == None:
            return
        msg = str(msg).upper()
        rate = []
        rate.clear()
        option = False
        PFile = r"C:/Python/DataSheet/"  # + self.mat_item + ".pdf"
        contenido = os.scandir(PFile)
        for elemento in contenido:

            elementos = elemento.name[0:-4].upper()
            elementos = elementos.replace("POLYETHYLENE RESIN", "")
            elementos = elementos.replace("POLYOLEFIN PLASTOMER", "")
            elementos = elementos.replace("DOW", "")
            elementos = elementos.replace("ENHANCED", "")

            elementos = str(elementos)

            r = SequenceMatcher(None, str(elementos),
                                msg).ratio()
            namemat = elemento, r
            rate.append(namemat)

        max_ratio = max(rate, key=lambda x: x[1])

        if max_ratio[1] > 0.60:

            webbrowser.open(max_ratio[0])
            option = True

        if option == False:

            self.msg_box = tk.messagebox.askquestion('Datasheet report', "There is not datasheet file for " + "- " + str(msg) + " -"
                                                     + " \nDo you want to find it on the iProduct Quick Search?", icon='warning')
            if self.msg_box == "yes":
                copia.copy(msg)
                webbrowser.open(
                    'https://prodlist.intranet.dow.com/Search/Search.aspx')
            else:
                self.msg_box = tk.messagebox.askquestion(
                    'Datasheet report', "Do you want to open the DataSheet folder\nto find it there? ", icon='info')
                if self.msg_box == "yes":
                    webbrowser.open(str("file:")+PFile)


#######  -------- GUI - Debug Console -------- ######


class OpenConsola(tk.Toplevel):

    global text
    text = ""

    def __init__(self):
        self.resizable(0, 0)
        self.geometry("300x600+0+0")
        self.title("Debug window")
        my_text_box = Text(
            self, height=27, width=80, font="verdana 10")
        my_text_box.grid(row=1, column=0, columnspan=1, sticky='nsew')
        my_text_box.pack(fill='both', expand=1)

    def Set_Text(self, txt):
        global text
        text = text + "\n" + txt
        self.my_text_box.config(state="normal")
        self.my_text_box.delete(1.0, tk.END)
        # This inserts nothing when called from outside class
        self.my_text_box.insert(tk.INSERT, text)
        # But it inserts the correct text when called from this same class
        self.my_text_box.config(state="disabled")

#######  -------- GUI - Sample Window Options -------- ######


class Sample_Window(tk.Toplevel):

    def __init__(self, parent, SamplesCO, DatoStudy, obj1):
        super().__init__(parent)
        self.parent = parent
        self.obk = Window
        self.obj1 = obj1
        iD = SamplesCO.get()
        self.DatoStudy = DatoStudy
        self.Feed = tk.StringVar()
        self.Zone1 = tk.StringVar()
        self.Zone2 = tk.StringVar()
        self.Zone3 = tk.StringVar()
        self.Zone4 = tk.StringVar()
        self.Zone5 = tk.StringVar()
        self.Adapt = tk.StringVar()
        self.Die = tk.StringVar()
        self.Kgh = tk.StringVar()
        self.wr = False
        self.Next_Mix = False
        self.path_to_file = r"C:/Python/Temp/Temp.txt"
        self.filepath2 = r"C:/Python/Temp/Operation.txt"
        self.resizable(0, 0)
        wtotal = self.winfo_screenwidth()
        htotal = self.winfo_screenheight()
        wventana = 800
        hventana = 550
        pwidth = round(wtotal/2-wventana/2)
        pheight = round(htotal/2-hventana/2)
        self.geometry(str(wventana)+"x"+str(hventana) +
                      "+"+str(pwidth)+"+"+str(pheight))
        self.title("Sample Layers Distribution")
        self.boton_Kg = tk.Button(
            self, text="Kg/h\nCalculation", font=("Verdana", 10, "bold"), command=self.Kg_hour)
        self.boton_Kg.place(x=20, y=400)
        self.boton_Kg = tk.Button(self, text="Mixing\nCalculation", font=(
            "Verdana", 10, "bold"), command=self.Kg_Mixing)
        self.boton_Kg.place(x=20, y=450)
        self.boton_Density = tk.Button(self, text="Density/Mfi\nCalculation", font=(
            "Verdana", 10, "bold"), command=self.Denisty_MFI)
        self.boton_Density.place(x=20, y=500)
        self.boton_Temp = tk.Button(self, text=" Edit TEMP  \nprofile", font=(
            "Verdana", 10, "bold"), command=partial(self.Fill_Temp, True))
        self.boton_Temp.place(x=165, y=400)
        self.boton_Temp2 = tk.Button(self, text="Show\nDataSheet  ", font=(
            "Verdana", 10, "bold"), command=self.Open_Datasheet)
        self.boton_Temp2.place(x=165, y=450)
        self.boton_coment = tk.Button(self, text="Show\nComments ", font=(
            "Verdana", 10, "bold"), command=self.Open_Comments)
        self.boton_coment.place(x=165, y=500)
        self.label_Sample = tk.Label(
            self, text="Sample : " + str(iD), font=("Verdana bold", 18))
        self.label_Sample.place(x=20, y=5)
        self.label_Item = tk.Label(
            self, text="Layer Calculation", font=("Verdana bold", 10))
        self.label_Item.place(x=370, y=10)
        self.label_Kg = tk.Label(self, text="for", font=("Verdana bold", 10))
        self.label_Kg.place(x=640, y=10)
        self.label_Kg1 = tk.Label(
            self, textvariable=self.Kgh, font=("Verdana bold", 10))
        self.label_Kg1.place(x=670, y=10)
        self.Multi = tk.Entry(self, justify='center',
                              font=("Verdana bold", 10))
        self.Multi.place(x=570, y=10, w=60)
        self.ItemCO = ttk.Combobox(
            self, justify='center', state='readonly', font=("Verdana bold", 10))
        self.ItemCO.place(x=510, y=10, w=50)
        self.RawMat = tk.Entry(self, justify='center',
                               font=("Verdana bold", 10))
        self.RawMat.place(x=300, y=400, w=460)
        self.label_Feed = tk.Label(
            self, textvariable=self.Feed, font=("Verdana bold", 10))
        self.label_Feed.place(x=320, y=440)
        self.label_Zones = tk.Label(
            self, textvariable=self.Zone1, font=("Verdana bold", 10))
        self.label_Zones.place(x=320, y=460)
        self.label_Zones = tk.Label(
            self, textvariable=self.Zone2, font=("Verdana bold", 10))
        self.label_Zones.place(x=320, y=480)
        self.label_Zones = tk.Label(
            self, textvariable=self.Zone3, font=("Verdana bold", 10))
        self.label_Zones.place(x=320, y=500)
        self.label_Zones = tk.Label(
            self, textvariable=self.Zone4, font=("Verdana bold", 10))
        self.label_Zones.place(x=530, y=440)
        self.label_Zones = tk.Label(
            self, textvariable=self.Zone5, font=("Verdana bold", 10))
        self.label_Zones.place(x=530, y=460)
        self.label_Die = tk.Label(
            self, textvariable=self.Adapt, font=("Verdana bold", 10))
        self.label_Die.place(x=530, y=480)
        self.label_Adapt = tk.Label(
            self, textvariable=self.Die, font=("Verdana bold", 10))
        self.label_Adapt.place(x=530, y=500)

        self.focus()
        self.grab_set()

        self.treeTop = ttk.Treeview(self, selectmode='browse')
        self.treeTop['columns'] = (
            'Layer', 'Percentage %', 'Amount %', 'Density', 'Melt Index', 'Total Kg/h', 'Resin')
        self.treeTop.place(x=20, y=40, height=350)
        self.vsb1 = ttk.Scrollbar(
            self, orient="vertical", command=self.treeTop.yview)
        self.vsb1.place(x=770, y=41, height=348)
        self.treeTop.configure(yscrollcommand=self.vsb1.set)
        self.treeTop.bind('<ButtonRelease-1>', self.Select_Item)

        # format our column
        self.treeTop.column("#0", width=0,  stretch=NO)
        self.treeTop.column("Layer", anchor=CENTER, width=50)
        self.treeTop.column("Percentage %", anchor=CENTER, width=80)
        self.treeTop.column("Amount %", anchor=CENTER, width=80)
        self.treeTop.column("Density", anchor=CENTER, width=90)
        self.treeTop.column("Melt Index", anchor=CENTER, width=100)
        self.treeTop.column("Total Kg/h", anchor=CENTER, width=80)
        self.treeTop.column("Resin", anchor=CENTER, width=280)

        # Create Headings
        self.treeTop.heading("#0", text="", anchor=CENTER)
        self.treeTop.heading("Layer", text="Layer", anchor=CENTER)
        self.treeTop.heading(
            "Percentage %", text="Percentage %", anchor=CENTER)
        self.treeTop.heading("Amount %", text="Amount %", anchor=CENTER)
        self.treeTop.heading("Density", text="Density g/10 m", anchor=CENTER)
        self.treeTop.heading(
            "Melt Index", text="Melt Index g/cm³", anchor=CENTER)
        self.treeTop.heading("Total Kg/h", text="Total Kg/h", anchor=CENTER)
        self.treeTop.heading("Resin", text="Resin/Additives", anchor=CENTER)
        self.TKg_mix = []
        self.TKg_Final = []
        self.dad = ["All", "01", "02", "03",
                    "04", "05", "06", "07", "08", "09"]
        self.ItemCO['values'] = self.dad
        self.ItemCO.current(0)
        self.ItemCO.bind('<<ComboboxSelected>>',
                         lambda event: self.Layer_Calculation(event, iD))
        self.Layer_Calculation(self, iD)
        self.Next_Mix = False

    def Denisty_MFI(self):
        item1 = ""
        newList = []
        self.density = []
        self.mfiList = []
        nlayer = ""
        for row_id in self.treeTop.get_children():
            items = self.treeTop.item(row_id)['values']
            if item1 == "":
                item1 = items[0]
            if items[0] == item1:
                if items[3] == 'None':
                    items[3] = 0
                if items[4] == 'None':
                    items[4] = 0

                self.Sum_denisty(Decimal(items[2]), Decimal(
                    items[3]), False)
                self.Sum_MFI(Decimal(items[2]), Decimal(
                    items[4]),  False)
                if nlayer == "":
                    nlayer = str(items[0])
                else:
                    nlayer = nlayer + "+" + str(items[0])
            else:
                if items[3] == 'None':
                    items[3] = 0
                if items[4] == 'None':
                    items[4] = 0
                den1 = self.Sum_denisty(Decimal(items[2]), Decimal(
                    items[3]),  True)
                mfi1 = self.Sum_MFI(Decimal(items[2]), Decimal(
                    items[4]),  True)
                item1 = items[0]
                ap = nlayer, (round(den1, 3)), (round(mfi1, 3))
                newList.append(ap)
                nlayer = ""
                nlayer = str(items[0])
                print(ap)
        if items[3] == 'None':
            items[3] = 0
        if items[4] == 'None':
            items[4] = 0
        den1 = self.Sum_denisty(Decimal(items[2]), Decimal(
            items[3]),  True)
        mfi1 = self.Sum_MFI(Decimal(items[2]), Decimal(
            items[4]),  True)

        ap = nlayer, (round(den1, 3)), (round(mfi1, 3))
        newList.append(ap)
        print(ap)
        ShowDensity(self, newList)

    def Sum_denisty(self, amo, den, ret):
        dpl = []
        if ret == True:
            for a, d in self.density:
                if d > 0:
                    dp = (Decimal(a)/100)/Decimal(d)
                    dpl.append(dp)
                else:
                    dpl.append(0)
            if sum(dpl) != 0:
                dens = 1/(sum(dpl))
                self.density.clear()
                tot = amo, round(den, 3)
                self.density.append(tot)
                return dens
            else:
                return 0
        tot = amo, round(den, 3)
        self.density.append(tot)

    def Sum_MFI(self, amo, mfi, ret):
        dpl = []
        if ret == True:
            for a, m in self.mfiList:
                if m > 0:
                    dp = (float(a)/100)*((float(m))**(-0.277))
                    dpl.append(dp)
                else:
                    dpl.append(0)
            if sum(dpl) != 0:
                mfis = (sum(dpl))**(-1/0.277)
                self.mfiList.clear()
                totM = amo, round(mfi, 3)
                self.mfiList.append(totM)  # totM
                return mfis
            else:
                return 0
        totM = amo, round(mfi, 3)
        self.mfiList.append(totM)  # totM

    def Layer_Calculation(self, event, iD):

        self.multi_2 = False
        multi = self.Multi.get()
        multi_1 = multi.split(",")
        if len(multi_1) > 1:
            self.multi_2 = True

        self.Kgh.set(str("0.0") + " Kg")
        self.Next_Mix = False
        self.treeTop.delete(*self.treeTop.get_children())
        iD1 = self.ItemCO.get()
        ab = 0
        tag1 = True
        dato = []
        nrowss = len(self.DatoStudy)
        for i in range(0, nrowss):
            dato = self.DatoStudy[i]
            if self.multi_2 == False:
                if iD1 == "All":
                    iD2 = dato[1]
                else:
                    iD2 = iD1
                if (dato[0] == iD and dato[1] == iD2):
                    if tag1 == True:
                        tag1 = False
                        self.treeTop.insert(parent='', index='end', iid=ab, text='', values=(
                            dato[1], dato[2], dato[6], dato[4], dato[5], "", dato[3]), tag='gray')
                    else:
                        tag1 = True
                        self.treeTop.insert(parent='', index='end', iid=ab, text='', values=(
                            dato[1], dato[2], dato[6], dato[4], dato[5], "", dato[3]))
                        self.treeTop.tag_configure(
                            'gray', background='#cccccc')
                    ab = ab + 1
            else:
                for mu in multi:
                    iD2 = mu
                    if (dato[0] == iD and dato[1] == "0" + iD2):
                        if tag1 == True:
                            tag1 = False
                            self.treeTop.insert(parent='', index='end', iid=ab, text='', values=(
                                dato[1], dato[2], dato[6], dato[4], dato[5], "", dato[3]), tag='gray')
                        else:
                            tag1 = True
                            self.treeTop.insert(parent='', index='end', iid=ab, text='', values=(
                                dato[1], dato[2], dato[6], dato[4], dato[5], "", dato[3]))
                            self.treeTop.tag_configure(
                                'gray', background='#cccccc')
                        ab = ab + 1

    def Check_Item_Tree(self):
        curItem = self.treeTop.focus()
        loc_value = self.treeTop.item(curItem)
        Value = loc_value.get("values")
        if Value == "":
            messagebox.showinfo(message="Please, select a Layer of the Sample",
                                title="Layer not Selected", parent=self)
            return False
        else:
            True

    def Open_Datasheet(self, op=False):

        if self.Check_Item_Tree() == False and op == False:
            return

        rate = []
        rate.clear()
        option = False
        PFile = r"C:/Python/DataSheet/"  # + self.mat_item + ".pdf"
        contenido = os.scandir(PFile)
        for elemento in contenido:
            # print(str(elemento.name))

            elementos = elemento.name[0:-4]
            r = SequenceMatcher(None, str(elementos),
                                self.mat_item).ratio()
            rate.append(r)
            # print(r)
            if r > 0.96:
                LogtText = f"Rate search pass ---  {r}"
                print(LogtText)
                self.obk.set_Log_Text(self.obj1, LogtText)
                webbrowser.open(elemento)
                option = True
                break
        if option == False:
            LogtText = f"Rate search fail ---  {max(rate)} ---  {self.mat_item}"
            print(LogtText)
            self.obk.set_Log_Text(self.obj1, LogtText)

            self.msg_box = tk.messagebox.askquestion('Datasheet report', "There is not datasheet file for " + "- " + str(self.mat_item) + " -"
                                                     + " \nDo you want to find it?", icon='warning', parent=self)
            if self.msg_box == "yes":
                copia.copy(self.mat_item)
                webbrowser.open(
                    'https://prodlist.intranet.dow.com/Search/Search.aspx')
            else:
                self.msg_box = tk.messagebox.askquestion(
                    'Datasheet report', "Do you want to open the DataSheet folder\nto find it there? ", icon='info', parent=self)
                if self.msg_box == "yes":
                    webbrowser.open(PFile)

    def Kg_hour(self):
        ab = 0
        TotalKG = 0
        TotalKG = askfloat(
            'Total Kg', '-       Please, entry the Kg/hour.       -')
        if TotalKG == None:
            TotalKG = 0.0
            return
        self.Kgh.set(str(TotalKG) + " Kg")
        self.Next_Mix = True
        t = len(self.treeTop.get_children())
        for row_id in self.treeTop.get_children():
            datos = self.treeTop.item(row_id)['values']
            dat = Decimal(datos[2])
            dat1 = Decimal(datos[1])
            if self.ItemCO.get() == "All" and self.multi_2 == False:
                dat = (((dat*Decimal(TotalKG))/100)*dat1/100)
                dat = round(dat, 4)
            else:
                dat = ((Decimal(TotalKG))*(dat/100))
                dat = round(dat, 4)

            if dat < 1:
                dat = dat*Decimal(1000)
                dat = str(round(dat, 1)) + " grams"
            else:
                dat = str(round(dat, 2)) + " Kg"
            self.treeTop.item(row_id, text="blub", values=(
                datos[0], datos[1], datos[2], datos[3], datos[4], dat, datos[6]))
            ab = ab + 1

    def Kg_Mixing(self):

        if self.Next_Mix == False:
            messagebox.showinfo(
                message="Please, set the Kg to mix first!", title="Kg/h not set", parent=self)
            return

        Layers = []
        self.contenedor = []
        self.TotalMix = []
    #### Layers Hacer array con items Sample selecionado ####
        flag = ""
        x = -1
        y = 0
        rot = ""

        for row_id in self.treeTop.get_children():
            row = self.treeTop.item(row_id)['values']
            del row[1]
            del row[2]
            del row[2]
            Layers.append(row)

    #### Contnedor Agrupar materiales por cada capa ####
        for row in Layers:
            if row[0] == flag or flag == "":
                if len(self.contenedor) != 0:
                    rot = str(row[1]), str(row[2]), str(row[3])
                    self.contenedor[-1].extend(rot)
                else:
                    rot = str(row[1]), str(row[2]), str(row[3]), (str(row[2]))
                    self.contenedor.append(row)
            else:
                rot = str(row[1]), str(row[2]), str(row[3])
                self.contenedor.append(row)
                y = y + 1
                LogText = f"one pair ---- {str(y)} --- {rot}"
                print(LogText)
                self.obk.set_Log_Text(self.obj1, LogText)
            flag = row[0]

    #### Borrar Número de capa de cada row ####
        for row in self.contenedor:
            del row[0]
    # Hacer copia de array self.contenedor
        contenedor1 = copy.deepcopy(self.contenedor)

    #### Borrar array con datos kg por capa y buscar por cada capa ####
        self.TKg_mix.clear()
        self.TKg_Final.clear()
        for rows1 in contenedor1:
            # print("---- > " + str(rows1))
            x = contenedor1.index(rows1)
            self.Total_Kg_Item(rows1, True)
        self.check = []
        f = -1
        for lay in self.TKg_Final:
            f = f + 1
            f1 = -1
            for lay1 in self.TKg_mix:
                f1 = f1 + 1
                if lay == lay1:
                    iguales = f, f1
                    self.check.append(iguales)
        # print("----- " +str(self.check))
        ck = ""
        ck = ""
        self.TKg_mix.clear()
        for jk in self.contenedor:
            self.Total_Kg_Item(jk, False)
        # print( str(self.TKg_mix))
        partialMix = []
        self.Capas = []
        self.Kg_grams = []
        flag2 = ""
        self.Nlayers = ""
        for ls in self.check:
            ck = ls[1]
            if ls[0] != flag2:
                if len(self.Capas) > 0:
                    # self.Nlayers = str(self.Nlayers) + "-" + str(ck)
                    self.Save_Materials(
                        self.Capas, self.Kg_grams, self.Nlayers)
                    self.Nlayers = ""
                    self.Add_mat(self.contenedor, ck)
                else:
                    self.Add_mat(self.contenedor, ck)
                flag2 = ls[0]
            else:
                self.Add_mat(self.contenedor, ck)
                flag2 = ls[0]
                # self.Save_Materials(self.Capas,self.Kg_grams,self.Nlayers)
                # self.Nlayers = ""
        self.Save_Materials(self.Capas, self.Kg_grams, self.Nlayers)
        self.Nlayers = ""
        # print(self.TotalMix)
        self.ventana_secundaria = ShowMatmix(
            self, self.TotalMix, self.check, self.TKg_mix)
#### Añade todos los materiales de cada capa ####

    def Add_mat(self, conten, ck):
        ck1 = -1
        if self.Nlayers == "":
            self.Nlayers = str(int(ck)+1)
        else:
            self.Nlayers = self.Nlayers + "-" + str(int(ck)+1)
        for item in conten:
            ck1 = ck1 + 1  # self.contenedor.index(item)
            if ck1 == ck:
                for it1 in range(1, len(item), 3):
                    self.Capas.append(str(item[it1+1]))
                    self.Kg_grams.append(str(item[(it1)]))
                # print(str(self.Capas))

#### Suma los Kg/gramos de todas las capas ####
    def Save_Materials(self, capas, kg_grams, Nlayers):
        res = dict.fromkeys(capas, 0)
        for a, b in zip(capas, kg_grams):
            res[a] += Decimal(str(b))
        # print(*(f"{key}, {value}" for key, value in res.items()), sep="\n")
        for key, value in res.items():
            join = key, value, Nlayers
            if len(capas) == 2:
                self.TotalMix.append(join)
            self.TotalMix.append(join)
        self.Capas.clear()
        self.Kg_grams.clear()

#### compara listas entre ellas ####
    def Sub_Lista(self, lista1, lista2):
        return [x for x in lista1 if x in lista2]

#### Generar matriz con items con/sin Kg/grams de cada capa y material ####
    def Total_Kg_Item(self, amount, bol):

        for n in amount:
            x = amount.index(n)
            if " Kg" in str(n):
                if bol == True:
                    amount[x] = ""
                else:
                    amount[x] = str(amount[x]).replace(" Kg", "")
            if " grams" in str(n):
                if bol == True:
                    amount[x] = ""
                else:
                    amount[x] = str(amount[x]).replace(" grams", "")

        tkg = amount
        self.TKg_mix.append(tkg)
        if tkg not in self.TKg_Final:
            self.TKg_Final.append(tkg)
        # print ("----->" + str(tkg))
        tkg = ""

    def Temp_Profile(self):

        with open(self.path_to_file) as f:
            for line in f.readlines():
                match1 = line.split(";")
                if match1[0] in self.mat_item:
                    self.Fill_Labels(str(line))
                    LogText = f"Match --- {str(line)}"
                    print(LogText)
                    self.obk.set_Log_Text(self.obj1, LogText)
                    f.close()
                    self.match2 = line
                    return
        LogText = f"Match material {match1[0]} --- No found"
        print(LogText)
        self.obk.set_Log_Text(self.obj1, LogText)
        self.newData = self.Fill_Temp(False)
        if self.newData is not None:
            self.Fill_Labels(self.newData)

        else:
            send = (
                "No Data;No Data;No Data;No Data;No Data;No Data;No Data;No Data;No Data")
            self.Fill_Labels(send)
            f.close()
            return

    def Fill_Labels(self, data):

        match1 = data.split(";")
        self.Feed.set("Temp. Feeding  : " + str(match1[1]).strip() + " °C")
        self.Zone1.set("Temp. Zones 1  : " + str(match1[2]).strip() + " °C")
        self.Zone2.set("Temp. Zones 2  : " + str(match1[3]).strip() + " °C")
        self.Zone3.set("Temp. Zones 3  : " + str(match1[4]).strip() + " °C")
        self.Zone4.set("Temp. Zones 4    : " + str(match1[5]).strip() + " °C")
        self.Zone5.set("Temp. Zones 5    : " + str(match1[6]).strip() + " °C")
        self.Adapt.set("Temp. Adapters  : " + str(match1[7]).strip() + " °C")
        self.Die.set("Temp. Die Head  : " + str(match1[8]).strip() + " °C")

    def Fill_Temp(self, option):

        if self.Check_Item_Tree() == False:
            return

        if option == False:
            self.msg_box = tk.messagebox.askquestion('Profile Temp. not stored', "There is no data about " + str(self.mat_item) + " Temperatures\nDo you want to entry?",
                                                     icon='warning', parent=self)
        if option == True:
            self.msg_box = tk.messagebox.askquestion('Edit Profile Temp', "Do you want to change\n" + str(self.mat_item) + " Temperatures profile",
                                                     icon='warning', parent=self)

        if self.msg_box == 'yes':
            self.NewTemp = ""
            self.entry_temp1 = askinteger(
                "Saving Temp", "-                    Temp. Feeding Zone?                    -", parent=self)
            self.entry_temp2 = askinteger(
                "Saving Temp", "-                     Temp. Zone 1?                         -", parent=self)
            self.entry_temp3 = askinteger(
                "Saving Temp", "-                     Temp. Zone 2?                         -", parent=self)
            self.entry_temp4 = askinteger(
                "Saving Temp", "-                     Temp. Zone 3?                         -", parent=self)
            self.entry_temp5 = askinteger(
                "Saving Temp", "-                     Temp. Zone 4?                         -", parent=self)
            self.entry_temp6 = askinteger(
                "Saving Temp", "-                     Temp. Zone 5?                         -", parent=self)
            self.entry_temp7 = askinteger(
                "Saving Temp", "-                    Temp. Adapter?                         -", parent=self)
            self.entry_temp8 = askinteger(
                "Saving Temp", "-                    Temp. Die Head?                        -", parent=self)
            self.NewTemp = (str(self.mat_item) + ";" + str(self.entry_temp1) + ";" + str(self.entry_temp2) + ";" + str(self.entry_temp3) + ";" + str(self.entry_temp4)
                            + ";" + str(self.entry_temp5) + ";" + str(self.entry_temp6) + ";" + str(self.entry_temp7) + ";" + str(self.entry_temp8))
            # print(self.NewTemp.strip())

            if option == False:
                with open(str(self.path_to_file), "a+") as f:
                    f.write(self.NewTemp)
                    f.write("\n")
                    f.close
                    messagebox.showinfo(
                        message="Profile Temp. saved", title="Data Saved", parent=self)
                    return self.NewTemp
            if option == True:
                self.Replace_Line(self.path_to_file,
                                  self.match2, self.NewTemp + '\n')
                messagebox.showinfo(
                    message="Profile Temp. changed", title="Data Changed", parent=self)
                return self.NewTemp
        return None

    def Replace_Line(self, filepath, oldline, newline):

        # quick parameter checks
        assert os.path.isfile(filepath)          # !
        assert (oldline and str(oldline))  # is not empty and is a string
        assert (newline and str(newline))

        replaced = False
        written = False

        try:

            with open(filepath, 'r+') as f:    # open for read/write -- alias to f

                lines = f.readlines()            # get all lines in file

                if oldline not in lines:
                    pass                         # line not found in file, do nothing

                else:
                    # with open(self.filepath2, 'r+') as f1:  # temp file opened for writing
                    f1 = []
                    for line in lines:           # process each line
                        if line == oldline:        # find the line we want
                            f1.append(newline)   # replace it
                            replaced = True
                        else:
                            f1.append(line)   # write old line unchanged

                    if replaced:                   # overwrite the original file
                        f.seek(0)                    # beginning of file
                        f.truncate()                 # empties out original file

                        for tmplines in f1:
                            f.write(tmplines)
                        f.close()             # writes each line to original file
                        self.Temp_Profile()

                        # tmpfile auto deleted
                    f.close()                          # we opened it , we close it

        except IOError:                 # if something bad happened.
            print("ERROR", IOError)
            self.obk.set_Log_Text(self.obj1, "ERROR " + str(IOError))
            f.close()
            return False

    def Select_Item(self, ar):

        Value = []
        curItem = self.treeTop.focus()
        if curItem == '':
            return
        loc_value = self.treeTop.item(curItem)
        Value = loc_value.get("values")
        self.mat_item = Value[6]
        self.RawMat.delete(0, "end")
        self.RawMat.insert(0, self.mat_item)
        self.Temp_Profile()
        # print("---- > " + str(Value[6]))

    def Open_Comments(self):

        if self.Check_Item_Tree() == False:
            return
        option = False
        PFile = r"C:\Python\Comments/"  # + self.mat_item + ".pdf"
        contenido = os.scandir(PFile)
        for elemento in contenido:
            # print(str(elemento.name))
            elementos = elemento.name.split(".")
            r = SequenceMatcher(None, str(elementos[0]), self.mat_item).ratio()
            # print(r)
            if r > 0.98:
                # webbrowser.open(elemento)
                ShowCommnets(self, elemento)
                option = True
                break
        if option == False:
            self.msg_box = tk.messagebox.askquestion('Resin Comments', "There is no comments for " + str(self.mat_item)
                                                     + " \nDo you want to write it?", icon='warning', parent=self)
            if self.msg_box == "yes":
                elemento = r"C:\Python\Comments/" + str(self.mat_item) + ".txt"
                ShowCommnets(self, elemento)

#######  -------- GUI - Comments Window  -------- ######


class ShowCommnets(tk.Toplevel):

    def __init__(self, parent, path):
        super().__init__(parent)
        self.resizable(0, 0)
        wtotal = self.winfo_screenwidth()
        htotal = self.winfo_screenheight()
        wventana = 800
        hventana = 500
        pwidth = round(wtotal/2-wventana/2)
        pheight = round(htotal/2-hventana/2)
        self.geometry(str(wventana)+"x"+str(hventana) +
                      "+"+str(pwidth)+"+"+str(pheight))
        self.title("Resins Comments")

        self.focus()
        self.grab_set()
        # Creating a text box widget
        self.my_text_box = Text(
            self, height=27, width=80, font="verdana 10 bold")
        self.my_text_box.pack()

        # Create a button to save the text
        save = Button(self, text="Save File",
                      command=self.Save_text, font="verdana 10 bold")
        save.pack()
        self.path = path
        self.Open_text()

    def Open_text(self):
        check_file = os.path.isfile(self.path)
        if check_file == False:
            self.content = ""
            return
        text_file = open(self.path, "r+")
        self.content = text_file.read()
        self.my_text_box.tag_configure("tag_name", justify='center')
        self.my_text_box.insert(END, self.content + "\n")
        self.my_text_box.tag_add("tag_name", "1.0", "end")

        text_file.close()

    def Save_text(self):

        text_get = self.my_text_box.get(1.0, END).rstrip()
        if self.content.rstrip() == text_get:
            self.destroy()
            return
        text_file = open(self.path, "w")
        DayNow = datetime.datetime.now().strftime("%m/%d/%Y, %H:%M:%S")
        text_to_write = f"{text_get}\n{DayNow}\n"
        text_file.write(text_to_write)
        text_file.close()
        self.destroy()

#######  -------- GUI - Mixing Window Results -------- ######


class ShowMatmix(tk.Toplevel):

    def __init__(self, parent, list, check, allLayers):
        super().__init__(parent)
        self.resizable(0, 0)
        wtotal = self.winfo_screenwidth()
        htotal = self.winfo_screenheight()
        wventana = 800
        hventana = 500
        pwidth = round(wtotal/2-wventana/2)
        pheight = round(htotal/2-hventana/2)
        self.geometry(str(wventana)+"x"+str(hventana) +
                      "+"+str(pwidth)+"+"+str(pheight))
        self.title("Weight distribution")
        self.focus()
        self.grab_set()

        Label(master=self, text="Amount Materials Mixing Percentage",
              font=("Verdana bold", 14)).place(x=190, y=20)

#####  agrupar numero de capas #####

        txt = ""
        l1 = ""
        l2 = []
        l = ""

        for cl in check:
            if l1 == "":
                l1 = cl[0]
            if cl[0] == l1:
                if l == "":
                    l = str(int(str(cl[1])) + 1)
                else:
                    l = l + "-" + str(int(str(cl[1])) + 1)
            else:
                l2.append(l)
                l = str(int(str(cl[1])) + 1)
                l1 = cl[0]
        l2.append(l)
        a1 = 0
        b1 = 0
        i = 0
        c = 0
#####  Generar labels con materiales y su peso  #####
        for k in l2:
            txt = "Layers : " + str(k)
            i = i + 22
            ShowMatmix.Labels_fill(self, txt, i, True, c)
            for jk in list:
                if jk[2] == k:
                    b1 = b1 + 1
            sameText = ""
            for n in range(a1, b1):

                txt = str(list[n][0]) + " : " + \
                    ShowMatmix.WeightConversion(str(list[n][1]))
                if txt != sameText:
                    i = i + 22
                    ShowMatmix.Labels_fill(self, txt, i, False, c)
                sameText = txt
            a1 = b1

    def WeightConversion(we):
        nn = str(we).split(".")
        if len(nn[1]) == 1:
            dat = str(we) + " grams"
        else:
            dat = str(we) + " Kg"
        return dat

    def Labels_fill(self, txt, pos, b, c):

        if b == True:
            tk.Label(master=self, text=txt, font=(
                "Verdana bold", 12)).place(x=(40+c), y=(60+pos))
            return
        tk.Label(master=self, text=txt, font=(
            "Verdana", 12)).place(x=(40+c), y=(60+pos))

#######  -------- GUI - MFI & Density Window Results -------- ######


class ShowDensity(tk.Toplevel):

    def __init__(self, parent, lista):
        super().__init__(parent)
        self.resizable(0, 0)
        wtotal = self.winfo_screenwidth()
        htotal = self.winfo_screenheight()
        wventana = 800
        hventana = 500
        pwidth = round(wtotal/2-wventana/2)
        pheight = round(htotal/2-hventana/2)
        self.geometry(str(wventana)+"x"+str(hventana) +
                      "+"+str(pwidth)+"+"+str(pheight))
        self.title("Density & MFI calculations results")
        self.focus()
        self.grab_set()

        Label(master=self, text="Amount Densities & MFI layers",
              font=("Verdana bold", 14)).place(x=190, y=20)
        i = 0
        c = 0
        txt = ""

        for k in lista:
            if i > 264:
                i = 0
                c = 300
            txt = "Layers : " + str(k[0])
            i = i + 22
            ShowDensity.Labels_fill(self, txt, i, True, c)
            i = i + 22
            ShowDensity.Labels_fill(
                self, "New Density : " + str(k[1]) + " g/cm³", i, False, c)
            i = i + 22
            ShowDensity.Labels_fill(
                self, "New MFI : " + str(k[2]) + " g/10 min", i, False, c)

    def Labels_fill(self, txt, pos, b, c):

        if b == True:
            tk.Label(master=self, text=txt, font=(
                "Verdana bold", 12)).place(x=(40+c), y=(60+pos))
            return
        tk.Label(master=self, text=txt, font=(
            "Verdana", 12)).place(x=(40+c), y=(60+pos))


class Mdotools(tk.Toplevel):

    def __init__(self, parent):
        super().__init__(parent)
        self.resizable(0, 0)
        wtotal = self.winfo_screenwidth()
        htotal = self.winfo_screenheight()
        wventana = 800
        hventana = 600
        pwidth = round(wtotal/2-wventana/2)
        pheight = round(htotal/2-hventana/2)
        self.geometry(str(wventana)+"x"+str(hventana) +
                      "+"+str(pwidth)+"+"+str(pheight))
        self.title("MDO tools calculation")
        self.focus()
        self.grab_set()

        Label(self, text="MDO tools calculation",
              font=("Verdana bold", 14)).grid(row=0, column=0, padx=(0, 0), columnspan=6, pady=10)
        Label(self, text="GSM Calculation",
              font=("Verdana bold", 12)).grid(row=2, column=2, padx=(0, 0), columnspan=5, pady=5, sticky="w")
        Label(self, text="Set GSM",
              font=("Verdana bold", 10)).grid(row=3, column=0, padx=(20, ), pady=5, sticky="w")
        Label(self, text="Real GSM",
              font=("Verdana bold", 10)).grid(row=3, column=1, padx=(15, 0), columnspan=1, pady=5, sticky="w")
        Label(self, text="Current Speed",
              font=("Verdana bold", 10)).grid(row=3, column=2, padx=(0, 0), columnspan=1, pady=5, sticky="w")
        Label(self, text="Current Kg/h",
              font=("Verdana bold", 10)).grid(row=3, column=3, padx=(5, 0), columnspan=1, pady=5, sticky="w")
        Label(self, text="Current rpm",
              font=("Verdana bold", 10)).grid(row=3, column=4, padx=(0, 0), columnspan=1, pady=5, sticky="w")
        Label(self, text="New Speed",
              font=("Verdana bold", 10)).grid(row=5, column=2, padx=(15, 0), pady=10, sticky="w")
        Label(self, text="New Kg/h",
              font=("Verdana bold", 10)).grid(row=5, column=3, padx=(10, 0), columnspan=1, pady=10, sticky="w")
        Label(self, text="New rpm",
              font=("Verdana bold", 10)).grid(row=5, column=4, padx=(5, 0), columnspan=1, pady=10, sticky="w")

        # Stretcht ratio

        Label(self, text="Stretch Ratio Calculation",
              font=("Verdana bold", 12)).grid(row=7, column=2, padx=(0, 0), columnspan=5, pady=15, sticky="w")
        Label(self, text="Source Speed",
              font=("Verdana bold", 10)).grid(row=8, column=0, padx=(10, 0), pady=5, sticky="w")
        Label(self, text="Stretch Speed",
              font=("Verdana bold", 10)).grid(row=8, column=1, padx=(10, 0), columnspan=1, pady=5, sticky="w")
        Label(self, text="Stretch RATIO",
              font=("Verdana bold", 10)).grid(row=8, column=2, padx=(10, 0), columnspan=1, pady=5, sticky="w")
        Label(self, text="Micron Stretch Calculation",
              font=("Verdana bold", 12)).grid(row=10, column=2, padx=(0, 0), columnspan=5, pady=15, sticky="w")
        Label(self, text="Set Microns",
              font=("Verdana bold", 10)).grid(row=11, column=0, padx=(10, 0), pady=5, sticky="w")
        Label(self, text="Real Microns",
              font=("Verdana bold", 10)).grid(row=11, column=1, padx=(10, 0), columnspan=1, pady=5, sticky="w")
        Label(self, text="Current Speed",
              font=("Verdana bold", 10)).grid(row=11, column=2, padx=(10, 0), columnspan=1, pady=5, sticky="w")
        Label(self, text="Stretch RATIO",
              font=("Verdana bold", 10)).grid(row=11, column=3, padx=(10, 0), columnspan=1, pady=5, sticky="w")
        Label(self, text="New Speed",
              font=("Verdana bold", 10)).grid(row=13, column=2, padx=(10, 0), columnspan=1, pady=5, sticky="w")
        Label(self, text="New SR",
              font=("Verdana bold", 10)).grid(row=13, column=3, padx=(10, 0), columnspan=1, pady=5, sticky="w")

        # Text to entry GSM
        self.gsm = tk.Entry(self, font="Verdana 10 bold",
                            width=10, justify="center")
        self.gsm.grid(row=4, column=0, columnspan=1, padx=10,
                      pady=0, ipady=5, ipadx=0, sticky="w")
        self.gsm.delete(0, END)
        self.gsm.insert(0, "15")

        self.realgsm = tk.Entry(
            self, font="Verdana 10 bold", width=10, justify="center")
        self.realgsm.grid(row=4, column=1, columnspan=1, padx=10,
                          pady=0, ipady=5, ipadx=0, sticky="w")

        self.speed = tk.Entry(self, font="Verdana 10 bold",
                              width=10, justify="center")
        self.speed.grid(row=4, column=2, columnspan=1, padx=10,
                        pady=0, ipady=5, ipadx=0, sticky="w")
        self.speed.delete(0, END)
        self.speed.insert(0, "0")

        self.real_micron = tk.Entry(self, font="Verdana 10 bold",
                                    width=10, justify="center")
        self.real_micron.grid(row=12, column=1, columnspan=1, padx=10,
                              pady=0, ipady=5, ipadx=0, sticky="w")

        self.micron = tk.Entry(self, font="Verdana 10 bold",
                               width=10, justify="center")
        self.micron.grid(row=12, column=0, columnspan=1, padx=10,
                         pady=0, ipady=5, ipadx=0, sticky="w")
        self.micron_ratio = tk.Entry(self, font="Verdana 10 bold",
                                     width=10, justify="center")
        self.micron_ratio.grid(row=12, column=3, columnspan=1, padx=10,
                               pady=0, ipady=5, ipadx=0, sticky="w")

        self.kgh = tk.Entry(self, font="Verdana 10 bold",
                            width=10, justify="center")
        self.kgh.grid(row=4, column=3, columnspan=1, padx=10,
                      pady=0, ipady=5, ipadx=0, sticky="w")
        self.kgh.delete(0, END)
        self.kgh.insert(0, "0")

        self.rpm = tk.Entry(self, font="Verdana 10 bold",
                            width=10, justify="center")
        self.rpm.grid(row=4, column=4, columnspan=1, padx=0,
                      pady=0, ipady=5, ipadx=0, sticky="w")
        self.rpm.delete(0, END)
        self.rpm.insert(0, "0")

        # Results
        self.Newspeed = tk.Entry(
            self, font="Verdana 10 bold", width=10, justify="center",)
        self.Newspeed.grid(row=6, column=2, columnspan=1, padx=15,
                           pady=0, ipady=5, ipadx=0, sticky="w")

        self.Newkgh = tk.Entry(
            self, font="Verdana 10 bold", width=10, justify="center",)
        self.Newkgh.grid(row=6, column=3, columnspan=1, padx=10,
                         pady=0, ipady=5, ipadx=0, sticky="w")

        self.Newrpm = tk.Entry(
            self, font="Verdana 10 bold", width=10, justify="center",)
        self.Newrpm.grid(row=6, column=4, columnspan=1, padx=0,
                         pady=0, ipady=5, ipadx=0, sticky="w")

        self.Currentspeed1 = tk.Entry(
            self, font="Verdana 10 bold", width=10, justify="center",)
        self.Currentspeed1.grid(row=12, column=2, columnspan=1, padx=15,
                                pady=0, ipady=5, ipadx=0, sticky="w")
        self.Currentspeed1.insert("a", "0")
        # Results microns
        self.Newspeed1 = tk.Entry(
            self, font="Verdana 10 bold", width=10, justify="center",)
        self.Newspeed1.grid(row=14, column=2, columnspan=1, padx=15,
                            pady=0, ipady=5, ipadx=0, sticky="w")

        self.Newratio = tk.Entry(
            self, font="Verdana 10 bold", width=10, justify="center",)
        self.Newratio.grid(row=14, column=3, columnspan=1, padx=10,
                           pady=0, ipady=5, ipadx=0, sticky="w")

        # Button calculate
        self.btnCalculate = tk.Button(self, text="Calculate", font="Verdana 12",
                                      command=self.calculate_gsm)
        self.btnCalculate.grid(row=4, column=5, padx=30, pady=0,
                               ipady=0, ipadx=0, sticky="W")
        # Button calculate SR
        self.btnCalculateSR = tk.Button(self, text="Cal. Stretch ratio", font="Verdana 10",
                                        command=partial(self.calculate_sr, "0"))
        self.btnCalculateSR.grid(row=9, column=3, padx=5, pady=0,
                                 ipady=0, ipadx=0, sticky="W")
        # Button calculate SR1
        self.btnCalculateSR1 = tk.Button(self, text="Cal. Str. speed", font="Verdana 10",
                                         command=partial(self.calculate_sr, "1"))
        self.btnCalculateSR1.grid(row=9, column=4, padx=0, pady=0,
                                  ipady=0, ipadx=0, sticky="W")
        # Button calculate SR2
        self.btnCalculateSR2 = tk.Button(self, text="Cal. Sourc. speed", font="Verdana 10",
                                         command=partial(self.calculate_sr, "2"))
        self.btnCalculateSR2.grid(row=9, column=5, padx=5, pady=0,
                                  ipady=0, ipadx=0, sticky="W")
        # Button calculate Micron SR
        self.btnCalculateMSR = tk.Button(self, text="Cal. Str. Ratio", font="Verdana 10",
                                         command=partial(self.calculate_micron, "1"))
        self.btnCalculateMSR.grid(row=12, column=5, padx=5, pady=0,
                                  ipady=0, ipadx=0, sticky="W")
        # Button calculate Micron Speed
        self.btnCalculateMS = tk.Button(self, text="Cal. Str. Speed", font="Verdana 10",
                                        command=partial(self.calculate_micron, "0"))
        self.btnCalculateMS.grid(row=12, column=4, padx=5, pady=0,
                                 ipady=0, ipadx=0, sticky="W")

        # Stretch Ratio Results
        self.SourceSpeed = tk.Entry(
            self, font="Verdana 10 bold", width=10, justify="center")
        self.SourceSpeed.grid(row=9, column=0, columnspan=1, padx=15,
                              pady=0, ipady=5, ipadx=0, sticky="w")

        self.SRSpeed = tk.Entry(
            self, font="Verdana 10 bold", width=10, justify="center")
        self.SRSpeed.grid(row=9, column=1, columnspan=1, padx=10,
                          pady=0, ipady=5, ipadx=0, sticky="w")

        self.SR = tk.Entry(
            self, font="Verdana 10 bold", width=10, justify="center",)
        self.SR.grid(row=9, column=2, columnspan=1, padx=10,
                     pady=0, ipady=5, ipadx=0, sticky="w")

    def calculate_gsm(self):

        try:
            set_gsm = self.gsm.get().replace(",", ".")
            real_gsm = self.realgsm.get().replace(",", ".")
            speed = self.speed.get().replace(",", ".")
            kgh = self.kgh.get().replace(",", ".")
            rpm = self.rpm.get().replace(",", ".")

            if real_gsm == "0" or real_gsm == "":
                messagebox.showwarning(
                    "GSM values", "The GSM value must be bigger than 0 or not empty", parent=self)
                return

            res = (Decimal(float(speed)) * Decimal(float(real_gsm))) / \
                Decimal(float(set_gsm))

            self.Newspeed.delete(0, END)
            self.Newspeed.insert(0, round(res, 2))

            res = (Decimal(float(kgh)) * Decimal(float(set_gsm))) / \
                Decimal(float(real_gsm))

            self.Newkgh.delete(0, END)
            self.Newkgh.insert(0, round(res, 2))

            res = (Decimal(float(rpm)) * Decimal(float(set_gsm))) / \
                Decimal(float(real_gsm))

            self.Newrpm.delete(0, END)
            self.Newrpm.insert(0, round(res, 2))
        except:
            messagebox.showerror(
                "Values Wrong", "Please, fill with correct values the boxes\nor put 0 before press Calculate button", parent=self)

    def calculate_sr(self, tp):

        source_SR = self.SourceSpeed.get().replace(",", ".")
        srspeed = self.SRSpeed.get().replace(",", ".")
        sr_set = self.SR.get().replace(",", ".")

        if tp == "0":
            if not source_SR.replace(".", "").isdigit() or not srspeed.replace(".", "").isdigit():
                messagebox.showerror(message="Please, fill with correct values!!",
                                     title="No valid value", parent=self)
            else:
                res = Decimal(float(srspeed)) / Decimal(float(source_SR))
                self.SR.delete(0, END)
                self.SR.insert(0, round(res, 2))

        if tp == "1":
            if not source_SR.replace(".", "").isdigit() or not sr_set.replace(".", "").isdigit():
                messagebox.showerror(message="Please, fill with correct values!!",
                                     title="No valid value", parent=self)
            else:
                res = Decimal(float(sr_set)) * Decimal(float(source_SR))
                self.SRSpeed.delete(0, END)
                self.SRSpeed.insert(0, round(res, 2))

        if tp == "2":
            if not srspeed.replace(".", "").isdigit() or not sr_set.replace(".", "").isdigit():
                messagebox.showerror(message="Please, fill with correct values!!",
                                     title="No valid value", parent=self)
            else:
                res = Decimal(float(srspeed)) / Decimal(float(sr_set))
                self.SourceSpeed.delete(0, END)
                self.SourceSpeed.insert(0, round(res, 2))

    def calculate_micron(self, tp):

        try:
            set_micron = self.micron.get().replace(",", ".")
            real_micron = self.real_micron.get().replace(",", ".")
            speed_micron = self.Currentspeed1.get().replace(",", ".")
            stretch_ratio = self.micron_ratio.get().replace(",", ".")

            if real_micron == "0" or real_micron == "" or set_micron == "0" or set_micron == "":
                messagebox.showwarning(
                    "Microns values", "The Microns value must be bigger than 0 or not empty", parent=self)
                return

            if tp == "0":

                res = (Decimal(float(speed_micron)) * Decimal(float(real_micron))) / \
                    Decimal(float(set_micron))
                self.Newspeed1.delete(0, END)
                self.Newspeed1.insert(0, round(res, 2))

            if tp == "1":

                res = (Decimal(float(real_micron)) *
                       Decimal(float(stretch_ratio)))/Decimal(float(set_micron))
                self.Newratio.delete(0, END)
                self.Newratio.insert(0, round(res, 2))

            if tp == "2":
                res = (Decimal(float(rpm)) * Decimal(float(set_gsm))) / \
                    Decimal(float(real_gsm))

                self.Newrpm.delete(0, END)
                self.Newrpm.insert(0, round(res, 2))
        except:
            messagebox.showerror(
                "Values Wrong", "Please, fill with correct values the boxes\nor put 0 before press Calculate button", parent=self)


#######  -------- GUI - Start app -------- ######


def main():

    sql = Sql_Data
    sql.Connection()
    sql.Create_db()
    sql.Insert_data("01-01-2024", "'LDPE 310 E'", 0.86, 0.967)

    root = tk.Tk()
    Window(root)
    root.mainloop()


if __name__ == '__main__':
    main()
