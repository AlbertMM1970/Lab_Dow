
from PyPDF2 import PdfReader
from tkinter import filedialog
import tkinter as tk
from tkinter import *
from tkinter import ttk
import os
from tkinter import messagebox
import xlsxwriter
import datetime
from datetime import timedelta
import webbrowser

path = ""
folder = ""


class PDftoexcel():

    def __init__(self, master):

        self.master = master
        self.master.geometry("150x600")
        self.master.resizable(0, 0)
        self.master.title("Conversor PDF To Excel File")
        #  Obtenemos el largo y  ancho de la pantalla
        wtotal = self.master.winfo_screenwidth()
        htotal = self.master.winfo_screenheight()
        #  Guardamos el largo y alto de la ventana
        wventana = 600
        hventana = 150
        #  Aplicamos la siguiente formula para calcular donde debería posicionarse
        pwidth = round(wtotal/2-wventana/2)
        pheight = round(htotal/2-hventana/2)
        #  Se lo aplicamos a la geometría de la ventana
        self.master.geometry(str(wventana)+"x"+str(hventana) +
                             "+"+str(pwidth)+"+"+str(pheight))
        self.master.grab_set()

        self.entry1 = tk.Entry(self.master)
        self.entry1.place(x=10, y=90, w=150)
        self.entry = tk.Entry(self.master)
        self.entry.place(x=10, y=35, w=580)
        self.Start_Time = Label(
            self.master, text="Start Time").place(x=190, y=70)
        self.Stop_Time = Label(
            self.master, text="Stop Time").place(x=320, y=70)
        self.Interval = Label(
            self.master, text="Interval (min)").place(x=450, y=70)
        # InterEntry = Combobox(self.master, width = 8).place(x = 450, y = 90)
        self.InterEntry = ttk.Combobox(self.master)
        self.InterEntry.place(x=450, y=90, w=45)
        self.intervals = ["0", "1", "2", "5", "12"]
        self.InterEntry['values'] = self.intervals
        self.InterEntry.set("0")
        self.StartT = ttk.Combobox(self.master)
        self.StartT.place(x=190, y=90, w=40)
        self.StartT['values'] = self.FillCombos(1)
        self.StartT.set("0")
        self.StartT1 = ttk.Combobox(self.master)
        self.StartT1.place(x=235, y=90, w=40)
        self.StartT1['values'] = self.FillCombos(2)
        self.StartT1.set("0")
        self.StopT = ttk.Combobox(self.master)
        self.StopT.place(x=320, y=90, w=40)
        self.StopT['values'] = self.FillCombos(1)
        self.StopT.set("23")
        self.StopT1 = ttk.Combobox(self.master)
        self.StopT1.place(x=365, y=90, w=40)
        self.StopT1['values'] = self.FillCombos(2)
        self.StopT1.set("59")
        self.btn = tk.Button(
            self.master, text="Choose File ...", command=self.Choose_File)
        self.btn.place(x=10, y=5)
        self.btn1 = tk.Button(self.master,
                              text="Process & Save File ...", command=self.Convert)
        self.btn1.place(x=10, y=60)
        # btn2 = tk.Button(text="Convert File...", command=Export_Excel)
        # btn2.place(x=10, y=85)
        self.entry.focus()
        self.master.mainloop()

    def TimeAdd(self, TimeTest):
        sec = int(self.AjustTime())
        formato = r"%Y-%m-%d %H:%M:%S"
        TimePlus1 = timedelta(seconds=sec)
        Time1 = datetime.datetime.strptime(TimeTest, formato)
        print(Time1)
        TimePlus = Time1 + TimePlus1
        print("FFFFFF ---- " + str(TimePlus))
        return TimePlus

    def CleanTXT(self):
        self.entry.delete('0', END)
        self.entry1.delete('0', END)

    def Choose_File(self):
        self.CleanTXT()
        try:
            filename = filedialog.askopenfilename(
                title="Examinar archivo", filetypes=[("PDF files", "*.pdf")])
            if filename == None or filename == '':
                return
            path = filename.split(".")
            f = len(path)
            if path[f-1] == "pdf":
                self.entry.insert(0, str(filename))
                nFile = os.path.split(filename)
                rFile = str(nFile[1]).replace(".pdf", "")
                self.entry1.insert(0, str(rFile))
            else:
                messagebox.showerror(
                    message="This file is not a PDF", title="Error extension file", parent=self.master)
        except:
            return

    def open_popup(self):
        top = Toplevel(self.root)
        top.geometry("200x250")
        top.title("Child Window")
        Label(top, text="Something was wrong!!! ", font=(
            'Mistral 18 bold')).place(x=150, y=80)

    def FillCombos(self, id):
        ret = []
        if id == 1:
            for i in range(0, 24):
                ret.append(str(i))
            return ret
        if id == 2:
            for i in range(0, 60):
                ret.append(str(i))
            return ret

    def AjustTime(self):

        nn = self.InterEntry.get()

        match str(nn):
            case "0":
                return "1"
            case "1":
                return "55"
            case "2":
                return "124"
            case "5":
                return "304"
            case "12":
                return "725"

    def Convert(self):

        if str(self.entry.get()) == "":
            messagebox.showerror(
                message="Please, choose file to convert!!", title="Error choosing file", parent=self.master)
            return

        NameFile = str(self.entry1.get())
        if NameFile == "":
            messagebox.showerror(
                message="Please, fill the name file box!!", title="Error name file", parent=self.master)
            return
        path = str(self.entry.get())
        pdf = PdfReader(path)
        number_of_pages = len(pdf.pages)
        print(f"Total Pages: {number_of_pages}")
        page_content = ""

        spl = []
        spl1 = []
        data = [[], [], [], []]

        n = 0
        l = 0
        nameFile = str(self.entry1.get())

        folder = os.path.dirname(path)
        filename = folder + "/" + nameFile + ".xlsx"

        try:
            workbook = xlsxwriter.Workbook(
                filename, {'default_date_format': 'dd/mm/yy'})
        except IOError:
            messagebox.showerror(
                message="Problems opening the Excel file.", title="Error opening file")

        for i in range(0, number_of_pages):  # number_of_pages
            page = pdf.pages[i]
            page_content += page.extract_text().strip()
            spl = page_content.split(" ")
            # print(spl)

        nn = 0
        row = 0
        ex = False
        con = False
        NoTime = False

        St = datetime.time(int(self.StartT.get()), int(self.StartT1.get()))
        St1 = datetime.time(int(self.StopT.get()), int(self.StopT1.get()))

        if int(self.StartT.get()) == 0 and int(self.StartT1.get()) == 0:
            NoTime = False
        else:
            NoTime = True

        if St > St1:
            messagebox.showerror(
                message="The StopTime is bigger than StartTime", title="Error Time")
            return
        NewTime = ""
        for txt in spl:
            n = n + 1
            txt = txt.replace("\n", "#")
            spl1 = txt.split("#")

            if n > 10:
                ab = len(spl1)
                sg = False
                ia = 0

                if ab == 1:

                    if ":" in spl1[ia]:
                        fix = []
                        fix = spl1[ia].split(":")
                        if len(fix[2]) > 2:
                            sg = True

                if ab > 1:
                    ia = 1
                    if "Page" in spl1[ia]:
                        print(spl1)
                        sg = False

                if ab > 1 or sg == True:
                    sg = False
                    if ":" in spl1[ia]:
                        txt = spl1[ia].strip()
                        txt1 = txt.split(":")
                        if len(txt1[0]) >= 2:
                            a = 8
                            b = 8
                        else:
                            a = 7
                            b = 7
                        print("1 -- Hora : " + txt[:a].strip())
                        time = txt[:a].strip()  # Time
                        txt = txt[b:].strip()
                        print("2 -- Melt Pressure : " + txt)
                        meltP = txt  # Melt Pressure
                        nn = nn + 1
                        con = True
                        Ts = time.split(":")
                        # print(Ts)
                        TimeSave = time
                        if NoTime == True:
                            if len(Ts) == 3:
                                time = "1900-01-01 " + time
                                # timeOld = TimeAdd(time.format(r"%d-%m-%Y %H:%M:%S"))
                                timeT = datetime.time(
                                    int(Ts[0]), int(Ts[1]), int(Ts[2]))
                                T = str(timeT).split(":")
                                formato = r"%Y-%m-%d %H:%M:%S"
                                timeT1 = datetime.datetime.strptime(
                                    "1900-01-01 " + str(T[0]) + ":" + str(T[1]) + ":" + str(T[2]), formato)

                                if NewTime == "":
                                    NewTime = self.TimeAdd(str(timeT1))
                                    messagebox.showinfo(message="The interval time selected is  " + str(
                                        self.InterEntry.get()), title="Interval time selected")

                                if timeT > St:
                                    ex = False
                                    if timeT < St1:
                                        ex = False

                                        if timeT1 > NewTime:
                                            NewTime = self.TimeAdd(str(timeT1))
                                            # messagebox.showerror(message="TimeT -- " + str(timeT) + "  TimeOld -- " + str(NewTime), title="New Time 2")
                                            ex = False
                                        else:
                                            ex = True
                                    else:
                                        ex = True
                                else:
                                    ex = True
                    r = 0
                    if con == True:

                        if "/" in spl1[ia]:
                            r = len(spl1)
                            if r > 2:

                                day = str(spl1[ia]) + str(spl1[ia+1])
                            else:
                                day = spl1[ia]  # Day
                            rpm = spl1[0]  # RPM
                            nn = nn + 1
                            nn = 5
                            print("4 --  Fecha : " + str(day))

                else:
                    if con == True:
                        l = l + 1

                        if l == 1:
                            # print (str(row) + " ---- " + str(l))
                            meltT = spl1[0]  # Melt Temp
                            nn = nn + 1

                            # print ("3 -- " + spl1[0])

                        if l == 2:
                            # print (str(row) + " ---- " + str(l))
                            mc = spl1[0]  # Motor Current
                            nn = nn + 1
                            # print("nn " + str(nn))
                            # print ("4 -- " + spl1[0])

                        if l == 3:
                            # print (str(row) + " ---- " + str(l))
                            feed = spl1[0]  # Feeding
                            nn = nn + 1
                            # print("nn " + str(nn))
                            print("5 -- Feeding : " + spl1[0])
                        l = 0

                if nn == 5:
                    nn = 0
                    con = False
                    row = row + 1
                    print(" row -- " + str(row))
                    if ex == False:
                        try:
                            data[0].append(str(day))
                            data[1].append(str(TimeSave))
                            data[2].append(int(meltP))
                            data[3].append(int(feed))
                        except:
                            print(("------ >" + str(data[0])))

                    else:
                        row = row - 1

        worksheet = workbook.add_worksheet(nameFile)
        bold = workbook.add_format({"bold": 1, "align": CENTER})
        bold1 = workbook.add_format({"bold": 0, "align": CENTER, 'border': 1})
        headings = ["DATE", "TIME", "MELT PRESS", "KG/HORA"]

        chart1 = workbook.add_chart({"type": "line"})
        chart1.add_series(
            {
                "name": "='" + str(nameFile).strip() + "'!$C$1",
                "categories": "='" + str(nameFile).strip() + "'!$B$2:$B$" + str(row+1),
                "values": "='" + str(nameFile).strip() + "'!$C$2:$C$" + str(row+1),
            }  # + str(row+1)
        )
        chart1.set_title({"name": "Results of batch analysis"})
        chart1.set_x_axis({"name": "Test number : " + str(nameFile).strip()})
        chart1.set_y_axis({"name": "Melt Pressure (bar)"})
        chart1.set_size({'width': 1150, 'height': 500})

        # Set an Excel chart style. Colors with white outline and shadow.
        chart1.set_style(10)

        worksheet.write_row("A1", headings, bold)
        worksheet.write_column("A2", data[0], bold1)
        worksheet.write_column("B2", data[1], bold1)
        worksheet.write_column("C2", data[2], bold1)
        worksheet.write_column("D2", data[3], bold1)
        worksheet.autofit()
        worksheet.insert_chart("E2", chart1, {"x_offset": 15, "y_offset": 10})
        # CheckIfOpen(filename)
        workbook.close()

        messagebox.showinfo(message="    Job DONE!!      ",
                            title="Status Conversion", parent=self.master)

        self.CleanTXT()
        msg = messagebox.askokcancel(
            "Excel created", "Do you want to open the Excel file?", icon="info", parent=self.master)
        if msg == True:
            webbrowser.open(str("file:")+filename)
        else:
            return
        print("Excel created ...")

    def CheckIfOpen(self, fileN):
        while True:   # repeat until the try statement succeeds
            try:
                myfile = open(str(fileN), "r+")  # or "a+", whatever you need
                break                             # exit the loop
            except IOError:
                messagebox.showinfo(
                    message="The file is opened! Please close Excel file before save the new Excel file. Press Enter to retry.", title="Error saving file", parent=self.master)
                # restart the loop


# def main():
    # root = tk.Tk()
    # PDftoexcel(root)


# if __name__ == '__main__':

   # main()
