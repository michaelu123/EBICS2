import locale
import os
import os.path
import logging
from tkinter import *
from tkinter.filedialog import askopenfilenames, askopenfilename
from tkinter.messagebox import *

import ebics


# not allowed to allow serviceaccount access to spreadsheet!?
# import gspread
# from oauth2client.service_account import ServiceAccountCredentials


class ButtonEntry(Frame):
    def __init__(self, master, buttontext, stringtext, w, cmd):
        super().__init__(master)
        self.btn = Button(self, text=buttontext, bg="red", bd=4, width=15, height=0, relief=RAISED, command=cmd)
        self.svar = StringVar()
        self.svar.set(stringtext)
        self.entry = Entry(self, textvariable=self.svar, bd=4,
                           width=w, borderwidth=2)
        self.grid_rowconfigure(0, weight=0)
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.btn.grid(row=0, column=0, sticky="w")
        self.entry.grid(row=0, column=1, sticky="we")

    def get(self):
        return self.svar.get()

    def set(self, s):
        return self.svar.set(s)


class LabelEntry(Frame):
    def __init__(self, master, labeltext, stringtext, w):
        super().__init__(master)
        self.label = Label(self, text=labeltext, bd=4, width=15, height=0, relief=RIDGE)
        self.svar = StringVar()
        self.svar.set(stringtext)
        self.entry = Entry(self, textvariable=self.svar, bd=4,
                           width=w, borderwidth=2)
        self.grid_rowconfigure(0, weight=0)
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.label.grid(row=0, column=0, sticky="w")
        self.entry.grid(row=0, column=1, sticky="we")

    def get(self):
        return self.svar.get()

    def set(self, s):
        return self.svar.set(s)


class LabelOM(Frame):
    def __init__(self, master, labeltext, options, initVal, **kwargs):
        super().__init__(master)
        self.options = options
        self.label = Label(self, text=labeltext, bd=4, width=15, relief=RIDGE)
        self.svar = StringVar()
        self.svar.set(initVal)
        self.optionMenu = OptionMenu(self, self.svar, *options, **kwargs)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.label.grid(row=0, column=0, sticky="w")
        self.optionMenu.grid(row=0, column=1, sticky="w")

    def get(self):
        return self.svar.get()

    def set(self, s):
        self.svar.set(s)


class MyApp(Frame):
    def __init__(self, master):
        super().__init__(master)
        w = 50
        self.inputFilesBE = ButtonEntry(master, "Eingabedateien", ebics.inputFilesDefault, w, self.inpFilesSetter)
        self.templateBE = ButtonEntry(master, "Template-Datei", ebics.templateFileDefault, w, self.templFileSetter)
        self.outputLE = LabelEntry(master, "Ausgabedatei", "ebics.xml", w)
        self.betragLE = LabelEntry(master, "Betrag", "100,00", w)
        self.zweckLE = LabelEntry(master, "Zweck", "ADFC Fahrradkurs", w)
        self.mandatLE = LabelEntry(master, "Mandat", "ADFC-M-RFS-2018", w)
        self.sepOM = LabelOM(master, "CSV Separator", ["Komma", "Semikolon"], "Komma")
        self.startBtn = Button(master, text="Start", bd=4, bg="red", width=15, command=self.starten)
        for x in range(1):
            Grid.columnconfigure(master, x, weight=1)
        for y in range(8):
            Grid.rowconfigure(master, y, weight=1)

        self.inputFilesBE.grid(row=0, column=0, sticky="we")
        self.templateBE.grid(row=1, column=0, sticky="we")
        self.outputLE.grid(row=2, column=0, sticky="we")
        self.betragLE.grid(row=3, column=0, sticky="we")
        self.zweckLE.grid(row=4, column=0, sticky="we")
        self.mandatLE.grid(row=5, column=0, sticky="we")
        self.sepOM.grid(row=6, column=0, sticky="w")
        self.startBtn.grid(row=7, column=0, sticky="w")

    def inpFilesSetter(self):
        x = askopenfilenames(title="CSV Dateien auswählen", defaultextension=".csv", filetypes=[("CSV", ".csv")])
        l = list(x)
        self.inputFilesBE.set(",".join(l))

    def templFileSetter(self):
        x = askopenfilename(title="Template Datei auswählen", defaultextension=".xml", filetypes=[("XML", ".xml")])
        self.templateBE.set(x)

    def starten(self):
        if self.inputFilesBE.get() == "":
            showerror("Fehler", "keine Eingabedatei")
            return
        if self.outputLE.get() == "":
            showerror("Fehler", "keine Ausgabedatei")
            return
        if self.zweckLE.get() == "":
            showerror("Fehler", "keine Ausgabedatei")
            return
        eb = ebics.Ebics(self.inputFilesBE.get(),
                   self.outputLE.get(),
                   self.betragLE.get(),
                   "," if self.sepOM.get() == "Komma" else ";",
                   self.zweckLE.get(),
                   self.mandatLE.get(),
                   self.templateBE.get())
        try:
            res = eb.createEbicsXml()
            stats = eb.getStatistics()
            msg = f"Anzahl Abbuchungsaufträge: {stats[0]}\nSchon bezahlt: {stats[1]}\nNoch abzubuchen: {stats[2]}"
            if res is None:
                showinfo("Nichts", msg)
                # showinfo("Nichts",
                #        "Anzahl Abbuchungsaufträge: {}".format(stats[0]),
                #        "Schon bezahlt: {}".format(stats[1]),
                #        "Noch abzubuchen: {}".format(stats[2]))
            else:
                showinfo("Erfolg", msg + f"\nAusgabe in Datei {self.outputLE.get()} erzeugt")
        except Exception as e:
            logging.exception("Fehler")
            showerror("Fehler", str(e))


#if __name__ == '__main__':
#    main()
# locale.setlocale(locale.LC_ALL, "de_DE")
locale.setlocale(locale.LC_TIME, "German")
root = Tk()
app = MyApp(root)
app.master.title("ADFC Lastschrift")
app.mainloop()
