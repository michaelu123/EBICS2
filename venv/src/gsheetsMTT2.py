from decimal import Decimal
from tkinter.messagebox import *
from tkinter.filedialog import askopenfilename
from gsheets import is_iban
import openpyxl

class GSheetMTT2:
    def __init__(self, stdBetrag, stdZweck):
        # diese Felder brauchen wir für den Einzug
        self.ebicsnames = ebicsnames = ["Kontoinhaber", "IBAN", "Saldo EUR", "Buchungsnummer"]
        self.ktoinh = ebicsnames[0]
        self.iban = ebicsnames[1]
        self.betrag = ebicsnames[2]
        self.zweck = ebicsnames[3]
        self.nr_einzuziehen = 0



    @classmethod
    def getDefaults(self):
        return ("je nach Reise", "ADFC Mehrtagestouren", "ADFC-M-MTT-2022")

    def fillEingezogen(self):
        pass

    def getStatistics(self):
        return ( -1, -1, -1, -1, self.nr_einzuziehen)

    def getEntries(self):
        path = askopenfilename(title="Excel-Datei (edoobox) auswählen", defaultextension=".xlsx", filetypes=[("XLSX", ".xlsx")])
        wb = openpyxl.load_workbook(filename=path)
        sheetnames = wb.get_sheet_names()
        if not "Export" in sheetnames:
            showerror("Ungültige Datei", "Kein Tabellenblatt 'Export' in Excel-Datei")
            return
        ws = wb.get_sheet_by_name("Export")
        headerMap = {}
        entries = []
        for (i, row) in enumerate(ws.rows):
            if i == 0:
                row0 = row
                for (j, v) in enumerate(row0):
                    headerMap[v.value] = j
                print("hdr", headerMap)
                if headerMap.get(self.zweck) is None or \
                    headerMap.get(self.zweck) is None \
                    or headerMap.get(self.zweck) is None \
                    or headerMap.get(self.zweck) is None:
                    showerror("Ungültige Datei", "Tabellenblatt 'Export' hat ungültige Headerzeile")
                    return
                continue
            buchung = {x: row[headerMap[x]].value for x in self.ebicsnames}
            if buchung[self.betrag] is None or \
                    buchung[self.ktoinh] is None or \
                    buchung[self.iban] is None or \
                    buchung[self.zweck] is None:
                continue
            if buchung[self.ktoinh] == "" or \
                    buchung[self.iban] == "" or \
                    buchung[self.zweck] == "":
                continue # skip empty rows
            iban = buchung[self.iban]
            if not is_iban(iban):
                showerror("Falsche IBAN", "IBAN " + iban + " von " + buchung[self.ktoinh] + " ist ungültig")
                continue
            betrag = buchung[self.betrag]
            if betrag >= 0:
                showerror("Betrag ungültig", "Betrag " + str(betrag) + " ist nicht negativ")
                continue
            buchung[self.betrag] = -betrag
            buchung[self.zweck] = "ADFC Mehrtagestour mit Buchungsnummer " + str(buchung[self.zweck])
            entries.append(buchung)
            self.nr_einzuziehen += 1
        return entries
