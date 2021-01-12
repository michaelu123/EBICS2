import gsheets
from decimal import Decimal


class GSheetTK(gsheets.GSheet):
    def __init__(self, stdBetrag, stdZweck):
        super().__init__("", stdZweck)
        self.spreadSheetId = "1jm8GL-Xblyh7vORDvWljbBgz0hbFXArrWpl96WCZVGU"  # Backend-Technikkurse
        self.spreadSheetName = "Backend-Technikkurse"

        # diese Felder brauchen wir für den Einzug
        self.ebicsnames = ebicsnames = ["Name des Kontoinhabers", "IBAN-Kontonummer", "Betrag", "Zweck"]
        self.ktoinh = ebicsnames[0]
        self.iban = ebicsnames[1]
        self.betrag = ebicsnames[2]
        self.zweck = ebicsnames[3]

        # Felder die wir überprüfen
        self.formnames = formnames = ["Vorname", "Name", "ADFC-Mitgliedsnummer", "Zustimmung zur SEPA-Lastschrift", "Bestätigung"]
        self.vorname = formnames[0]
        self.name = formnames[1]
        self.mitglied = formnames[2]
        self.zustimmung = formnames[3]
        self.bestätigung = formnames[4]  # Bestätigung der Teilnahmebedingungen

        # diese Felder fügen wir hinzu
        self.zusatzFelder = zusatzFelder = ["Verifikation", "Anmeldebestätigung", "Eingezogen", "Zahlungseingang",
                                            "Kommentar"]
        self.verifikation = zusatzFelder[0]
        self.anmeldebest = zusatzFelder[1]  # wird vom Skript Radfahrschule/Anmeldebestätigung senden ausgefüllt
        self.eingezogen = zusatzFelder[2]
        self.zahlungseingang = zusatzFelder[3]  # händisch
        self.kommentar = zusatzFelder[4]

    @classmethod
    def getDefaults(self):
        return ("10/15", "ADFC Technikkurse", "ADFC-M-TK-2021")

    def validSheetName(self, sname):
        return sname.startswith("Buchungen") or sname == "Email-Verifikation"

    def checkKtoinh(self, row):
        inh = row[self.ktoinh]
        if len(inh) < 5 or inh.startswith("dto") or inh.startswith("ditto"):
            row[self.ktoinh] = row[self.vorname] + " " + row[self.name]
        return True

    def checkBetrag(self, row):
        mitglied = row[self.mitglied] != ""
        if not self.betrag in row:
            row[self.betrag] = "10" if mitglied else "15"
        row[self.betrag] = Decimal(row[self.betrag].replace(',', '.'))  # 3,14 -> 3.14
        return True

    def parseEmailVerif(self):
        emailVerifSheet = self.data.pop("Email-Verifikation", None)
        if emailVerifSheet is None:
            return
        headers = emailVerifSheet[0]
        if headers[0] != "Zeitstempel" or \
                headers[1] != "Haben Sie sich gerade für einen Technikkurs angemeldet?" or \
                headers[2] != "Mit dieser Email-Adresse (bitte nicht ändern!) :":
            print("Arbeitsblatt Email-Verifikation hat falsche Header-Zeile", headers)
        for row in emailVerifSheet[1:]:
            if len(row) != 3:
                continue
            if row[1] == "Ja":
                self.emailAdresses[row[2]] = row[0]
