import gsheets
from decimal import Decimal

class GSheetRFSF(gsheets.GSheet):
    def __init__(self, stdBetrag, stdZweck):
        super().__init__(stdBetrag, stdZweck)
        self.spreadSheetId = "19w4WEvKSZBGgEkVeYNkPQDNGOXsQDhJcZ-TQkOgy4ac"  # RFS_0FXX Backend
        self.spreadSheetName = "RFS_0FXX Backend" \
                               ""
        # diese Felder brauchen wir für den Einzug
        self.ebicsnames = ebicsnames = ["Lastschrift: Name des Kontoinhabers", "Lastschrift: IBAN-Kontonummer", "Betrag", "Zweck"]
        self.ktoinh = ebicsnames[0]
        self.iban = ebicsnames[1]
        self.betrag = ebicsnames[2]
        self.zweck = ebicsnames[3]

        # Felder die wir überprüfen
        self.formnames = formnames = ["Vorname", "Name", "ADFC-Mitgliedsnummer falls Mitglied", "Zustimmung zur SEPA-Lastschrift", "Bestätigung"]
        self.vorname = formnames[0]
        self.name = formnames[1]
        self.mitglied = formnames[2]
        self.zustimmung = formnames[3]
        self.bestätigung = formnames[4]  # Bestätigung der Teilnahmebedingungen

        # diese Felder fügen wir hinzu
        self.zusatzFelder = zusatzFelder = ["Verifikation", "Anmeldebestätigung", "Eingezogen", "Zahlungseingang", "Kommentar"]
        self.verifikation = zusatzFelder[0]
        self.anmeldebest = zusatzFelder[1]  # wird vom Skript Radfahrschule/Anmeldebestätigung senden ausgefüllt
        self.eingezogen = zusatzFelder[2]
        self.zahlungseingang = zusatzFelder[3]  # händisch
        self.kommentar = zusatzFelder[4]

    @classmethod
    def getDefaults(self):
        return ("12/24", "ADFC Radfahrschule", "ADFC-M-RFS-2020")

    def validSheetName(self, sname):
        return sname.startswith("RFS") or sname == "Email-Verifikation"

    def checkKtoinh(self, row):
        inh = row[self.ktoinh]
        if len(inh) < 5 or inh.startswith("dto") or inh.startswith("ditto"):
            row[self.ktoinh] = row[self.vorname] + " " + row[self.name]
        return True

    def checkBetrag(self, row):
        mitglied = row[self.mitglied] != ""
        if not self.betrag in row:
            row[self.betrag] = "12" if mitglied else "24"
        row[self.betrag] = Decimal(row[self.betrag].replace(',', '.'))  # 3,14 -> 3.14
        return True

    def parseEmailVerif(self):
        emailVerifSheet = self.data.pop("Email-Verifikation", None)
        if emailVerifSheet is None:
            return
        headers = emailVerifSheet[0]
        if headers[0] != "Zeitstempel" or \
                headers[1] != "Haben Sie sich gerade für einen Radfahrkurs angemeldet?" or \
                headers[2] != "Mit dieser Email-Adresse (bitte nicht ändern!) :":
            print("Arbeitsblatt Email-Verifikation hat falsche Header-Zeile", headers)
        for row in emailVerifSheet[1:]:
            if len(row) != 3:
                continue
            if row[1] == "Ja":
                self.emailAdresses[row[2]] = row[0]

