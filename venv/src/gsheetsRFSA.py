import gsheets

class GSheetRFSA(gsheets.GSheet):
    def __init__(self, stdBetrag, stdZweck):
        super().__init__(stdBetrag, stdZweck)
        self.spreadSheetId = "1xRwSYtnmB4Y3_2f8ZPHxLuzMy7WuuZIW8jOY0nsIzN8"  # RFS_0AXX Backend
        self.spreadSheetName = "RFS_0AXX Backend" \

        # diese Felder brauchen wir für den Einzug
        self.ebicsnames = ebicsnames = ["Lastschrift: Name des Kontoinhabers", "Lastschrift: IBAN-Kontonummer", "Betrag", "Zweck"]
        self.ktoinh = ebicsnames[0]
        self.iban = ebicsnames[1]
        self.betrag = ebicsnames[2]
        self.zweck = ebicsnames[3]

        # Felder die wir überprüfen
        self.formnames = formnames = ["Vorname", "Name", "Zustimmung zur SEPA-Lastschrift", "Bestätigung"]
        self.vorname = formnames[0]
        self.name = formnames[1]
        self.zustimmung = formnames[2]
        self.bestätigung = formnames[3]  # Bestätigung der Teilnahmebedingungen

        # diese Felder fügen wir hinzu
        self.zusatzFelder = zusatzFelder = ["Verifikation", "Anmeldebestätigung", "Eingezogen", "Zahlungseingang", "Kommentar"]
        self.verifikation = zusatzFelder[0]
        self.anmeldebest = zusatzFelder[1]  # wird vom Skript Radfahrschule/Anmeldebestätigung senden ausgefüllt
        self.eingezogen = zusatzFelder[2]
        self.zahlungseingang = zusatzFelder[3]  # händisch
        self.kommentar = zusatzFelder[4]

    @classmethod
    def getDefaults(self):
        return ("100", "ADFC Radfahrschule", "ADFC-M-RFS-2020")

    def validSheetName(self, sname):
        return sname.startswith("RFS") or sname == "Email-Verifikation"

    def checkKtoinh(self, row):
        inh = row[self.ktoinh]
        if len(inh) < 5 or inh.startswith("dto") or inh.startswith("ditto"):
            row[self.ktoinh] = row[self.vorname] + " " + row[self.name]
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

