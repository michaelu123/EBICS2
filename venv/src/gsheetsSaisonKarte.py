import gsheets

class GSheetSK(gsheets.GSheet):
    def __init__(self, stdBetrag, stdZweck):
        super().__init__("22", "Saisonkarte 2020")
        self.spreadSheetId = "1IsG9HpZlDU97Sf82LG-XkTPY1stI_xONH_pDP3x72BU"  # Saisonkarten-Bestellungen
        # Mit dem aktuellen credentials.json brauchen wir dafür eine externe Linkfreigabe für adfc-muc zum Bearbeiten
        self.spreadSheetName = "Saisonkarten-Bestellungen"

        self.ebicsnames = ebicsnames = ["Name des Kontoinhabers (kann leer bleiben falls gleich Mitgliedsname)", "IBAN-Kontonummer", "Betrag", "Zweck" ]
        self.ktoinh = ebicsnames[0]
        self.iban = ebicsnames[1]
        self.betrag = ebicsnames[2]
        self.zweck = ebicsnames[3]

        # Felder die wir überprüfen
        self.formnames = formnames = ["ADFC-Mitgliedsname", "ADFC-Mitgliedsnummer", "Zustimmung zur SEPA-Lastschrift"]
        self.mitgliedsname = formnames[0]
        self.mitgliedsnummer = formnames[1]
        self.zustimmung = formnames[2]

        # diese Felder fügen wir hinzu
        self.zusatzFelder = zusatzFelder = ["Verifiziert", "Gesendet", "Abgebucht", "Bezahlt", "Kommentar"]
        self.verifikation = zusatzFelder[0]
        self.gesendet = zusatzFelder[1]  # wird vom Skript der Tabelle ausgefüllt
        self.eingezogen = zusatzFelder[2]
        self.zahlungseingang = zusatzFelder[3] # händisch
        self.kommentar = zusatzFelder[4]

    @classmethod
    def getDefaults(self):
        return ("22", "ADFC Saisonkarte", "ADFC-M-SK-2020")

    def validSheetName(self, sname):
        return sname == "Bestellungen" or sname == "Email-Verifikation"

    def checkKtoinh(self, row):
        if row[self.ktoinh] == "":
            row[self.ktoinh] = row[self.mitgliedsname]
        return True

    def parseEmailVerif(self):
        emailVerifSheet = self.data.pop("Email-Verifikation", None)
        if emailVerifSheet is None:
            return
        headers = emailVerifSheet[0]
        if headers[0] != "Zeitstempel" or \
                headers[1] != "Haben Sie gerade eine Saisonkarte bestellt?" or \
                headers[2] != "Mit dieser Email-Adresse (bitte nicht ändern!) :":
            print("Arbeitsblatt Email-Verifikation hat falsche Header-Zeile", headers)
        for row in emailVerifSheet[1:]:
            if len(row) != 3:
                continue
            if row[1] == "Ja":
                self.emailAdresses[row[2]] = row[0]

