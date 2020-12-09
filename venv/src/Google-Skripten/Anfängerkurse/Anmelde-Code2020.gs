//#########################################################
function anredeText(herrFrau, name) {
  if (herrFrau === 'Herr') {
    return 'Sehr geehrter Herr ' + name;
  } else {
    return 'Sehr geehrte Frau ' + name;
  }
}

//#########################################################
function kursName(fileName) {
  var nameString = fileName.toString(fileName);
  var codeName = "";
  codeName = nameString.substring(0,8);
  return codeName;
}

//#########################################################
function kursTermine(code) {
  var text = ""
  if (code == "RFS_0A01") {
    text += 'Samstag, den 18. April 2020\n';
    text += 'Sonntag, den 19. April 2020\n';
    text += 'Samstag, den 25. April 2020\n';
    text += 'Sonntag, den 26. April 2020\n';
    text += 'jeweils von 10:00 bis 13:00\n';
    return text;
  }
  if (code == "RFS_0A02") {
    text += 'Mittwoch, den 15. Juli 2020\n';
    text += 'Donnerstag, den 16. Juli 2020\n';
    text += 'Mittwoch, den 22. Juli 2020\n';
    text += 'Donnerstag, den 23. Juli 2020\n';
    text += 'Donnerstag, den 30. Juli 2020\n';
    text += 'Freitag, den 31. Juli 2020\n';
    text += 'jeweils abends von 19:00 Uhr bis 21:00 Uhr\n';
    return text;
  }
  if (code == "RFS_0A03") {
    text += 'Montag, den 22. Juni 2020\n';
    text += 'Dienstag, den 23. Juni 2020\n';
    text += 'Montag, den 29. Juni 2020\n';
    text += 'Dienstag, den 30. Juni 2020\n';
    text += 'Montag, den 6. Juli 2020\n';
    text += 'Dienstag, den 7. Juli 2020\n';
    text += 'Montag, den 13. Juli 2020\n';
    text += 'Dienstag, den 14. Juli 2020\n';
    text += 'jeweils abends von 17:00 Uhr bis 18:30 Uhr\n';
    return text;
  }
  if (code == "RFS_0A04") {
    text += 'Mittwoch, den 29. Juli 2020\n';
    text += 'Mittwoch, den 5. August 2020\n';
    text += 'Mittwoch, den 12. August 2020\n';
    text += 'Mittwoch, den 19. August 2020\n';
    text += 'Mittwoch, den 26. August 2020\n';
    text += 'Mittwoch, den 2. September 2020\n';
    text += 'jeweils abends von 18:30 Uhr bis 20:30 Uhr\n';
    return text;
  }

  if (code == "RFS_0A05") {
    text += 'Dienstag, den 16. Juni 2020\n';
    text += 'Freitag, den 19. Juni 2020\n';
    text += 'Dienstag, den 23. Juni 2020\n';
    text += 'Freitag, den 26. Juni 2020\n';
    text += 'Dienstag, den 30. Juni 2020\n';
    text += 'Freitag, den 3. Juli 2020\n';
    text += 'jeweils abends von 19:00 Uhr bis 21:00 Uhr\n';
    return text;
  }

  if (code == "RFS_0A06") {
    text += 'Samstag, den 22. August 2020\n';
    text += 'Sonntag, den 23. August 2020\n';
    text += 'Samstag, den 29. August 2020\n';
    text += 'Sonntag, den 30. August 2020\n';
    text += 'jeweils von 15:00 Uhr bis 18:00 Uhr\n';
    return text;
  }

  if (code == "RFS_0A07") {
    text += 'Samstag, den 4. Juli 2020\n';
    text += 'Sonntag, den 5. Juli 2020\n';
    text += 'Samstag, den 11. Juli 2020\n';
    text += 'Sonntag, den 12. Juli 2020\n';
    text += 'jeweils von 15:00 Uhr bis 18:00 Uhr\n';
    return text;
  }

  if (code == "RFS_0A08") {
    text += 'Samstag, den 9. Mai 2020\n';
    text += 'Sonntag, den 10. Mai 2020\n';
    text += 'Samstag, den 16. Mai 2020\n';
    text += 'Sonntag, den 17. Mai 2020\n';
    text += 'jeweils von 10:00 Uhr bis 13:00 Uhr\n';
    return text;
  }

if (code == "RFS_0A09") {
    text += 'Montag, den 28. September 2020\n';
    text += 'Donnerstag, den 01. Oktober 2020\n';
    text += 'Montag, den 05. Oktober 2020\n';
    text += 'Donnerstag, den 08. Oktober 2020\n';
    text += 'Montag, den 12. Oktober 2020\n';
    text += 'Donnerstag, den 15. Oktober 2020\n';
    text += 'jeweils abends von 17:00 Uhr bis 19:00 Uhr\n';
    return text;
  }

  if (code == "RFS_0A10") {
    text += 'Dienstag, den 7. Juli 2020\n';
    text += 'Freitag, den 10. Juli 2020\n';
    text += 'Dienstag, den 14. Juli 2020\n';
    text += 'Freitag, den 17. Juli 2020\n';
    text += 'Dienstag, den 21. Juli 2020\n';
    text += 'Freitag, den 24. Juli 2020\n';
    text += 'jeweils von 19:00 Uhr bis 21:00 Uhr\n';
    return text;
  }


  text = "Interner Fehler: Für diesen Kurs ist kein Kursdatum im Skript hinterlegt";
  SpreadsheetApp.getUi().alert(text);
  return text;
}

//#########################################################
function heuteString() {
  return Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'YYYY-MM-dd HH:mm:ss');
}

//#########################################################
function attachmentFiles() {
  var thisFileId = SpreadsheetApp.getActive().getId();
  var thisFile = DriveApp.getFileById(thisFileId);
  var parent = thisFile.getParents().next();
  var grandPa = parent.getParents().next();
  var attachmentFolder = grandPa.getFoldersByName("Anhänge für Anmelde-Bestätigung").next();
  var PDFs = attachmentFolder.getFilesByType(MimeType.PDF);
  var files = [];
  while (PDFs.hasNext()) {
    files.push(PDFs.next());
  }
  return files;
}

var Kursgebühr = '100,00';

// Indices are 1-based
var mailIndex = 2; // B=E-Mail-Adresse
var herrFrauIndex = 3; // C=Anrede
var nameIndex = 5; // E=Name
var ZahlungsartIndex = 10; // J=Überweisung
var bestätigungIndex = 15; // O=Bestätigung
var verifikationIndex = 17; // Q=Verifikation
var anmeldebestIndex = 18; // R=Anmeldebestätigung


//#########################################################
function mailSchicken() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var anmeldeSheets = SpreadsheetApp.openById("1Qx1t4dPqbrZYQ8cck9kPv_P8RPQfZHAjxeWr2nJavkQ").getSheets(); // Textbausteine
  var anmeldeTexte = anmeldeSheets[0];
  var textColumn = 2;
  var subjectRow = 3;
  var bodyRow = 4;
  var subjectTemplate = anmeldeTexte.getRange(subjectRow, textColumn).getValue().toString();
  var bodyTemplate = anmeldeTexte.getRange(bodyRow,textColumn).getValue().toString();
  var row = sheet.getSelection().getCurrentCell().getRow();
  if (row < 2 || row > 15) {
    SpreadsheetApp.getUi().alert("Die ausgewählte Zeile ist ungültig, bitte zuerst Teilnehmerzeile selektieren");
    return;
  }
  if (sheet.getRange(row, bestätigungIndex).getValue().toString().indexOf("Ich habe") != 0) {
    SpreadsheetApp.getUi().alert("Hat Teilnahmebedingungen nicht zugestimmt");
    return;
  }
  // update sheet
  sheet.getRange(row, anmeldebestIndex).setValue(heuteString());

  // setting up mail
  var code = kursName(sheet.getName());
  var empfaenger = sheet.getRange(row, mailIndex).getValue();
  var subject = Utilities.formatString(subjectTemplate, code);
  // Anrede
  var anrede = anredeText(sheet.getRange(row, herrFrauIndex).getValue(), sheet.getRange(row, nameIndex).getValue());
  // Zahlungstext
  var zahlungsText = '';
  var zahlungsart = sheet.getRange(row, ZahlungsartIndex).getValue().toString();
  if (zahlungsart === 'Überweisung') {
    var überweisungsRow = 5;
    var überweisungsTemplate = anmeldeTexte.getRange(überweisungsRow, textColumn).getValue().toString();
    zahlungsText = Utilities.formatString(überweisungsTemplate, Kursgebühr, code);
  } else {
    var sepaRow = 6;
    var sepaTemplate = anmeldeTexte.getRange(sepaRow, textColumn).getValue().toString();
    zahlungsText = Utilities.formatString(sepaTemplate, Kursgebühr);
  }
  // Zusammensetzung Body
  var body = Utilities.formatString(bodyTemplate, anrede, code, kursTermine(code), zahlungsText);

  // setting up mail
  GmailApp.sendEmail(empfaenger, subject, body, {
    name: 'Radfahrschule ADFC München e.V.',
    replyTo: 'radfahrschule@adfc-muenchen.de',
    attachments: attachmentFiles()
});
}

//#########################################################
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Radfahrschule')
      .addItem('Anmeldebestätigung senden', 'mailSchicken')
//    .addItem('Emailverif', 'verifEmail')
      .addToUi();
}

function verifEmail() {
  var ssheet = SpreadsheetApp.getActiveSpreadsheet();
  var evSheet = ssheet.getSheetByName('Email-Verifikation');
  var evalues = evSheet.getSheetValues(2, 1, evSheet.getLastRow()-1, evSheet.getLastColumn()); // Mit dieser Email-Adresse

  sheets = ssheet.getSheets();
  for (rx in sheets) {
    var rfsSheet = sheets[rx];
    Logger.log("rfsSheet %s %s", rfsSheet, rfsSheet.getName());
    if (rfsSheet.getName().indexOf("RFS") != 0)
      continue;
    var numRows = rfsSheet.getLastRow();
    if (numRows <= 1)
      continue;
    var rvalues = rfsSheet.getSheetValues(2, 1, numRows - 1, rfsSheet.getLastColumn())
    Logger.log("rvalues %s", rvalues)

    for (var bx in rvalues) {
      bx = +bx // confusingly, bx is initially a string, and is interpreted as A1Notation in sheet.getRange(bx) !
      rrow = rvalues[bx];
      if (rrow[mailIndex-1] != "" && rrow[verifikationIndex - 1] == "") {
        raddr = rrow[1];
        for (var ex in evalues) {
          erow = evalues[ex];
          if (erow[1] != "Ja" || erow[2] == "")
            continue;
          eaddr = erow[2];
          if (eaddr != raddr)
            continue;
          // Bestellungen[Verifiziert] = Email-Verif[Zeitstempel]
          rfsSheet.getRange(bx + 2, verifikationIndex).setValue(erow[0]);
          rrow[verifikationIndex - 1] = erow[0];
          break;
        }
      }
    }


  }
}
