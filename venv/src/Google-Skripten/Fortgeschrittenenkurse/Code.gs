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
//## Achtung: unten noch Fallunterscheidung für Kursgebühr!!
function kursTermine(code) {
  var text = ""
  if (code == "RFS_0F01") {
    text += '\nMittwoch, den 08. Juli 2020 von 17:00 bis ca. 20:00. \n';
    text += "(Ersatztermin bei schlechtem Wetter: Mittwoch, den 15. Juli, gleiche Uhrzeit.)\n";
    return text;
  }
  if (code == "RFS_0F02") {
    text += '\nSamstag, den 08. August 2020 von 14:00 bis ca. 17:00. \n';
    text += "(Ersatztermin bei schlechtem Wetter: Samstag, den 22. August, gleiche Uhrzeit.)\n";
    return text;
  }
  if (code == "RFS_0F03") {
    text += '\nFreitag, den 17. Juli 2020 von 16:00 bis 20:00.\n' ;
    text += "(Ersatztermin bei schlechtem Wetter: Freitag, den 24. Juli, gleiche Uhrzeit.)\n";
    return text;
  }
  if (code == "RFS_0F04") {
    text += '\nFreitag, den 14. August 2020 von 15:00 bis 19:00.\n' ;
    text += "(Ersatztermin bei schlechtem Wetter: Freitag, den 21. August, gleiche Uhrzeit.)\n";
    return text;
  }
  if (code == "RFS_0F05") {
    text += '\nSamstag, den 5. September 2020 von 15:00 bis 19:00.\n' ;
    text += "(Ersatztermin bei schlechtem Wetter: Samstag, den 12. September, gleiche Uhrzeit.)\n";
    return text;
  }
  text = "Interner Fehler: Für diesen Kurs ist kein Kursdatum im Skript hinterlegt";
  SpreadsheetApp.getUi().alert(text);
  return text;
}

//#########################################################
function heuteString() {
  return Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'YYYY-MM-dd hh:mm:ss');
}

//#########################################################
function attachmentFiles() {
  var thisFileId = SpreadsheetApp.getActive().getId();
  var thisFile = DriveApp.getFileById(thisFileId);
  var parent = thisFile.getParents().next();
  var grandPa = parent.getParents().next();
  var attachmentFolder = grandPa.getFoldersByName("Texte für Fahrsicherheitstrainings").next();
  var PDFs = attachmentFolder.getFilesByType(MimeType.PDF);
  var files = [];
  while (PDFs.hasNext()) {
    files.push(PDFs.next());
  }
  return files;
}

// Indices are 1-based
var mailIndex = 2; // B=E-Mail-Adresse
var herrFrauIndex = 3; // C=Anrede
var nameIndex = 5; // E=Name
var mitgliedIndex = 10; // J=ADFC-Mitgliedsnummer falls Mitglied
var ZahlungsartIndex = 11; // K=Überweisung
var bestätigungIndex = 16; // P=Bestätigung
var verifikationIndex = 18; // R=Verifikation
var anmeldebestIndex = 19; // S=Anmeldebestätigung


//#########################################################
function mailSchicken() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var anmeldeSheets = SpreadsheetApp.openById("1Qx1t4dPqbrZYQ8cck9kPv_P8RPQfZHAjxeWr2nJavkQ").getSheets(); // Textbausteine
  var anmeldeTexte = anmeldeSheets[0];
  var textColumn = 2;
  var subjectRow = 10;
  var bodyRow = 11;
  var subjectTemplate = anmeldeTexte.getRange(subjectRow, textColumn).getValue().toString();
  var bodyTemplate = anmeldeTexte.getRange(bodyRow,textColumn).getValue().toString();
  var row = sheet.getSelection().getCurrentCell().getRow();
  if (row < 2 || row > 15) {
    SpreadsheetApp.getUi().alert("Die ausgewählte Zeile ist ungültig, bitte zuerst Teilnehmerzeile selektieren");
    return;
  }
  var b = sheet.getRange(row, bestätigungIndex).getValue().toString();
  Logger.log("Best %s %s", b, b.indexOf("Ich habe"));
  if (sheet.getRange(row, bestätigungIndex).getValue().toString().indexOf("Ich habe") != 0) {
    SpreadsheetApp.getUi().alert("Hat Teilnahmebedingungen nicht zugestimmt");
    return;
  }
  var v = sheet.getRange(row, verifikationIndex).getValue().toString();
  Logger.log("Veri %s", v);
  if (sheet.getRange(row, verifikationIndex).getValue().toString() == "") {
    SpreadsheetApp.getUi().alert("Emailadresse nicht verifiziert");
    return;
  }

  // setting up mail
  var code = kursName(sheet.getName());
  var empfaenger = sheet.getRange(row, mailIndex).getValue();
  var subject = Utilities.formatString(subjectTemplate, code);
  // Anrede
  var anrede = anredeText(sheet.getRange(row, herrFrauIndex).getValue(), sheet.getRange(row, nameIndex).getValue());
  // Zahlungstext
  var mitgliedsnummer = sheet.getRange(row, mitgliedIndex).getValue().toString();
  var Kursgebühr = "";
  if (code == "RFS_0F03" || code == "RFS_0F04")
    Kursgebühr = mitgliedsnummer != "" ? "16" : "32";
  else
    Kursgebühr = mitgliedsnummer != "" ? "12" : "24";
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
 // update sheet
  sheet.getRange(row, anmeldebestIndex).setValue(heuteString());
}





//#########################################################
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Radfahrschule')
      .addItem('Anmeldebestätigung senden', 'mailSchicken')
      .addItem('Emailverif', 'verifEmail')
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
      var rrow = rvalues[bx];
      if (rrow[mailIndex-1] != "" && rrow[verifikationIndex - 1] == "") {
        var raddr = rrow[1];
        for (var ex in evalues) {
          ex = +ex
          var erow = evalues[ex];
          if (erow[1] != "Ja" || erow[2] == "")
            continue;
          var eaddr = erow[2];
          if (eaddr != raddr)
            continue;
          // Bestellungen[Verifiziert] = Email-Verif[Zeitstempel]
          Logger.log("bx=%s ex=%s", bx, ex);
          rfsSheet.getRange(bx + 2, verifikationIndex).setValue(erow[0]);
          rrow[verifikationIndex - 1] = erow[0];
          break;
        }
      }
    }
  }
}
