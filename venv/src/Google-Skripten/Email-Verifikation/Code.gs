function verifyEmaiAddress(e) {
  var emailTo = e.response.getRespondentEmail()
  Logger.log("emailTo=" + emailTo)
  var html = HtmlService.createTemplateFromFile("email.html")
  html.anrede = anrede(e)
  html.verifLink = "https://docs.google.com/forms/d/e/1FAIpQLSfxNZUwFdOKWXFbb1PDeEH4E3nOalf120gcgG3ieIwzMYRgIQ/viewform?usp=pp_url&entry.1730791681=Ja&entry.1561755994=" + encodeURIComponent(emailTo)
  var htmlText = html.evaluate().getContent()
  Logger.log("htmlText=" + htmlText)
  var subject = "Bestätigung Ihrer Email-Adresse"
  var textbody = "HTML only"
  var options = { htmlBody: htmlText, name:'Radfahrschule ADFC München e.V.', replyTo: "anmeldungen-radfahrschule@adfc-muenchen.de"}

  if (emailTo !== undefined){
    GmailApp.sendEmail(emailTo, subject, textbody, options);
  }
}

function anrede(e) {
  var res = "";
  var anrede = "";
  var vorname = "";
  var name = "";
  var formResponse = e.response;
  var itemResponses = formResponse.getItemResponses();
  for (var j = 0; j < itemResponses.length; j++) {
    var itemResponse = itemResponses[j];
    var q = itemResponse.getItem().getTitle().toString();
    var a = itemResponse.getResponse().toString();
    if (q == "Anrede") {
      if (a == "Herr")
        a = "Sehr geehrter Herr ";
      else
        a = "Sehr geehrte Frau ";
      anrede = a;
    }
    if (q == "Vorname")
      vorname = a + " ";
    if (q == "Name")
      name = a;
  }
  return anrede + vorname + name;

}
