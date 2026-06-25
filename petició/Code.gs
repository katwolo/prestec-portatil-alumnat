function doGet() {
  var idDocument = "1jO1_Oh9lK_vKXR4l4Shh2ZD2-P-wVbfGF2pS1xUGbw8"; // <--- ENGANXA AQUÍ L'ID DEL TEU GOOGLE DOC

  var template = HtmlService.createTemplateFromFile('Index');
  template.idDocument = idDocument;

  return template.evaluate()
      .setTitle('Petició de Préstec de Portàtils')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function textCondicionsDoc(id) {
  try {
    var doc = DocumentApp.openById(id);
    var body = doc.getBody().getText();
    return body.replace(/\n/g, '<br>');
  } catch(e) {
    return "Error al carregar les condicions. Si us plau, contacta amb l'administrador del centre.";
  }
}

// Estructura de 'dades':
// {
//   nomTutor:   string,  (Coach)
//   emailTutor: string,  (Mail coach)
//   curs:       string,  (Curs)
//   classe:     string,  (Classe)
//   timestamp:  string (ISO),
//   alumnes: [
//     { nom: string, cognoms: string, emailAlum: string, preferencia: "Alta"|"Mitja"|"Baixa" },
//     ...
//   ]
// }
function guardarRegistre(dades) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var timestamp = new Date(dades.timestamp);

  dades.alumnes.forEach(function(alumne) {
    sheet.appendRow([
      timestamp,          // A: Registre
      alumne.nom,         // B: Nom
      alumne.cognoms,     // C: Cognoms
      dades.curs,         // D: Curs
      dades.classe,       // E: Classe
      alumne.emailAlum,   // F: Email Alum
      dades.nomTutor,     // G: Coach
      dades.emailTutor,   // H: Mail coach
      alumne.preferencia  // I: Preferència
    ]);
  });

  return "Sol·licitud registrada correctament. S'han afegit " + dades.alumnes.length + " alumne(s) a la llista.";
}
