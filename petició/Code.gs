// ─────────────────────────────────────────────────────────────
//  API per a GitHub Pages — respon JSONP (GET + ?callback=...)
//  Desplegament GAS: "Execute as: Me · Who has access: Anyone"
// ─────────────────────────────────────────────────────────────
function doGet(e) {
  var p  = e && e.parameter ? e.parameter : {};
  var cb = p.callback || '';
  var result;

  if (p.action === 'guardar') {
    try {
      var dades = JSON.parse(p.data || '{}');
      var msg = guardarRegistre(dades);
      result = { ok: true, msg: msg };
    } catch (ex) {
      result = { ok: false, msg: ex.message };
    }
  } else {
    result = { ok: false, msg: 'Acció no reconeguda.' };
  }

  var json = JSON.stringify(result);
  return ContentService
    .createTextOutput(cb ? cb + '(' + json + ');' : json)
    .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
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
