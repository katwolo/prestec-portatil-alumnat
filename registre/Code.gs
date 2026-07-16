// ══════════════════════════════════════════════════════════════
//  Registres petició portàtils — Code.gs
// ══════════════════════════════════════════════════════════════

var ID_DOCUMENT_CONDICIONS = "1jO1_Oh9lK_vKXR4l4Shh2ZD2-P-wVbfGF2pS1xUGbw8"; // 🔴 ID del Google Doc de condicions
var ID_GESTOR = "15b0FVkvr7eVeeyjokrk2ely9sQmwLHyBZaQwqKEOtVM";             // 🔴 ID del Gestor de Préstec

// ──────────────────────────────────────────────────────────────
//  API per a GitHub Pages — respon JSONP (GET + ?callback=...)
//  Desplegament GAS: "Execute as: Me · Who has access: Anyone"
// ──────────────────────────────────────────────────────────────

function doGet(e) {
  var p  = e && e.parameter ? e.parameter : {};
  var cb = p.callback || '';
  var result;

  if (p.action === 'textCondicions') {
    try {
      var text = textCondicionsDoc(ID_DOCUMENT_CONDICIONS);
      result = { ok: true, text: text };
    } catch (ex) {
      result = { ok: false, msg: ex.message };
    }
  } else if (p.action === 'guardar') {
    try {
      var datos = JSON.parse(p.data || '{}');
      var msg = guardarRegistro(datos);
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

function guardarRegistro(datos) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.appendRow([
    new Date(),        // A: Registre (Timestamp)
    datos.nombre,      // B: Nom
    datos.apellidos,   // C: Cognoms
    datos.curso,       // D: Curs
    datos.clase,       // E: Classe
    datos.emailAlumno, // F: Email Alum
    datos.coach,       // G: Nom Coach
    datos.emailCoach   // H: Email Coach
  ]);
  try { sincronitzarRegistre_(); } catch(e) { Logger.log('Sync error: ' + e.message); }
  return "Sol·licitud registrada correctament. Les condicions de l'Institut de l'Esport de Barcelona han estat acceptades.";
}

// ──────────────────────────────────────────────────────────────
//  Menú
// ──────────────────────────────────────────────────────────────

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Gestió Portàtils')
    .addItem('Obrir aplicació', 'obrirEnllac')
    .addSeparator()
    .addItem('🔄 Sincronitzar registre de peticions', 'sincronitzarRegistre')
    .addToUi();
}

function obrirEnllac() {
  var url = "https://script.google.com/a/macros/ieb.cat/s/AKfycbxibKCzqxdi7Nz_cBwjpllf1EwXwHLcqdq0z1t4Be3QMfMRP2pbo_IeETXRTpsRiIiu8A/exec";
  var htmlOutput = HtmlService
    .createHtmlOutput(
      'Pots accedir directament fent clic al següent enllaç:<br><br>' +
      '<a href="' + url + '" target="_blank" style="padding: 10px 20px; background-color: #4CAF50; color: white; text-decoration: none; border-radius: 4px; font-family: sans-serif;">Anar a Gestió Portàtils</a>' +
      '<br><br><p style="font-size: 12px; color: gray;">(S\'obrirà en una pestanya nova)</p>'
    )
    .setWidth(350)
    .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Gestió Portàtils');
}

// ──────────────────────────────────────────────────────────────
//  Sincronització cap al Gestor
// ──────────────────────────────────────────────────────────────

function generateId_() {
  var chars = 'abcdefghijklmnopqrstuvwxyz0123456789';
  var id = '';
  for (var i = 0; i < 12; i++) {
    id += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return id;
}

function generateUniqueId_(shAlumnes) {
  var lastRow = shAlumnes.getLastRow();
  var existingIds = lastRow > 1
    ? shAlumnes.getRange(2, 1, lastRow - 1, 1).getValues().map(function(r) { return r[0]; })
    : [];
  var id;
  do { id = generateId_(); } while (existingIds.indexOf(id) !== -1);
  return id;
}

function sincronitzarRegistre() {
  var ui = SpreadsheetApp.getUi();
  try {
    var msg = sincronitzarRegistre_();
    ui.alert('Sincronització completada', msg, ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('❌ Error', 'S\'ha produït un error durant la sincronització:\n' + e.message, ui.ButtonSet.OK);
  }
}

function sincronitzarRegistre_() {
  var ssRegistre = SpreadsheetApp.getActiveSpreadsheet();
  var shRegistre = ssRegistre.getSheetByName('Full 1');
  if (!shRegistre) throw new Error('No s\'ha trobat la pestanya "Full 1".');

  var ssGestor  = SpreadsheetApp.openById(ID_GESTOR);
  var shAlumnes = ssGestor.getSheetByName('alumnes');
  if (!shAlumnes) throw new Error('No s\'ha trobat la pestanya "alumnes" al Gestor.');

  var lastRow = shRegistre.getLastRow();
  if (lastRow < 2) return 'ℹ️ No hi ha dades al full de registre.';

  var data = shRegistre.getRange(2, 1, lastRow - 1, 9).getValues();
  // A=0: Timestamp | B=1: Nom | C=2: Cognoms | D=3: Curs
  // E=4: Classe    | F=5: Email Alum | G=6: Coach | H=7: Mail coach | I=8: ✅ Copiat

  var copiats = 0;
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[1] && !row[2]) continue;
    if (row[8] === '✅')    continue;

    var nouId = generateUniqueId_(shAlumnes);
    shAlumnes.appendRow([
      nouId,        // A — id
      row[1] || '', // B — nom
      row[2] || '', // C — cog
      row[3] || '', // D — curs
      row[4] || '', // E — grup
      row[5] || '', // F — email
      row[6] || '', // G — tnom (Coach)
      row[7] || ''  // H — temail (Mail coach)
    ]);
    shRegistre.getRange(i + 2, 9).setValue('✅');
    copiats++;
  }

  return copiats > 0
    ? '✅ ' + copiats + ' alumne' + (copiats !== 1 ? 's' : '') + ' nou' + (copiats !== 1 ? 's' : '') + ' afegit' + (copiats !== 1 ? 's' : '') + ' correctament.'
    : 'ℹ️ No hi ha registres nous per sincronitzar.';
}
