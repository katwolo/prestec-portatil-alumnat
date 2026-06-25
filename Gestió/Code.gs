// ══════════════════════════════════════════════════════════════
//  Préstec Portàtils Alumnat — Google Apps Script · Code.gs
//  ─ Persistència: Google Sheets (pestanyes: portatils, alumnes, prestecs)
//  ─ Motor de correus: llegeix plantilles HTML de la pestanya "Plantilles"
//  ─ Registre d'enviaments: pestanya "Registre"
// ══════════════════════════════════════════════════════════════

var HEADERS = {
  portatils: ['id','marca','model','sace','estat','usr','pwd','obs'],
  alumnes:   ['id','nom','cog','curs','grup','email','tnom','temail'],
  prestecs:  ['id','aId','pId','dataE','dataD','curs','obs','obsD','estat']
};

var ID_PETICIONS = "1Pm06Uu3Y350NTZqNuWs4g6e6r4Zqaq5zQAOyrbD8lzU";
var ID_REGISTRE  = "1bAYP40kOYT7C006R3bUuWuIxtGg79XeOr55DvJ5t_eo";

// ══════════════════════════════════════════════════════════════
//  INTERFÍCIE D'USUARI
// ══════════════════════════════════════════════════════════════

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('💻 Préstec Portàtils')
    .addItem('Obrir aplicació', 'openApp')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('✉️ Enviaments')
        .addItem('Enviar ara (Petició Alumnat)', 'enviarCorreus')
    )
    .addSeparator()
    .addItem('🗑️ Netejar registre d\'enviaments', 'netejarRegistre')
    .addToUi();
}

function openApp() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
    .setWidth(1200)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Préstec Portàtils Alumnat');
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Préstec Portàtils Alumnat')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ══════════════════════════════════════════════════════════════
//  HELPERS DE PESTANYA
// ══════════════════════════════════════════════════════════════

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (HEADERS[name]) {
      sh.appendRow(HEADERS[name]);
      sh.setFrozenRows(1);
      var hRange = sh.getRange(1, 1, 1, HEADERS[name].length);
      hRange.setBackground('#143f7a').setFontColor('#ffffff').setFontWeight('bold');
    }
  }
  return sh;
}

function sheetToObjects(name) {
  var sh = getSheet(name);
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      obj[h] = (row[i] === undefined || row[i] === null) ? '' : String(row[i]);
    });
    return obj;
  });
}

// ══════════════════════════════════════════════════════════════
//  API DE DADES
// ══════════════════════════════════════════════════════════════

function loadDB() {
  var db = {
    portatils: sheetToObjects('portatils'),
    alumnes:   sheetToObjects('alumnes'),
    prestecs:  sheetToObjects('prestecs')
  };
  return JSON.stringify(db);
}

function saveDB(dbStr) {
  var db = JSON.parse(dbStr);
  var tables = ['portatils', 'alumnes', 'prestecs'];
  tables.forEach(function(table) {
    var sh = getSheet(table);
    var lastRow = sh.getLastRow();
    if (lastRow > 1) sh.deleteRows(2, lastRow - 1);
    db[table].forEach(function(rec) {
      sh.appendRow(HEADERS[table].map(function(h) { return rec[h] !== undefined ? rec[h] : ''; }));
    });
  });
}

function saveChanges(changesStr) {
  var changes = JSON.parse(changesStr);
  ['portatils', 'alumnes', 'prestecs'].forEach(function(table) {
    var recs = changes[table];
    if (!recs || !recs.length) return;
    recs.forEach(function(rec) {
      saveRecord(table, JSON.stringify(rec));
    });
  });
}

function saveRecord(table, recStr) {
  var rec     = JSON.parse(recStr);
  var sh      = getSheet(table);
  var headers = HEADERS[table];
  var data    = sh.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(rec.id)) {
      var row = headers.map(function(h) { return rec[h] !== undefined ? rec[h] : ''; });
      sh.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      return;
    }
  }
  var newRow = headers.map(function(h) { return rec[h] !== undefined ? rec[h] : ''; });
  sh.appendRow(newRow);
}

function saveRecords(table, recsStr) {
  var recs = JSON.parse(recsStr);
  recs.forEach(function(rec) { saveRecord(table, JSON.stringify(rec)); });
}

// ══════════════════════════════════════════════════════════════
//  MOTOR DE PLANTILLES
// ══════════════════════════════════════════════════════════════

function loadPlantilles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Plantilles');
  if (!sh) return JSON.stringify({ assumpte: '', plantilles: [] });

  var assumpte = String(sh.getRange('B1').getValue() || '');
  var data = sh.getDataRange().getValues();

  var plantilles = [];
  for (var i = 4; i < data.length; i++) {
    var row = data[i];
    if (!row[0] || String(row[0]).trim() === '') continue;
    plantilles.push({
      id:   String(row[0]).trim(),
      cos:  String(row[1] || ''),
      dest: String(row[2] || 'alumnat').trim().toLowerCase()
    });
  }
  return JSON.stringify({ assumpte: assumpte, plantilles: plantilles });
}

function renderPlaceholders_(txt, d) {
  if (!txt) return '';
  return txt
    .replace(/\{\{nombreAlumno\}\}/gi,   d.nombreAlumno   || '')
    .replace(/\{\{correoAlumno\}\}/gi,   d.correoAlumno   || '')
    .replace(/\{\{nombreTutor\}\}/gi,    d.nombreTutor    || '')
    .replace(/\{\{correoTutor\}\}/gi,    d.correoTutor    || '')
    .replace(/\{\{Dadesportatils\}\}/gi, d.dadesPortatils || '');
}

function buildDadesPortatilsByAluId_(alumneId) {
  try {
    var prestecs = sheetToObjects('prestecs');
    var prestec  = null;
    for (var i = 0; i < prestecs.length; i++) {
      if (prestecs[i].aId === alumneId && prestecs[i].estat === 'actiu') {
        prestec = prestecs[i]; break;
      }
    }
    if (!prestec) return '';

    var portatils = sheetToObjects('portatils');
    var pc = null;
    for (var j = 0; j < portatils.length; j++) {
      if (portatils[j].id === prestec.pId) { pc = portatils[j]; break; }
    }
    if (!pc) return '';

    return '<ul>' +
      '<li><b>Model:</b> '                + (pc.marca || '') + ' ' + (pc.model || '') + '</li>' +
      '<li><b>Identificador (SACE):</b> ' + (pc.sace  || '') + '</li>' +
      '<li><b>Usuari:</b> '               + (pc.usr   || '') + '</li>' +
      '<li><b>Password:</b> '             + (pc.pwd   || '') + '</li>' +
      '</ul>';
  } catch(e) {
    return '';
  }
}

// ══════════════════════════════════════════════════════════════
//  ENVIAMENT DIRECTE DES DE LA WEB APP
// ══════════════════════════════════════════════════════════════

function enviarCorreuDirecte(templateId, alumneId, destOverride) {
  try {
    var tplData   = JSON.parse(loadPlantilles());
    var plantilla = null;
    for (var i = 0; i < tplData.plantilles.length; i++) {
      if (tplData.plantilles[i].id === templateId) {
        plantilla = tplData.plantilles[i]; break;
      }
    }
    if (!plantilla) {
      return JSON.stringify({ ok: false, msg: 'Plantilla "' + templateId + '" no trobada a la pestanya Plantilles.' });
    }

    var alumnes = sheetToObjects('alumnes');
    var alumne  = null;
    for (var j = 0; j < alumnes.length; j++) {
      if (alumnes[j].id === alumneId) { alumne = alumnes[j]; break; }
    }
    if (!alumne) {
      return JSON.stringify({ ok: false, msg: 'Alumne/a no trobat/da a la base de dades.' });
    }

    var dadesPortatils = buildDadesPortatilsByAluId_(alumneId);

    var dic = {
      nombreAlumno:   alumne.nom    || '',
      correoAlumno:   alumne.email  || '',
      nombreTutor:    alumne.tnom   || '',
      correoTutor:    alumne.temail || '',
      dadesPortatils: dadesPortatils
    };

    var assumpte = renderPlaceholders_(tplData.assumpte || 'Correu — Préstec portàtil', dic);
    var cos      = renderPlaceholders_(plantilla.cos, dic);

    var dest = (destOverride || plantilla.dest || 'alumnat').toLowerCase();
    var destinataris = [];

    if (dest === 'alumnat' || dest === 'alumne') {
      if (alumne.email)  destinataris.push({ email: alumne.email,  nom: (alumne.nom || '') + ' ' + (alumne.cog || '') });
    } else if (dest === 'coach' || dest === 'tutor') {
      if (alumne.temail) destinataris.push({ email: alumne.temail, nom: alumne.tnom || 'Tutor/a' });
    } else {
      if (alumne.email)  destinataris.push({ email: alumne.email,  nom: (alumne.nom || '') + ' ' + (alumne.cog || '') });
      if (alumne.temail) destinataris.push({ email: alumne.temail, nom: alumne.tnom || 'Tutor/a' });
    }

    if (!destinataris.length) {
      return JSON.stringify({ ok: false, msg: 'No hi ha adreces de correu vàlides per al destinatari seleccionat.' });
    }

    var enviats = [];
    destinataris.forEach(function(d) {
      try {
        GmailApp.sendEmail(d.email, assumpte, '', {
          htmlBody: cos,
          name:     'Departament AFD'
        });
        registrarEnviament_(d.email, assumpte);
        enviats.push(d.nom + ' <' + d.email + '>');
      } catch(e) {
        throw new Error('Error enviant a ' + d.email + ': ' + e.message);
      }
    });

    return JSON.stringify({ ok: true, msg: 'Correu enviat a: ' + enviats.join(' i ') });

  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

// ══════════════════════════════════════════════════════════════
//  ENVIAMENT PER LOTS
// ══════════════════════════════════════════════════════════════

function enviarCorreus() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Petició Alumnat');

  if (!sh) {
    ui.alert('No s\'ha trobat la pestanya "Petició Alumnat".');
    return;
  }

  var tplData    = JSON.parse(loadPlantilles());
  var plantilles = tplData.plantilles;
  var assBase    = tplData.assumpte || 'Préstec portàtil';

  if (!plantilles.length) {
    ui.alert('No hi ha plantilles a la pestanya "Plantilles".');
    return;
  }

  var data    = sh.getDataRange().getValues();
  var alertes = [];
  var enviats = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[1] !== true) continue;

    var nomPlantilla = String(row[2]  || '').trim();
    var nomAlumne    = String(row[4]  || '').trim();
    var correoAlumne = String(row[5]  || '').trim();
    var nomTutor     = String(row[7]  || '').trim();
    var correoTutor  = String(row[8]  || '').trim();
    var adjuntsRaw   = String(row[9]  || '').trim();
    var sacId        = String(row[11] || '').trim();

    var plantilla = null;
    for (var p = 0; p < plantilles.length; p++) {
      if (plantilles[p].id === nomPlantilla) { plantilla = plantilles[p]; break; }
    }
    if (!plantilla) {
      alertes.push('Fila ' + (i + 1) + ': plantilla "' + nomPlantilla + '" no trobada.');
      sh.getRange(i + 1, 2).setValue(false);
      continue;
    }

    var destinataris = [];
    var dest = plantilla.dest;
    if (dest === 'alumnat' || dest === 'alumne') {
      if (correoAlumne) destinataris.push(correoAlumne);
    } else if (dest === 'coach' || dest === 'tutor') {
      if (correoTutor) destinataris.push(correoTutor);
    } else {
      if (correoAlumne) destinataris.push(correoAlumne);
      if (correoTutor)  destinataris.push(correoTutor);
    }

    if (!destinataris.length) {
      alertes.push('Fila ' + (i + 1) + ' (' + nomAlumne + '): no hi ha destinataris vàlids.');
      sh.getRange(i + 1, 2).setValue(false);
      continue;
    }

    var dadesPortatils = '';
    if (sacId) {
      var shP = ss.getSheetByName('portatils');
      var dP  = shP ? shP.getDataRange().getValues() : [];
      for (var k = 1; k < dP.length; k++) {
        if (String(dP[k][4]).trim() === sacId) {
          dadesPortatils = '<ul>' +
            '<li><b>Model:</b> '                + (dP[k][1] || '') + ' ' + (dP[k][2] || '') + '</li>' +
            '<li><b>Identificador (SACE):</b> ' + (dP[k][4] || '') + '</li>' +
            '<li><b>Usuari:</b> '               + (dP[k][6] || '') + '</li>' +
            '<li><b>Password:</b> '             + (dP[k][7] || '') + '</li>' +
            '</ul>';
          break;
        }
      }
    }

    var dic = {
      nombreAlumno:   nomAlumne,
      correoAlumno:   correoAlumne,
      nombreTutor:    nomTutor,
      correoTutor:    correoTutor,
      dadesPortatils: dadesPortatils
    };

    var assumpte = renderPlaceholders_(assBase,       dic);
    var cos      = renderPlaceholders_(plantilla.cos, dic);

    var blobs = [];
    if (adjuntsRaw) {
      adjuntsRaw.split(',').forEach(function(url) {
        url = url.trim();
        var match = url.match(/[-\w]{25,}/);
        if (match) {
          try {
            blobs.push(DriveApp.getFileById(match[0]).getBlob());
          } catch(e) {
            alertes.push('Fila ' + (i + 1) + ': adjunt no llegit "' + url + '".');
          }
        }
      });
    }

    destinataris.forEach(function(email) {
      try {
        var opts = { htmlBody: cos, name: 'Departament AFD' };
        if (blobs.length) opts.attachments = blobs;
        GmailApp.sendEmail(email, assumpte, '', opts);
        enviats++;
        registrarEnviament_(email, assumpte);
      } catch(e) {
        alertes.push('Fila ' + (i + 1) + ' → ' + email + ': ' + e.message);
      }
    });

    sh.getRange(i + 1, 2).setValue(false);
  }

  var msg = '✅ ' + enviats + ' correu' + (enviats !== 1 ? 's' : '') + ' enviat' + (enviats !== 1 ? 's' : '') + ' correctament.';
  if (alertes.length) msg += '\n\n⚠️ Advertències:\n' + alertes.join('\n');
  ui.alert(msg);
}

// ══════════════════════════════════════════════════════════════
//  REGISTRE D'ENVIAMENTS
// ══════════════════════════════════════════════════════════════

function registrarEnviament_(destinatari, assumpte) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Registre');
  if (!sh) {
    sh = ss.insertSheet('Registre');
    sh.appendRow(['Data', 'Hora', 'Destinatari', 'Assumpte']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 4).setBackground('#143f7a').setFontColor('#ffffff').setFontWeight('bold');
  }
  var ara  = new Date();
  var data = Utilities.formatDate(ara, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var hora = Utilities.formatDate(ara, Session.getScriptTimeZone(), 'HH:mm:ss');
  sh.appendRow([data, hora, destinatari, assumpte]);
}

function netejarRegistre() {
  var ui  = SpreadsheetApp.getUi();
  var res = ui.alert(
    'Netejar registre',
    'Estàs segur/a que vols esborrar tot l\'historial d\'enviaments? Aquesta acció no es pot desfer.',
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Registre');
  if (!sh) { ui.alert('No existeix la pestanya "Registre".'); return; }

  var lastRow = sh.getLastRow();
  if (lastRow > 1) sh.deleteRows(2, lastRow - 1);
  ui.alert('✅ Registre netejat correctament.');
}

// ══════════════════════════════════════════════════════════════
//  PETICIONS — lectura del Sheets de peticions dels tutors
// ══════════════════════════════════════════════════════════════

function loadPeticions() {
  try {
    var ss = SpreadsheetApp.openById(ID_PETICIONS);
    var sh = ss.getSheetByName('Petició inicial');
    if (!sh) return JSON.stringify({ ok: false, msg: 'No s\'ha trobat la pestanya "Petició inicial".' });
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return JSON.stringify({ ok: true, data: [] });
    var data = sh.getRange(2, 1, lastRow - 1, 9).getValues();
    var tz = Session.getScriptTimeZone();
    var result = data
      .filter(function(r) { return r[1] || r[2]; })
      .map(function(r) {
        return {
          timestamp:   r[0] ? Utilities.formatDate(new Date(r[0]), tz, 'dd/MM/yyyy HH:mm') : '',
          nom:         String(r[1] || ''),
          cognoms:     String(r[2] || ''),
          curs:        String(r[3] || ''),
          classe:      String(r[4] || ''),
          emailAlum:   String(r[5] || ''),
          coach:       String(r[6] || ''),
          emailCoach:  String(r[7] || ''),
          preferencia: String(r[8] || '')
        };
      });
    return JSON.stringify({ ok: true, data: result });
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}

// ══════════════════════════════════════════════════════════════
//  SYNC REGISTRE → GESTIÓ (des de la web app, botó manual)
// ══════════════════════════════════════════════════════════════

function genId_() {
  var chars = 'abcdefghijklmnopqrstuvwxyz0123456789';
  var id = '';
  for (var i = 0; i < 12; i++) id += chars.charAt(Math.floor(Math.random() * chars.length));
  return id;
}

function genUniqueId_(sh) {
  var lastRow = sh.getLastRow();
  var existing = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 1).getValues().map(function(r) { return r[0]; }) : [];
  var id;
  do { id = genId_(); } while (existing.indexOf(id) !== -1);
  return id;
}

function syncRegistreManual() {
  try {
    var ssReg     = SpreadsheetApp.openById(ID_REGISTRE);
    var shReg     = ssReg.getSheetByName('Full 1');
    if (!shReg) return JSON.stringify({ ok: false, msg: 'No s\'ha trobat la pestanya "Full 1" al Registre.' });

    var shAlumnes = getSheet('alumnes');
    var lastRow   = shReg.getLastRow();
    if (lastRow < 2) return JSON.stringify({ ok: true, copiats: 0 });

    var data = shReg.getRange(2, 1, lastRow - 1, 9).getValues();
    // A=0:Timestamp | B=1:Nom | C=2:Cognoms | D=3:Curs | E=4:Classe | F=5:Email | G=6:Coach | H=7:EmailCoach | I=8:✅
    var copiats = 0;

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[1] && !row[2]) continue;
      if (row[8] === '✅') continue;

      shAlumnes.appendRow([
        genUniqueId_(shAlumnes),
        row[1] || '',
        row[2] || '',
        row[3] || '',
        row[4] || '',
        row[5] || '',
        row[6] || '',
        row[7] || ''
      ]);
      shReg.getRange(i + 2, 9).setValue('✅');
      copiats++;
    }

    return JSON.stringify({ ok: true, copiats: copiats });
  } catch(e) {
    return JSON.stringify({ ok: false, msg: e.message });
  }
}
