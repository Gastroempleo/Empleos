var SPREADSHEET_ID = 'https://docs.google.com/spreadsheets/d/1fQKRt4236PBctr0Z2m2wv7ZcICJD7_SkndsWNBORh5U/edit?gid=0#gid=0'; // ← Lo cambiás después

function doGet(e) {
  var page = e.parameter.page;
  if (page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('Admin')
      .setTitle('Panel RRHH')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else if (page === 'dashboard') {
    return HtmlService.createHtmlOutputFromFile('Dashboard')
      .setTitle('Dashboard')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Postulaciones')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function generarID(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return 'POST-00001';
  var lastId = sheet.getRange(last, 13).getValue();
  if (!lastId) return 'POST-00001';
  var num = parseInt(lastId.toString().split('-')[1]) + 1;
  return 'POST-' + ('00000' + num).slice(-5);
}

function validarEmail(e) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e); }
function validarTelefono(t) { return /^[\d\s\-\+\(\)]{7,20}$/.test(t); }

function getMasterSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Base General');
  if (!sheet) {
    sheet = ss.insertSheet('Base General');
    sheet.appendRow([
      'Fecha','Nombre','Email','Teléfono','Barrio','Puesto',
      'Años Exp','Disponibilidad','CV','Nombre CV','Link CV',
      'Experiencia Manual','ID','Estado'
    ]);
  }
  return sheet;
}

function checkDuplicate(email, phone) {
  var sheet = getMasterSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] === email || data[i][3] === phone) {
      return { exists: true, row: i+1 };
    }
  }
  return { exists: false };
}

function uploadCV(base64Data, fileName, mime, nombre, puesto, fecha) {
  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mime, fileName);
  var root = DriveApp.getFoldersByName('CV Postulaciones');
  var main = root.hasNext() ? root.next() : DriveApp.createFolder('CV Postulaciones');
  var subName = puesto;
  var subs = main.getFoldersByName(subName);
  var dest = subs.hasNext() ? subs.next() : main.createFolder(subName);
  var archivo = dest.createFile(blob);
  return { id: archivo.getId(), url: archivo.getUrl(), name: archivo.getName() };
}

function procesarPostulacion(form, archivoBase64, fileName, fileMime) {
  var sheet = getMasterSheet();
  var id = generarID(sheet);
  var fecha = new Date().toISOString().slice(0,10);
  var cvInfo = { id: null, url: null, name: null };
  
  if (archivoBase64) {
    cvInfo = uploadCV(archivoBase64, fileName, fileMime, form.nombre, form.puesto, fecha);
  }
  
  sheet.appendRow([
    fecha,
    form.nombre,
    form.email,
    form.telefono,
    form.barrio,
    form.puesto,
    form.experiencia,
    form.disponibilidad,
    archivoBase64 ? 'Sí' : 'No',
    cvInfo.name || '',
    cvInfo.url || '',
    form.experienciaManual || '',
    id,
    'Nuevo'
  ]);
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var viewName = form.puesto;
  if (!ss.getSheetByName(viewName)) {
    var newSheet = ss.insertSheet(viewName);
    newSheet.getRange('A1').setFormula(
      '=QUERY(\'Base General\'!A:N, "SELECT * WHERE F = \'' + form.puesto + '\'", 1)'
    );
  }
  return { success: true, id: id };
}

function getCandidatos(filtros) {
  var data = getMasterSheet().getDataRange().getValues();
  data.shift();
  return data.filter(r => {
    if (filtros.puesto && r[5] !== filtros.puesto) return false;
    if (filtros.estado && r[13] !== filtros.estado) return false;
    if (filtros.search) {
      var texto = (r[1]+' '+r[2]+' '+r[3]).toLowerCase();
      if (!texto.includes(filtros.search.toLowerCase())) return false;
    }
    return true;
  });
}

function cambiarEstado(row, nuevoEstado) {
  getMasterSheet().getRange(row, 14).setValue(nuevoEstado);
}

function getDashboardData() {
  var data = getMasterSheet().getDataRange().getValues();
  data.shift();
  var porPuesto = {}, porEstado = {}, porFecha = {};
  data.forEach(r => {
    var puesto = r[5]; if(puesto) porPuesto[puesto] = (porPuesto[puesto]||0)+1;
    var estado = r[13]; if(estado) porEstado[estado] = (porEstado[estado]||0)+1;
    var fecha = r[0] ? r[0].toString().slice(0,10) : '';
    if(fecha) porFecha[fecha] = (porFecha[fecha]||0)+1;
  });
  return { porPuesto, porEstado, porFecha, total: data.length };
}
