/**
 * Registro de consultas PESV  -  Google Apps Script
 *
 * Recibe las consultas de los conductores desde la app y las guarda en la hoja
 * "Consultas" de ESTE libro de Google Sheets.
 *
 * IMPORTANTE: cambia TOKEN por una clave larga y propia (mínimo 24 caracteres,
 * letras y números). La misma clave va en la app como SHEETS_TOKEN.
 */
const TOKEN = 'CAMBIA_ESTE_TOKEN_POR_UNO_LARGO_Y_PROPIO';
const NOMBRE_HOJA = 'Consultas';
const ENCABEZADOS = ['fecha_hora', 'conductor', 'cedula', 'ruta', 'id'];
const MAX_CARACTERES = 200;

function obtenerHoja_() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(NOMBRE_HOJA);
  if (!hoja) {
    hoja = libro.insertSheet(NOMBRE_HOJA);
    hoja.appendRow(ENCABEZADOS);
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function responder_(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}

function texto_(valor) {
  return String(valor === null || valor === undefined ? '' : valor).substring(0, MAX_CARACTERES);
}

// Abrir la dirección del servicio en el navegador sirve para comprobar que está activo.
function doGet() {
  return responder_({ ok: true, mensaje: 'Servicio de registro PESV activo' });
}

function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);
    if (!datos.token || datos.token !== TOKEN) {
      return responder_({ ok: false, error: 'no autorizado' });
    }
    const bloqueo = LockService.getScriptLock();
    bloqueo.waitLock(20000);
    try {
      if (datos.accion === 'registrar') return responder_(registrar_(datos.registros || []));
      if (datos.accion === 'listar') return responder_(listar_());
      return responder_({ ok: false, error: 'accion desconocida' });
    } finally {
      bloqueo.releaseLock();
    }
  } catch (err) {
    return responder_({ ok: false, error: String(err) });
  }
}

function registrar_(registros) {
  const hoja = obtenerHoja_();
  const ultima = hoja.getLastRow();
  const idsExistentes = new Set();
  if (ultima > 1) {
    hoja.getRange(2, 5, ultima - 1, 1).getValues().forEach(f => idsExistentes.add(String(f[0])));
  }
  let nuevos = 0;
  registros.forEach(r => {
    const id = texto_(r.id);
    if (id && idsExistentes.has(id)) return;          // ya registrado (reintento): se ignora
    const fila = hoja.getLastRow() + 1;
    const rango = hoja.getRange(fila, 1, 1, 5);
    rango.setNumberFormat('@');                         // texto plano: conserva ceros y evita fórmulas
    rango.setValues([[texto_(r.fecha_hora), texto_(r.conductor), texto_(r.cedula), texto_(r.ruta), id]]);
    idsExistentes.add(id);
    nuevos++;
  });
  return { ok: true, nuevos: nuevos };
}

function listar_() {
  const hoja = obtenerHoja_();
  const ultima = hoja.getLastRow();
  if (ultima < 2) return { ok: true, registros: [] };
  const valores = hoja.getRange(2, 1, ultima - 1, 5).getValues();
  return {
    ok: true,
    registros: valores.map(f => ({
      fecha_hora: String(f[0]), conductor: String(f[1]), cedula: String(f[2]),
      ruta: String(f[3]), id: String(f[4]),
    })),
  };
}
