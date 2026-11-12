/**
 * GASTOS CASA · back de la app (Google Apps Script)
 * Pega este código en Extensiones > Apps Script de la hoja "GastosCasa_datos_app".
 * La hoja debe tener la pestaña "Gastos" tal como viene en el archivo.
 *
 * Primera vez:
 *   1. Ejecuta `configurar` (botón ▶ arriba) y copia la clave que sale en el registro.
 *   2. Implementar > Nueva implementación > Aplicación web
 *      (Ejecutar como: yo · Quién tiene acceso: cualquier usuario). Copia la URL que acaba en /exec.
 *   3. Opcional: ejecuta una vez `activarCopias` para guardar una copia de la hoja cada lunes.
 */

const HOJA_GASTOS = 'Gastos';
const CAMPOS = ['id', 'fecha', 'categoria', 'importe', 'nota', 'creado', 'modificado'];
const CATEGORIAS = ['AGUA Y RESIDUOS', 'ARSENAL', 'BAJERA', 'COMUNIDAD', 'IBERDROLA', 'PISCINA',
  'SEGUROS CASA', 'SEGUROS COCHE', 'VARIOS INTERESES', 'VODAFONE', 'COMPRA', 'OTROS GASTOS'];

/** Genera la clave secreta de la app (solo hace falta una vez). */
function configurar() {
  const props = PropertiesService.getScriptProperties();
  let clave = props.getProperty('CLAVE');
  if (!clave) {
    clave = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
    props.setProperty('CLAVE', clave);
  }
  Logger.log('Clave de la app: ' + clave);
  return clave;
}

function doGet(e) {
  return responder_(() => {
    comprobarClave_(e.parameter.k);
    return { ok: true, gastos: leerGastos_(), hora: new Date().toISOString() };
  });
}

function doPost(e) {
  return responder_(() => {
    const p = JSON.parse(e.postData.contents);
    comprobarClave_(p.k);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const hechos = [];
      (p.ops || [p]).forEach(op => hechos.push(aplicar_(op)));
      return { ok: true, hechos: hechos };
    } finally {
      lock.releaseLock();
    }
  });
}

function aplicar_(op) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_GASTOS);
  const ahora = new Date();
  if (op.accion === 'crear') {
    const g = limpiar_(op.gasto);
    // si ya existe (reintento sin conexión), no se duplica
    if (g.id && buscarFila_(hoja, g.id) > 0) return { accion: 'crear', id: g.id, repetido: true };
    g.id = g.id || 'G' + Utilities.formatDate(ahora, 'UTC', 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    g.creado = ahora; g.modificado = ahora;
    hoja.appendRow(aFila_(g));
    return { accion: 'crear', id: g.id };
  }
  if (op.accion === 'editar') {
    const fila = buscarFila_(hoja, op.gasto.id);
    if (fila < 0) throw new Error('No encuentro el gasto ' + op.gasto.id);
    const actual = hoja.getRange(fila, 1, 1, CAMPOS.length).getValues()[0];
    const g = limpiar_(op.gasto);
    g.creado = actual[CAMPOS.indexOf('creado')] || '';
    g.modificado = ahora;
    hoja.getRange(fila, 1, 1, CAMPOS.length).setValues([aFila_(g)]);
    return { accion: 'editar', id: g.id };
  }
  if (op.accion === 'borrar') {
    const fila = buscarFila_(hoja, op.id);
    if (fila > 0) {
      // copia de seguridad antes de borrar
      const papelera = SpreadsheetApp.getActive().getSheetByName('Papelera') ||
        SpreadsheetApp.getActive().insertSheet('Papelera');
      papelera.appendRow(hoja.getRange(fila, 1, 1, CAMPOS.length).getValues()[0].concat([ahora]));
      hoja.deleteRow(fila);
    }
    return { accion: 'borrar', id: op.id };
  }
  throw new Error('Acción desconocida');
}

function limpiar_(g) {
  const cat = String(g.categoria || '').toUpperCase().trim();
  if (CATEGORIAS.indexOf(cat) < 0) throw new Error('Categoría desconocida: ' + g.categoria);
  const importe = Math.round(Number(g.importe) * 100) / 100;
  if (!isFinite(importe) || importe === 0) throw new Error('Importe no válido');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(g.fecha || ''))) throw new Error('Fecha no válida');
  return { id: g.id || '', fecha: g.fecha, categoria: cat, importe: importe, nota: String(g.nota || '').slice(0, 200) };
}

function leerGastos_() {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_GASTOS);
  const n = hoja.getLastRow() - 1;
  if (n < 1) return [];
  const tz = Session.getScriptTimeZone();
  return hoja.getRange(2, 1, n, 5).getValues()
    .filter(v => v[0] !== '' && v[2] !== '')
    .map(v => [String(v[0]), v[1] instanceof Date ? Utilities.formatDate(v[1], tz, 'yyyy-MM-dd') : String(v[1]).slice(0, 10),
      String(v[2]), Number(v[3]) || 0, String(v[4] || '')]);
}

function aFila_(g) {
  return CAMPOS.map(c => {
    const x = g[c];
    if (x === undefined || x === null) return '';
    if (c === 'fecha') { const p = String(x).split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
    return x;
  });
}

function buscarFila_(hoja, id) {
  const n = hoja.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = hoja.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function comprobarClave_(k) {
  const clave = PropertiesService.getScriptProperties().getProperty('CLAVE');
  if (!clave || k !== clave) throw new Error('CLAVE');
}

function responder_(fn) {
  let out;
  try { out = fn(); } catch (err) { out = { ok: false, error: String(err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ====================== copia de seguridad semanal ====================== */
/** Ejecútala una vez: guarda una copia de la hoja cada lunes en la carpeta "Gastos Casa copias" (las 8 últimas). */
function activarCopias() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'copiaSemanal').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('copiaSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(4).create();
  copiaSemanal();
}
function copiaSemanal() {
  const ss = SpreadsheetApp.getActive();
  const madre = DriveApp.getFileById(ss.getId()).getParents();
  const base = madre.hasNext() ? madre.next() : DriveApp.getRootFolder();
  const it = base.getFoldersByName('Gastos Casa copias');
  const carpeta = it.hasNext() ? it.next() : base.createFolder('Gastos Casa copias');
  DriveApp.getFileById(ss.getId()).makeCopy('Gastos Casa · copia ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), carpeta);
  const copias = []; const f = carpeta.getFiles(); while (f.hasNext()) copias.push(f.next());
  copias.sort((x, y) => y.getDateCreated() - x.getDateCreated()).slice(8).forEach(x => x.setTrashed(true));
}
