/**
 * PASE DE LISTA con QR + código de respaldo de 3 caracteres
 * Script vinculado a la hoja de cálculo (Extensiones > Apps Script).
 *
 * Hoja "Alumnos":    A Matrícula | B Nombre | C Grupo | D Código | E Token | F QR
 * Hoja "Asistencia": A Fecha | B Hora | C Matrícula | D Nombre | E Grupo | F Método | G Estatus
 */

const CFG = {
  HOJA_ALUMNOS: 'Alumnos',
  HOJA_ASISTENCIA: 'Asistencia',
  SESION_SEGUNDOS: 7200,     // la sesión dura 2 horas
  MAX_INTENTOS: 5,           // intentos de clave fallidos antes de bloquear
  BLOQUEO_SEGUNDOS: 600      // bloqueo de 10 minutos
};

/* ---------- Página web ---------- */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Pase de lista')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Pase de lista')
    .addItem('Generar códigos y QR', 'generarCodigosYQR')
    .addToUi();
}

/* ---------- Utilidades internas ---------- */

function _ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function _props() { return PropertiesService.getScriptProperties(); }

function _hash(texto, sal) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, sal + texto, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function _validar(sesion) {
  if (!sesion || !CacheService.getScriptCache().get('ses_' + sesion)) {
    throw new Error('SESION_EXPIRADA');
  }
}

function _alumnos() {
  const hoja = _ss().getSheetByName(CFG.HOJA_ALUMNOS);
  const n = hoja.getLastRow();
  if (n < 2) return [];
  return hoja.getRange(2, 1, n - 1, 5).getDisplayValues()
    .map(r => ({
      matricula: r[0].trim(),
      nombre: r[1].trim(),
      grupo: r[2].trim(),
      codigo: r[3].trim().toUpperCase(),
      token: r[4].trim()
    }))
    .filter(a => a.matricula);
}

/* ---------- Clave de acceso ---------- */

/**
 * EJECUTAR UNA SOLA VEZ desde el editor:
 * 1) escribe tu clave abajo, 2) ejecuta la función, 3) borra la clave del código.
 */
function configurarClave() {
  const NUEVA_CLAVE = 'CAMBIA_ESTA_CLAVE';
  if (NUEVA_CLAVE === 'CAMBIA_ESTA_CLAVE' || NUEVA_CLAVE.length < 8) {
    throw new Error('Escribe una clave de al menos 8 caracteres.');
  }
  const sal = Utilities.getUuid();
  _props().setProperties({ CLAVE_SAL: sal, CLAVE_HASH: _hash(NUEVA_CLAVE, sal) });
}

function iniciarSesion(clave) {
  const cache = CacheService.getScriptCache();
  if (cache.get('bloqueo')) {
    throw new Error('Demasiados intentos. Espera unos minutos.');
  }
  const p = _props().getProperties();
  if (!p.CLAVE_HASH) throw new Error('El sistema aún no tiene clave configurada.');

  if (_hash(String(clave || ''), p.CLAVE_SAL) !== p.CLAVE_HASH) {
    const n = Number(cache.get('intentos') || 0) + 1;
    if (n >= CFG.MAX_INTENTOS) {
      cache.put('bloqueo', '1', CFG.BLOQUEO_SEGUNDOS);
      cache.remove('intentos');
      throw new Error('Demasiados intentos. Espera unos minutos.');
    }
    cache.put('intentos', String(n), CFG.BLOQUEO_SEGUNDOS);
    throw new Error('Clave incorrecta.');
  }
  cache.remove('intentos');
  const sesion = Utilities.getUuid() + Utilities.getUuid();
  cache.put('ses_' + sesion, '1', CFG.SESION_SEGUNDOS);
  return sesion;
}

/* ---------- Consultas ---------- */

function listarGrupos(sesion) {
  _validar(sesion);
  const grupos = {};
  _alumnos().forEach(a => { if (a.grupo) grupos[a.grupo] = true; });
  return Object.keys(grupos).sort();
}

function buscarPorCodigo(sesion, codigo, grupo) {
  _validar(sesion);
  codigo = String(codigo || '').trim().toUpperCase();
  if (!/^[0-9]{2}[A-Z]$/.test(codigo)) {
    throw new Error('El código son 2 números y 1 letra (por ejemplo, 45J).');
  }
  const res = _alumnos()
    .filter(a => a.codigo === codigo && (!grupo || a.grupo === grupo))
    .map(a => ({ matricula: a.matricula, nombre: a.nombre, grupo: a.grupo }));
  if (!res.length) throw new Error('No hay ningún alumno con ese código en este grupo.');
  return res;
}

/* ---------- Registro de asistencia ---------- */

function _registrar(alumno, metodo) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const hoja = _ss().getSheetByName(CFG.HOJA_ASISTENCIA);
    const tz = Session.getScriptTimeZone();
    const ahora = new Date();
    const fecha = Utilities.formatDate(ahora, tz, 'yyyy-MM-dd');
    const hora = Utilities.formatDate(ahora, tz, 'HH:mm:ss');

    const n = hoja.getLastRow();
    if (n >= 2) {
      const previos = hoja.getRange(2, 1, n - 1, 3).getDisplayValues();
      const previo = previos.find(r => r[0] === fecha && r[2] === alumno.matricula);
      if (previo) {
        return { ok: false, duplicado: true, nombre: alumno.nombre, grupo: alumno.grupo, fecha: fecha, hora: previo[1] };
      }
    }
    const fila = n + 1;
    hoja.getRange(fila, 1, 1, 7).setNumberFormat('@')
      .setValues([[fecha, hora, alumno.matricula, alumno.nombre, alumno.grupo, metodo, 'Presente']]);
    return { ok: true, duplicado: false, nombre: alumno.nombre, grupo: alumno.grupo, fecha: fecha, hora: hora };
  } finally {
    lock.releaseLock();
  }
}

function registrarPorQR(sesion, token) {
  _validar(sesion);
  token = String(token || '').trim();
  const alumno = token && _alumnos().find(a => a.token === token);
  if (!alumno) throw new Error('QR no válido.');
  return _registrar(alumno, 'QR');
}

function registrarPorCodigo(sesion, matricula) {
  _validar(sesion);
  const alumno = _alumnos().find(a => a.matricula === String(matricula || '').trim());
  if (!alumno) throw new Error('Alumno no encontrado.');
  return _registrar(alumno, 'Código');
}

/* ---------- Descarga para Excel ---------- */

function _celdaSegura(valor) {
  let s = String(valor);
  if (/^[=+\-@]/.test(s)) s = "'" + s;      // evita fórmulas maliciosas en Excel
  return '"' + s.replace(/"/g, '""') + '"';
}

function exportarCSV(sesion, desde, hasta, grupo) {
  _validar(sesion);
  const datos = _ss().getSheetByName(CFG.HOJA_ASISTENCIA).getDataRange().getDisplayValues();
  const encabezado = datos[0];
  const filas = datos.slice(1).filter(r =>
    (!desde || r[0] >= desde) && (!hasta || r[0] <= hasta) && (!grupo || r[4] === grupo));
  const csv = [encabezado].concat(filas)
    .map(r => r.map(_celdaSegura).join(','))
    .join('\r\n');
  return '\uFEFF' + csv;                    // BOM para que Excel muestre bien los acentos
}

/* ---------- Generar códigos, tokens y QR (menú "Pase de lista") ---------- */

function generarCodigosYQR() {
  const hoja = _ss().getSheetByName(CFG.HOJA_ALUMNOS);
  const n = hoja.getLastRow();
  if (n < 2) return;

  const datos = hoja.getRange(2, 1, n - 1, 5).getDisplayValues();
  const invalidos = [];
  const salida = datos.map((r, i) => {
    const matricula = r[0].trim();
    const inicial = r[1].trim().normalize('NFD').replace(/[\u0300-\u036f]/g, '').charAt(0).toUpperCase();
    const codigo = r[3].trim() || (matricula.slice(-2) + inicial);
    const token = r[4].trim() || Utilities.getUuid().replace(/-/g, '');
    if (!/^[0-9]{2}[A-Z]$/.test(codigo)) invalidos.push(i + 2);
    return [codigo, token];
  });
  hoja.getRange(2, 4, n - 1, 2).setNumberFormat('@').setValues(salida);

  const formulas = salida.map((_, i) =>
    ['=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=180x180&data="&E' + (i + 2) + ')']);
  hoja.getRange(2, 6, n - 1, 1).setFormulas(formulas);
  hoja.setRowHeights(2, n - 1, 110);

  let msg = 'Listo: códigos, tokens y QR generados.';
  if (invalidos.length) msg += '\nRevisa las filas ' + invalidos.join(', ') + ' (la matrícula debe terminar en 2 números).';
  SpreadsheetApp.getUi().alert(msg);
}
