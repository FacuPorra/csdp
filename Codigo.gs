/**
 * CSDP Progresiones — base compartida en Google Sheets
 * ----------------------------------------------------
 * Pegá este archivo en Extensiones > Apps Script de tu planilla (reemplazando Código.gs).
 * Después corré una vez la función setup() y mirá la clave de admin en el registro de ejecución.
 */

// ====== CONFIGURACIÓN (podés editar esto) ======
const PREPARADORES = ['Preparador 1', 'Preparador 2', 'Preparador 3'];
const TZ = 'America/Argentina/Cordoba';
// Link de la app instalable en GitHub Pages (ej: 'https://tuusuario.github.io/csdp/').
// Si lo dejás vacío, los links de alumnos usan la URL de Google.
const APP_URL = '';

// Categorías con su color (tono 0-360) y ejercicios iniciales.
const CATEGORIAS = [
  { n: 'Fuerza Principales', h: 210, items: ['Sentadillas', 'Visagra de caderas', 'Pecho', 'Piernas unipodal', 'Espalda', 'Dominadas', 'Progresion trx', 'Flexiones de brazos'] },
  { n: 'Potencia Principales', h: 32, items: ['Cargadas', 'Arranques', 'Segundo tiempo'] },
  { n: 'Fuerza Secundarios', h: 150, items: ['Subidas al banco', 'Bulgaras', 'Pistolas', 'Estocadas', 'Zancadas', 'Sentadilla Split', 'Sentadilla Lateral', 'Copa', 'Frontal', 'Sentadilla Atrás', 'Sentadillas de arranque', 'Swing con pesa rusa', 'Peso muerto', 'Peso muerto 1 pie', 'Remo acostado', 'Remo con mancuerna', 'Remo Menton', 'Remo 1 Brazo beer dog', 'Biceps', 'Biceps con Hombros', 'Apertura Pecho', 'Press Inclinado', 'Press Mancuernas', 'Empujes de Barra diagonal', 'Pres militar', 'Pres Arnold', 'Vuelos frontales', 'Vuelos laterales', 'Pull Over', 'Pull over a 1 brazo', 'Push Press', 'Triceps', 'Sentadilla isometrica', 'Flexiones excentricas', 'Dominadas trx', 'Pull Over trx'] },
  { n: 'Potencia Secundarios', h: 300, items: ['Cargadas 1 pie', 'Angelitos Peso', 'Arranque a 1 brazo', 'Arranque 1 Pie (Barra)', 'Arranque 1 Pie (Mancuerna)', 'Remo al menton con impulso (Barra)', 'Remo al menton con impulso (Mancuerna)', 'Arranque con Rusa', 'Cargada con Rusa', 'Segundo tiempo con mancuernas', 'Segundo tiempo con barra'] },
  { n: 'Dinámicos (sin peso)', h: 330, items: ['Lanzamientos al piso', 'Lanzamientos a la pared', 'Recepcion en Flexion', 'Flexiones con aplauso', 'Impulso de tiron+frenado', 'Impulso de dominada+frenado', 'Tecnica de cargada', 'Tecnica de arranque'] },
  { n: 'Dinámicos Tren Inferior', h: 235, items: ['Recepciones en estocadas', 'Saltos con Carga', 'Saltos al cajón / 4 reps', 'Saltos al Cajón unipodales / 4+4 reps', 'Saltos Drop / 4 reps', 'Saltos Asimétricos al cajón / 4+4 reps', 'Saltos Pogo / 8 reps', 'Saltos Asistidos / 8 reps', 'Saltos con rebote previo / 4 reps', 'Saltos Sin Carga / 4 reps', 'Saltos de sentado / 6 reps', 'Saltos con Aproximación / 3+3 reps'] },
  { n: 'Zona Media', h: 190, items: ['Abdominales Peso', 'Twist', 'Espinales con Peso', 'Elevaciones de cadera', 'Bicho muerto', 'Bicho muerto unilateral', 'Oblicuos Diagonales', 'Espinales con mancuernas'] },
  { n: 'Zona Media (sin peso)', h: 175, items: ['Caminata', 'Zona media sobre pelota', 'Antimovimiento', 'Perro de casa', 'Planchas', 'Retracciones escapulares', 'Lanzamientos laterales'] }
];

const HOJAS = {
  Alumnos: ['id', 'token', 'nombre', 'prep', 'creado'],
  Ejercicios: ['categoria', 'nombre'],
  Planes: ['alumnoId', 'actualizado', 'items'],
  Registros: ['id', 'alumnoId', 'fecha', 'ejercicio', 'peso', 'reps', 'testigo']
};

// ====== WEB APP ======
function doGet(e) {
  const p = (e && e.parameter) || {};
  const t = HtmlService.createTemplateFromFile('Index');
  t.token = String(p.a || '');
  return t.evaluate()
    .setTitle('CSDP Progresiones')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // permite abrirla dentro de la app instalable
}

// ====== SETUP (correr una vez desde el editor) ======
function setup() {
  const ss = ss_();
  Object.keys(HOJAS).forEach(function (n) {
    let sh = ss.getSheetByName(n);
    if (!sh) sh = ss.insertSheet(n);
    if (sh.getLastRow() === 0) sh.appendRow(HOJAS[n]);
    sh.setFrozenRows(1);
    sh.getRange('A:Z').setNumberFormat('@'); // todo como texto: evita que Sheets convierta fechas/ids
  });
  const ej = ss.getSheetByName('Ejercicios');
  if (ej.getLastRow() < 2) {
    const rows = [];
    CATEGORIAS.forEach(function (c) { c.items.forEach(function (i) { rows.push([c.n, i]); }); });
    ej.getRange(2, 1, rows.length, 2).setValues(rows);
  }
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_KEY')) {
    props.setProperty('ADMIN_KEY', 'csdp' + Math.floor(1000 + Math.random() * 9000));
  }
  Logger.log('Listo. Tu clave de admin es: ' + props.getProperty('ADMIN_KEY'));
}

// Para cambiar la clave: editá el texto y corré esta función.
function cambiarClave() {
  PropertiesService.getScriptProperties().setProperty('ADMIN_KEY', 'NUEVA-CLAVE-ACA');
  Logger.log('Clave cambiada.');
}

// ====== API ADMIN ======
function adminLogin(clave) {
  checkKey_(clave);
  return estado_();
}

function adminAccion(clave, accion, datos) {
  checkKey_(clave);
  if (!ACCIONES[accion]) throw new Error('Acción desconocida: ' + accion);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let res;
  try {
    res = ACCIONES[accion](datos || {});
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { estado: estado_(), res: res || null };
}

const ACCIONES = {
  crearAlumno: function (d) {
    const nombre = limpia_(d.nombre);
    if (!nombre) throw new Error('Falta el nombre');
    const id = 'a' + uid_();
    hoja_('Alumnos').appendRow([id, nuevoToken_(), nombre, String(Number(d.prep) || 0), ahora_()]);
    return id;
  },
  editarAlumno: function (d) {
    const r = buscar_('Alumnos', 'id', d.id);
    if (!r) throw new Error('No encontré al alumno');
    const sh = hoja_('Alumnos');
    sh.getRange(r._row, 3, 1, 2).setValues([[limpia_(d.nombre) || r.nombre, String(Number(d.prep) || 0)]]);
  },
  nuevoLink: function (d) {
    const r = buscar_('Alumnos', 'id', d.id);
    if (!r) throw new Error('No encontré al alumno');
    hoja_('Alumnos').getRange(r._row, 2).setValue(nuevoToken_());
  },
  borrarAlumno: function (d) {
    borrarDonde_('Registros', 'alumnoId', d.id);
    borrarDonde_('Planes', 'alumnoId', d.id);
    borrarDonde_('Alumnos', 'id', d.id);
  },
  guardarPlan: function (d) {
    const items = (d.items || []).map(function (i) {
      return { n: limpia_(i.n), series: limpia_(i.series), reps: limpia_(i.reps), peso: limpia_(i.peso), nota: limpia_(i.nota) };
    }).filter(function (i) { return i.n; });
    const json = JSON.stringify(items);
    const r = buscar_('Planes', 'alumnoId', d.id);
    if (r) hoja_('Planes').getRange(r._row, 2, 1, 2).setValues([[ahora_(), json]]);
    else hoja_('Planes').appendRow([d.id, ahora_(), json]);
  },
  agregarRegistros: function (d) {
    const rows = (d.registros || []).map(function (x) {
      return ['r' + uid_(), d.id, limpia_(x.fecha) || hoy_(), limpia_(x.ejercicio), String(x.peso), String(x.reps), x.testigo ? 'si' : ''];
    }).filter(function (r) { return r[3]; });
    if (!rows.length) throw new Error('No hay registros para guardar');
    const sh = hoja_('Registros');
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return rows.length;
  },
  borrarRegistro: function (d) { borrarDonde_('Registros', 'id', d.id); },
  agregarEjercicios: function (d) {
    const cat = limpia_(d.categoria);
    const have = {};
    filas_('Ejercicios').forEach(function (e) { have[String(e.nombre).toLowerCase()] = 1; });
    const rows = [];
    (d.nombres || []).forEach(function (n) {
      n = limpia_(n);
      if (n && !have[n.toLowerCase()]) { rows.push([cat, n]); have[n.toLowerCase()] = 1; }
    });
    if (rows.length) {
      const sh = hoja_('Ejercicios');
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, 2).setValues(rows);
    }
    return rows.length;
  },
  borrarEjercicio: function (d) { borrarDonde_('Ejercicios', 'nombre', d.nombre); }
};

// ====== API ALUMNO (solo lectura, con su link) ======
function vistaAlumno(token) {
  token = String(token || '');
  if (token.length < 10) throw new Error('Link no válido');
  const a = filas_('Alumnos').filter(function (r) { return String(r.token) === token; })[0];
  if (!a) throw new Error('Este link ya no es válido. Pedile uno nuevo a tu profe.');
  const plan = filas_('Planes').filter(function (p) { return p.alumnoId === a.id; })[0];
  const marcas = filas_('Registros').filter(function (m) { return m.alumnoId === a.id; }).map(marca_);
  return {
    nombre: a.nombre,
    prep: PREPARADORES[Number(a.prep)] || '',
    plan: plan ? { actualizado: plan.actualizado, items: parse_(plan.items) } : { actualizado: '', items: [] },
    marcas: marcas,
    cat: catalogo_()
  };
}

// ====== ESTADO COMPLETO (admin) ======
function estado_() {
  const planes = {};
  filas_('Planes').forEach(function (p) { planes[p.alumnoId] = { actualizado: p.actualizado, items: parse_(p.items) }; });
  const marcas = {};
  filas_('Registros').forEach(function (m) { (marcas[m.alumnoId] = marcas[m.alumnoId] || []).push(marca_(m)); });
  return {
    preps: PREPARADORES,
    url: APP_URL ? APP_URL.replace(/\/?$/, '/') : ScriptApp.getService().getUrl(),
    hoy: hoy_(),
    cat: catalogo_(),
    alumnos: filas_('Alumnos').map(function (a) {
      return {
        id: a.id, token: a.token, nombre: a.nombre, prep: Number(a.prep) || 0,
        plan: planes[a.id] || { actualizado: '', items: [] },
        marcas: marcas[a.id] || []
      };
    })
  };
}

function catalogo_() {
  const orden = CATEGORIAS.map(function (c) { return c.n; });
  const map = {};
  filas_('Ejercicios').forEach(function (e) {
    if (!map[e.categoria]) map[e.categoria] = [];
    map[e.categoria].push(e.nombre);
  });
  const nombres = orden.concat(Object.keys(map).filter(function (k) { return orden.indexOf(k) < 0; }));
  return nombres.map(function (n, i) {
    const c = CATEGORIAS.filter(function (x) { return x.n === n; })[0];
    return { n: n, h: c ? c.h : (i * 47) % 360, items: map[n] || [] };
  });
}

// ====== HELPERS ======
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function hoja_(n) {
  const sh = ss_().getSheetByName(n);
  if (!sh) throw new Error('Falta la hoja "' + n + '". Corré setup() desde el editor.');
  return sh;
}
function filas_(n) {
  const v = hoja_(n).getDataRange().getValues();
  const h = v.shift() || [];
  return v.map(function (r, i) {
    const o = { _row: i + 2 };
    h.forEach(function (k, j) {
      let x = r[j];
      if (x instanceof Date) x = Utilities.formatDate(x, TZ, k === 'actualizado' ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd');
      o[k] = typeof x === 'string' ? x : String(x);
    });
    return o;
  });
}
function buscar_(n, col, val) {
  return filas_(n).filter(function (r) { return r[col] === String(val); })[0] || null;
}
function borrarDonde_(n, col, val) {
  const sh = hoja_(n);
  filas_(n).filter(function (r) { return r[col] === String(val); })
    .map(function (r) { return r._row; })
    .sort(function (a, b) { return b - a; })
    .forEach(function (row) { sh.deleteRow(row); });
}
function marca_(m) {
  return { id: m.id, fecha: m.fecha, ejercicio: m.ejercicio, peso: parseFloat(m.peso), reps: m.reps, testigo: m.testigo === 'si' };
}
function parse_(s) { try { return JSON.parse(s) || []; } catch (e) { return []; } }
function checkKey_(k) {
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!real) throw new Error('Falta configurar: corré setup() desde el editor.');
  if (String(k || '') !== real) throw new Error('Clave incorrecta');
}
function limpia_(s) {
  s = String(s == null ? '' : s).trim().slice(0, 300);
  return /^[=+\-@]/.test(s) && isNaN(Number(s)) ? "'" + s : s; // evita fórmulas inyectadas
}
function uid_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }
function nuevoToken_() { return Utilities.getUuid().replace(/-/g, '') + uid_(); }
function hoy_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function ahora_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
