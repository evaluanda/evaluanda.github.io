/**
 * API de Evaluanda para el servidor MCP — proyecto SEPARADO del principal.
 *
 * Por qué separado: la app principal se despliega como «ejecutar como el
 * usuario que accede» y con acceso restringido al dominio, así que una llamada
 * de servidor a servidor no puede autenticarse contra ella. Este mini-proyecto
 * se despliega como «ejecutar como YO», protegido por un token secreto que solo
 * conoce tu ordenador.
 *
 * Y por qué NO se hace con un segundo despliegue del proyecto principal: los
 * ajustes de despliegue son por despliegue, sí, pero doGet serviría la interfaz
 * completa ejecutándose con TU identidad a cualquiera que tuviese esa URL. Aquí
 * solo hay doPost y no hay interfaz que servir.
 *
 * Lee y escribe TU cuaderno directamente. Deliberadamente NO expone alumnado ni
 * calificaciones: la IA ve la planificación y nada más.
 *
 * Despliegue: ver README.md de esta carpeta.
 */

// ---------------------------------------------------------------------------
// CONFIGURACIÓN — los dos valores van en las PROPIEDADES DEL SCRIPT, no aquí:
//   ⚙ Configuración del proyecto → Propiedades de la secuencia de comandos
//     API_TOKEN     cadena larga y aleatoria, la misma que en el servidor MCP
//     CUADERNO_ID   id de tu «Evaluanda — Cuaderno» (está en su URL de Drive)
//     CIF_KEY       (opcional) la clave de cifrado de la app: sin ella la IA no
//                   puede leer ni escribir el TEXTO de las tareas (va cifrado).
// Así no acaban en ningún archivo ni se pierden al actualizar el código.
// ---------------------------------------------------------------------------

function doPost(e) {
  try {
    var props = PropertiesService.getScriptProperties();
    var token = props.getProperty('API_TOKEN');
    var datos = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!token || datos.token !== token) return json_({ ok: false, error: 'token' });

    var id = props.getProperty('CUADERNO_ID');
    if (!id) return json_({ ok: false, error: 'Falta la propiedad CUADERNO_ID.' });

    return json_({ ok: true, datos: ejecutar_(SpreadsheetApp.openById(id), datos) });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

/** Abrir la URL en el navegador confirma que el despliegue está vivo. */
function doGet() {
  return HtmlService.createHtmlOutput(
    '<meta charset="utf-8"><body style="font-family:system-ui;padding:24px">' +
    '<h2>Evaluanda — API</h2><p>El despliegue funciona. Esta API solo responde ' +
    'a peticiones POST con token.</p></body>');
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function ejecutar_(ss, d) {
  switch (d.accion) {
    case 'ping':             return { app: 'Evaluanda', cuaderno: ss.getName() };
    case 'listar_clases':    return listarClases_(ss);
    case 'listar_unidades':  return listarUnidades_(ss, d.area, d.curso);
    case 'listar_sesiones':  return listarSesiones_(ss, d.unidadId);
    case 'crear_sesion':     return crearSesion_(ss, d.sesion || {});
    case 'editar_sesion':    return editarSesion_(ss, d.sesionId, d.cambios || {});
    case 'listar_eventos':   return listarEventos_(ss, d.desde, d.hasta);
    case 'listar_tareas':    return listarTareas_(ss, d);
    case 'crear_tarea':      return crearTarea_(ss, d.tarea || {});
    case 'editar_tarea':     return editarTarea_(ss, d.tareaId, d.cambios || {});
    default: throw new Error('Acción desconocida: ' + d.accion);
  }
}

// ---------------------------------------------------------------------------
// Acceso a las hojas POR NOMBRE DE COLUMNA (no por posición): así, si la app
// principal añade columnas nuevas —ya ha pasado varias veces—, esto sigue
// funcionando sin tocar nada.
// ---------------------------------------------------------------------------

function tabla_(ss, nombre) {
  var sh = ss.getSheetByName(nombre);
  if (!sh) throw new Error('No encuentro la pestaña ' + nombre + '. ¿Es el cuaderno correcto?');
  var datos = sh.getDataRange().getValues();
  var cab = (datos[0] || []).map(function (c) { return String(c); });
  var filas = [];
  for (var i = 1; i < datos.length; i++) {
    if (!datos[i][0]) continue;
    var o = { _fila: i + 1 };
    cab.forEach(function (c, j) { o[c] = datos[i][j]; });
    filas.push(o);
  }
  return { sh: sh, cab: cab, filas: filas };
}

/** Índice (base 1) de una columna por su nombre; 0 si no existe. */
function col_(t, nombre) { return t.cab.indexOf(nombre) + 1; }

function parse_(v) {
  if (!v) return [];
  try { var x = JSON.parse(v); return Array.isArray(x) ? x : []; } catch (e) { return []; }
}

function nuevoId_(pref) {
  return pref + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 12);
}

// ---------------------------------------------------------------------------
// Lecturas
// ---------------------------------------------------------------------------

function listarClases_(ss) {
  var ev = tabla_(ss, '_evaluaciones');
  var cl = tabla_(ss, '_clases');
  var grupo = {};
  cl.filas.forEach(function (c) { grupo[c.claseId] = c.nombre; });
  return ev.filas
    .filter(function (e) { return !e.archivado; })
    .map(function (e) {
      return {
        claseId: e.evalId,
        nombre: e.nombre || e.area || 'Clase',
        area: e.area || '',
        curso: grupo[e.claseId] ? '' : '',   // el curso vive en el grupo
        grupo: grupo[e.claseId] || '',
        horario: parse_(e.horario)
      };
    });
}

function listarUnidades_(ss, area, curso) {
  return tabla_(ss, '_planUnidades').filas
    .filter(function (u) {
      return (!area || u.area === area) && (!curso || u.curso === curso);
    })
    .sort(function (a, b) { return (a.orden || 0) - (b.orden || 0); })
    .map(function (u) {
      return { unidadId: u.unidadId, nombre: u.nombre, area: u.area || '',
        curso: u.curso || '', orden: Number(u.orden) || 0 };
    });
}

function listarSesiones_(ss, unidadId) {
  return tabla_(ss, '_planner').filas
    .filter(function (s) {
      if ((s.tipo || 'clase') !== 'clase') return false;
      return !unidadId || s.unidadId === unidadId;
    })
    .sort(function (a, b) { return (a.orden || 0) - (b.orden || 0); })
    .map(function (s) {
      return {
        sesionId: s.sesionId, titulo: s.titulo,
        descripcion: htmlAMd_(s.descripcion),
        criterios: parse_(s.criterios),
        unidadId: s.unidadId || '',
        fechas: parse_(s.asignaciones).map(function (a) {
          return { claseId: a.evalId, fecha: a.fecha, estado: a.estado };
        }),
        orden: Number(s.orden) || 0
      };
    });
}

// ---------------------------------------------------------------------------
// Markdown ⇄ HTML. El editor de Evaluanda guarda HTML de una lista blanca
// (b, i, u, a, ul, ol, li, p, br); una IA escribe y lee Markdown. Sin esta
// traducción, la negrita y las listas que escribe la IA llegaban como texto
// plano, y al revés la IA veía «**» literales.
//
// La conversión va en los DOS sentidos a propósito: si solo se tradujera al
// escribir, en cuanto la IA editase una sesión existente le quitaría el formato
// que ya tenía.
// ---------------------------------------------------------------------------

/** HTML del editor → Markdown, para que la IA lea el formato que ya hay. */
function htmlAMd_(html) {
  var t = String(html || '');
  if (!/<[a-z][\s\S]*>/i.test(t)) return desescapar_(t).trim();
  t = t
    .replace(/<(b|strong)>([\s\S]*?)<\/\1>/gi, '**$2**')
    .replace(/<(i|em)>([\s\S]*?)<\/\1>/gi, '*$2*')
    .replace(/<u>([\s\S]*?)<\/u>/gi, '<u>$1</u>')   // sin equivalente en Markdown
    .replace(/<a[^>]*href="([^"]*)"[^>]*>([\s\S]*?)<\/a>/gi, '[$2]($1)');
  // Listas: numeradas con su número, con viñeta con guion.
  t = t.replace(/<ol>([\s\S]*?)<\/ol>/gi, function (_, dentro) {
    var n = 0;
    return '\n' + dentro.replace(/<li>([\s\S]*?)<\/li>/gi, function (__, li) {
      n++; return n + '. ' + li.trim() + '\n';
    });
  });
  t = t.replace(/<ul>([\s\S]*?)<\/ul>/gi, function (_, dentro) {
    return '\n' + dentro.replace(/<li>([\s\S]*?)<\/li>/gi, function (__, li) {
      return '- ' + li.trim() + '\n';
    });
  });
  t = t.replace(/<li>([\s\S]*?)<\/li>/gi, '- $1\n')   // <li> suelto, por si acaso
    .replace(/<\/p>/gi, '\n\n').replace(/<br\s*\/?>/gi, '\n')
    .replace(/<(?!\/?u>)[^>]+>/g, '');                  // el resto fuera (menos <u>)
  return desescapar_(t).replace(/\n{3,}/g, '\n\n').trim();
}

/**
 * Markdown de la IA → HTML de la lista blanca del editor.
 *
 * Se recorre LÍNEA A LÍNEA, agrupando las seguidas del mismo tipo. Analizarlo
 * por bloques separados por línea en blanco fallaba en el caso más frecuente:
 * un encabezado y su lista pegados, «**DICTADO**\n1. …\n2. …», donde el bloque
 * entero no es una lista y la numeración se quedaba como texto.
 */
function mdAHtml_(md) {
  var t = String(md || '');
  if (!t.trim()) return '';
  // Ya viene en HTML del editor (una edición que no tocó la descripción). Se
  // miran solo etiquetas de BLOQUE: <u> o <b> sueltas son Markdown con una
  // etiqueta dentro, no HTML ya formado.
  if (/<(p|br|ul|ol|div)\b/i.test(t)) return t;

  var out = '', buf = [], lista = null;   // lista: 'ul' | 'ol' | null
  function cerrarLista() {
    if (!lista) return;
    out += '<' + lista + '>' + buf.join('') + '</' + lista + '>';
    buf = []; lista = null;
  }
  function cerrarParrafo() {
    if (lista || !buf.length) return;
    out += '<p>' + buf.join('<br>') + '</p>';
    buf = [];
  }
  t.replace(/\r/g, '').split('\n').forEach(function (l) {
    if (l.trim() === '') { cerrarLista(); cerrarParrafo(); return; }
    var vin = l.match(/^\s*[-*•]\s+(.*)$/);
    var num = l.match(/^\s*\d+[.)]\s+(.*)$/);
    if (vin || num) {
      var tipo = num ? 'ol' : 'ul';
      if (lista !== tipo) { cerrarLista(); cerrarParrafo(); lista = tipo; }
      buf.push('<li>' + enLinea_((vin || num)[1]) + '</li>');
      return;
    }
    cerrarLista();
    buf.push(enLinea_(l));
  });
  cerrarLista(); cerrarParrafo();
  return out;
}

/** Marcas de dentro de una línea (negrita, cursiva, enlaces, subrayado). */
function enLinea_(txt) {
  // Se escapa primero: nada de lo que escriba la IA puede colar etiquetas…
  var t = String(txt)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
  // …salvo <u>, que no tiene equivalente en Markdown y el editor sí admite.
  t = t.replace(/&lt;u&gt;([\s\S]*?)&lt;\/u&gt;/gi, '<u>$1</u>');
  return t
    .replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>')
    .replace(/__([^_]+)__/g, '<b>$1</b>')
    .replace(/(^|[^*])\*([^*\n]+)\*/g, '$1<i>$2</i>')
    .replace(/\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)/g, '<a href="$2">$1</a>')
    .replace(/(^|\s)(https?:\/\/[^\s<]+)/g, '$1<a href="$2">$2</a>');
}

function desescapar_(t) {
  return String(t).replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&');
}

// ---------------------------------------------------------------------------
// Escrituras. No hay borrado a propósito: que una IA pueda crear y corregir
// está bien; que pueda borrar tu planificación, no. Se borra desde la app.
// ---------------------------------------------------------------------------

function crearSesion_(ss, s) {
  var titulo = String(s.titulo || '').trim();
  if (!titulo) throw new Error('La sesión necesita un título.');
  var t = tabla_(ss, '_planner');
  var orden = 0;
  t.filas.forEach(function (f) { orden = Math.max(orden, Number(f.orden) || 0); });
  var id = nuevoId_('ses');
  var fila = new Array(t.cab.length).fill('');
  function set(nombre, valor) { var c = col_(t, nombre); if (c) fila[c - 1] = valor; }
  set('sesionId', id);
  set('titulo', titulo.slice(0, 300));
  set('descripcion', descripcionValida_(mdAHtml_(s.descripcion)));
  set('criterios', JSON.stringify(codigos_(s.criterios)));
  set('asignaciones', JSON.stringify(asignaciones_(s.fechas)));
  set('creado', new Date().toISOString());
  set('orden', orden + 1);
  set('tipo', 'clase');
  set('unidadId', String(s.unidadId || ''));
  t.sh.appendRow(fila);
  return { sesionId: id, titulo: titulo };
}

function editarSesion_(ss, sesionId, cambios) {
  if (!sesionId) throw new Error('Falta el sesionId.');
  var t = tabla_(ss, '_planner');
  var f = t.filas.filter(function (x) { return x.sesionId === sesionId; })[0];
  if (!f) throw new Error('No encuentro esa sesión.');
  function set(nombre, valor) {
    var c = col_(t, nombre);
    if (c) t.sh.getRange(f._fila, c).setValue(valor);
  }
  if (cambios.titulo != null) set('titulo', String(cambios.titulo).slice(0, 300));
  if (cambios.descripcion != null) set('descripcion', descripcionValida_(mdAHtml_(cambios.descripcion)));
  if (cambios.criterios != null) set('criterios', JSON.stringify(codigos_(cambios.criterios)));
  if (cambios.unidadId != null) set('unidadId', String(cambios.unidadId));
  if (cambios.fechas != null) set('asignaciones', JSON.stringify(asignaciones_(cambios.fechas)));
  return { sesionId: sesionId, ok: true };
}

/** Mismo tope que la app (Planner.gs, MAX_DESC): mejor un error claro que
 *  cortar en silencio una descripción larga. */
function descripcionValida_(html) {
  if (html.length > 45000) throw new Error('La descripción es demasiado larga: divídela en varias sesiones.');
  return html;
}

function codigos_(lista) {
  var vistos = {}, out = [];
  (Array.isArray(lista) ? lista : []).forEach(function (c) {
    var cod = String(c || '').trim();
    if (cod && !vistos[cod]) { vistos[cod] = true; out.push(cod); }
  });
  return out;
}

/** [{claseId, fecha}] → el formato de asignaciones del cuaderno. */
function asignaciones_(fechas) {
  return (Array.isArray(fechas) ? fechas : []).map(function (a) {
    var iso = String((a && a.fecha) || '').trim();
    return {
      evalId: String((a && a.claseId) || '').trim(),
      fecha: /^\d{4}-\d{2}-\d{2}$/.test(iso) ? iso : '',
      estado: 'pendiente',
      observaciones: ''
    };
  }).filter(function (a) { return !!a.evalId; });
}

// ---------------------------------------------------------------------------
// Reuniones y eventos (para vincular tareas a ellos)
// ---------------------------------------------------------------------------

/** Reuniones, eventos y avisos, con su fecha y hora (opcionalmente entre dos fechas). */
function listarEventos_(ss, desde, hasta) {
  return tabla_(ss, '_planner').filas
    .filter(function (s) { return (s.tipo || 'clase') !== 'clase' && s.tipo !== 'tramo'; })
    .map(function (s) {
      var a = parse_(s.asignaciones)[0] || {};
      return { sesionId: s.sesionId, tipo: s.tipo, titulo: s.titulo, fecha: a.fecha || '', hora: a.hora || '' };
    })
    .filter(function (e) {
      return (!desde || (e.fecha && e.fecha >= desde)) && (!hasta || (e.fecha && e.fecha <= hasta));
    })
    .sort(function (a, b) { return (a.fecha + a.hora) < (b.fecha + b.hora) ? -1 : 1; });
}

// ---------------------------------------------------------------------------
// Tareas (pestaña _tareas). El TEXTO va cifrado en el cuaderno (puede nombrar
// alumnado), con la clave de la app. Si CIF_KEY no está en las propiedades del
// script, la IA ve fechas, clases y estado, pero no el texto, y no puede
// escribirlo. Es una copia exacta del cifrado de la app (Cripto.gs, «enc2:»).
// ---------------------------------------------------------------------------

var RE_ISO_ = /^\d{4}-\d{2}-\d{2}$/;

/** 'yyyy-MM-dd' o '' (Sheets convierte ese texto en fecha al escribirlo). */
function fechaIso_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  v = String(v == null ? '' : v).trim();
  return RE_ISO_.test(v) ? v : '';
}
/** Valida una fecha que llega de la IA ('' la borra). */
function fechaEntrada_(v, campo) {
  v = String(v == null ? '' : v).trim();
  if (v && !RE_ISO_.test(v)) throw new Error('La fecha de ' + campo + ' debe ser AAAA-MM-DD.');
  return v;
}

function claveCif_() {
  var k = PropertiesService.getScriptProperties().getProperty('CIF_KEY');
  return k ? String(k).trim() : '';
}
function firmar_(texto) { return Utilities.computeHmacSha256Signature(texto, claveCif_()); }
function aSigned_(x) { x = x & 0xff; return x > 127 ? x - 256 : x; }
function flujoCif_(n, nonce) {
  var out = [], i = 0;
  while (out.length < n) {
    var b = firmar_(nonce + ':' + i++);
    for (var j = 0; j < b.length && out.length < n; j++) out.push(b[j] & 0xff);
  }
  return out;
}
function tagCif_(nonce, b64) { return Utilities.base64Encode(firmar_('mac:' + nonce + ':' + b64)).slice(0, 12); }
function cifrar_(texto) {
  if (!texto) return '';
  if (!claveCif_()) throw new Error('Para escribir el texto de una tarea falta CIF_KEY en las propiedades del script (ver README).');
  var nonce = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  var bytes = Utilities.newBlob(String(texto)).getBytes(), ks = flujoCif_(bytes.length, nonce), enc = [];
  for (var i = 0; i < bytes.length; i++) enc.push(aSigned_((bytes[i] & 0xff) ^ ks[i]));
  var b64 = Utilities.base64Encode(enc);
  return 'enc2:' + nonce + ':' + tagCif_(nonce, b64) + ':' + b64;
}
function descifrar_(dato) {
  var s = String(dato == null ? '' : dato);
  if (s.indexOf('enc') !== 0) return s; // texto plano
  if (!claveCif_()) return '(texto cifrado: añade CIF_KEY para verlo)';
  var p = s.indexOf('enc2:') === 0 ? s.slice(5).split(':') : null, nonce, b64;
  if (p) {
    if (p.length !== 3 || tagCif_(p[0], p[2]) !== p[1]) return '(no se pudo descifrar: ¿CIF_KEY correcta?)';
    nonce = p[0]; b64 = p[2];
  } else { // formato antiguo «enc:<nonce>:<base64>»
    var c = s.slice(4), sep = c.indexOf(':');
    if (sep < 0) return s;
    nonce = c.slice(0, sep); b64 = c.slice(sep + 1);
  }
  var enc = Utilities.base64Decode(b64), ks = flujoCif_(enc.length, nonce), dec = [];
  for (var i = 0; i < enc.length; i++) dec.push(aSigned_((enc[i] & 0xff) ^ ks[i]));
  return Utilities.newBlob(dec).getDataAsString();
}

function tareaSalida_(f) {
  return {
    tareaId: String(f.id), texto: descifrar_(f.texto),
    inicio: fechaIso_(f.inicio), vence: fechaIso_(f.fin),
    clases: f.claseId ? String(f.claseId).split(',') : [],
    sesionId: String(f.sesionId || ''), hecha: f.hecha === true || f.hecha === 'TRUE'
  };
}

/** Tareas, filtrables por rango (de vencimiento o, sin él, de inicio), clase y estado. */
function listarTareas_(ss, d) {
  return tabla_(ss, '_tareas').filas.map(tareaSalida_)
    .filter(function (t) {
      var f = t.vence || t.inicio;
      if (d.desde && (!f || f < d.desde)) return false;
      if (d.hasta && (!f || f > d.hasta)) return false;
      if (d.claseId && t.clases.indexOf(d.claseId) < 0) return false;
      if (d.pendientes && t.hecha) return false;
      return true;
    })
    .sort(function (a, b) { return ((a.vence || '9') < (b.vence || '9')) ? -1 : 1; });
}

function crearTarea_(ss, t) {
  var texto = String(t.texto || '').trim().slice(0, 300);
  if (!texto) throw new Error('La tarea necesita un texto.');
  var tabla = tabla_(ss, '_tareas'), ahora = new Date();
  var id = 't_' + ahora.getTime().toString(36) + Math.random().toString(36).slice(2, 6);
  var fila = new Array(tabla.cab.length).fill('');
  function set(nombre, valor) { var c = col_(tabla, nombre); if (c) fila[c - 1] = valor; }
  set('id', id);
  set('texto', cifrar_(texto));
  set('inicio', fechaEntrada_(t.inicio, 'inicio') || Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  set('fin', fechaEntrada_(t.vence, 'vencimiento'));
  set('claseId', (Array.isArray(t.clases) ? t.clases : []).join(','));
  set('hecha', false);
  set('orden', ahora.getTime());
  set('creado', ahora.toISOString());
  set('actualizado', ahora.toISOString());
  set('sesionId', String(t.sesionId || ''));
  tabla.sh.appendRow(fila);
  return { tareaId: id, texto: texto };
}

function editarTarea_(ss, tareaId, c) {
  if (!tareaId) throw new Error('Falta el tareaId.');
  var tabla = tabla_(ss, '_tareas');
  var f = tabla.filas.filter(function (x) { return String(x.id) === String(tareaId); })[0];
  if (!f) throw new Error('No encuentro esa tarea.');
  function set(nombre, valor) {
    var col = col_(tabla, nombre);
    if (col) tabla.sh.getRange(f._fila, col).setValue(valor);
  }
  if (c.texto != null) {
    var texto = String(c.texto).trim().slice(0, 300);
    if (!texto) throw new Error('El texto no puede quedar vacío (las tareas se borran desde la app).');
    set('texto', cifrar_(texto));
  }
  if (c.inicio != null) set('inicio', fechaEntrada_(c.inicio, 'inicio'));
  if (c.vence != null) set('fin', fechaEntrada_(c.vence, 'vencimiento'));
  if (c.clases != null) set('claseId', (Array.isArray(c.clases) ? c.clases : []).join(','));
  if (c.sesionId != null) set('sesionId', String(c.sesionId));
  if (c.hecha != null) set('hecha', !!c.hecha);
  set('actualizado', new Date().toISOString());
  return { tareaId: tareaId, ok: true };
}
