/**
 * SCRIPT V29 — PRODUCCIÓN
 * CUMBRE CHP 2026
 * Fix: ACREDITACION 1/2 usa fecha del staff, no fecha del servidor
 */

const SHEET_ID = "1U8dfUYE4rVF_QQC3HOFENZ0e6gLTGFwZeqcLAPSJ3jk";

function doGet(e)  { return handleResponse(e); }
function doPost(e) { return handleResponse(e); }

function handleResponse(e) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (f) { return jsonOut({ ok: false, error: "Timeout" }); }

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let p = {};

    if (e.postData && e.postData.contents) {
      try { p = JSON.parse(e.postData.contents); } catch (err) { p = e.parameter; }
    } else {
      p = e.parameter;
    }

    if (!p || Object.keys(p).length === 0) {
      return jsonOut({ ok: true, mensaje: "Conexión activa V29 - OK" });
    }

    const nombreValido = p["tu-nombre-y-apellido"] || p.nombre || p.nombre_completo || "";
    const correoValido = p["tu-correo"] || p.email || "";

    if (nombreValido === "" && correoValido === "" && !p.tipo && !p.action && !p.accion) {
      return ContentService.createTextOutput("IGNORADO: SIN DATOS");
    }

    if (p.tipo || p.action || p.accion) {
      return procesarLogicaApps(ss, p);
    } else {
      return procesarInscripcionWeb(ss, p);
    }

  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

function procesarLogicaApps(ss, datos) {
  var accion = datos.tipo || datos.action || datos.accion;
  var esp = (datos.especialidad || datos.profesion || "N/A").toLowerCase();
  var sufijo = (esp.includes('enfermer') || esp.includes('otro')) ? ' 1' : ' 2';

  // 1. ACREDITACIÓN STAFF
  if (accion === 'acreditacion' || accion === 'acreditacion_staff') {
    // Usar la fecha enviada por el staff, no la del servidor
    var fechaStaff = datos.fecha || '';
    var esDia1 = fechaStaff.indexOf('05-07') !== -1 || 
                 fechaStaff.indexOf('07/05') !== -1 ||
                 (datos.dia || '').indexOf('1') !== -1 ||
                 (datos.dia || '').toLowerCase().indexOf('día 1') !== -1 ||
                 (datos.dia || '').toLowerCase().indexOf('dia 1') !== -1;
    var nombreHoja = esDia1 ? "ACREDITACION 1" : "ACREDITACION 2";
    var hojaA = obtenerHojaFormateada(ss, nombreHoja,
      ["Timestamp","UID","Nombre Completo","Especialidad","Pase","Dia","Fecha","Hora","Staff ID"],
      "#2E4053");
    hojaA.appendRow([
      new Date(),
      datos.uid          || "N/A",
      datos.nombre       || datos.nombre_completo || "N/A",
      datos.especialidad || datos.profesion || "N/A",
      datos.pase         || "N/A",
      datos.dia          || "N/A",
      datos.fecha        || "N/A",
      datos.hora         || new Date().toLocaleTimeString(),
      datos.staffId      || "N/A"
    ]);
    return jsonOut({ ok: true, success: true, hoja: nombreHoja });
  }

  // 2. REGISTRO PARTICIPANTES
  else if (accion === 'register' || accion === 'registro_asistente') {
    var hojaP = obtenerHojaFormateada(ss, 'Participantes' + sufijo,
      ['Timestamp','UID','Nombre Completo','Email','Telefono','Profesion'],
      '#1A5276');
    hojaP.appendRow([
      new Date(),
      datos.uid || 'CHP-NEW',
      datos.nombre_completo || datos.nombre || '',
      normalizarEmail(datos.email),
      limpiarTelefono(datos.telefono || datos.telf),
      esp
    ]);
    return jsonOut({ ok: true, success: true });
  }

  // 3. PRE-INSCRIPCIÓN A CONFERENCIAS
  else if (accion === 'preinscripcion_conferencia') {
    var hojaCP = obtenerHojaFormateada(ss, 'CONFERENCIA_PLANIFICADA' + sufijo,
      ['Timestamp','Nombre Completo','Email','Especialidad','Nombre Conferencia'],
      '#8E44AD');
    hojaCP.appendRow([
      new Date(),
      datos.nombre_completo || '',
      normalizarEmail(datos.email),
      esp,
      datos.nombre_conf || ''
    ]);
    return jsonOut({ ok: true, success: true });
  }

  // 4. TRACKING QR
  else if (accion === 'tracking' || accion.includes('qr')) {
    var hojaT = obtenerHojaFormateada(ss, 'CONFERENCIA_REAL_CERTIFICADO' + sufijo,
      ['Timestamp','UID','Email','ID Conf','Nombre Conf','Especialidad'],
      '#1E8449');
    hojaT.appendRow([
      new Date(),
      datos.uid || '',
      normalizarEmail(datos.email),
      datos.id_conf || '',
      datos.nombre_conf || '',
      esp
    ]);
    return jsonOut({ ok: true, success: true });
  }

  // 5. REGISTRO ENFERMERAS — Landing CTH CHP 2026
  else if (accion === 'registroEnfermera') {
    var hojaEnf = obtenerHojaFormateada(ss, 'ENFERMERAS_RED_CTH',
      ['Timestamp','Nombre','Email','Teléfono','Nivel Académico','Institución','Cargo','Curas a Domicilio','Fuente'],
      '#00796B');
    hojaEnf.appendRow([
      datos.timestamp   || new Date().toLocaleString('es-VE'),
      datos.nombre      || '',
      normalizarEmail(datos.email || ''),
      limpiarTelefono(datos.telefono || ''),
      datos.nivel       || '',
      datos.institucion || '',
      datos.cargo       || '',
      datos.domicilio   || '',
      datos.fuente      || 'Landing Enfermeras CHP 2026'
    ]);
    return jsonOut({ ok: true, success: true, accion: 'registroEnfermera' });
  }

  // 6. LOG GENERAL
  else {
    var hojaLog = obtenerHojaFormateada(ss, 'LOG_GENERAL',
      ['Timestamp','Accion','Email','Datos'],
      '#566573');
    hojaLog.appendRow([
      new Date(),
      accion,
      normalizarEmail(datos.email),
      JSON.stringify(datos)
    ]);
    return jsonOut({ ok: true, success: true });
  }
}

function procesarInscripcionWeb(ss, p) {
  const email = normalizarEmail(p["tu-correo"]);
  if (email === "") return ContentService.createTextOutput("SUCCESS");

  const cat = (p["categoria-chp"] || "").toLowerCase().trim();
  const esp = (p["tu-especialidad"] || "").trim();

  let nombreHoja = "MEDICO";
  if (cat === "estudiante" || (esp !== "" && p["tu-numero-de-carnet"] === "")) nombreHoja = "ESTUDIANTE";
  else if (cat === "enfermero" || esp.toLowerCase().includes("enfermer") || esp === "") nombreHoja = "ENFERMERO";

  let hoja = ss.getSheetByName(nombreHoja) || ss.insertSheet(nombreHoja);
  const fechaActual = Utilities.formatDate(new Date(), "GMT-4", "dd/MM/yyyy HH:mm:ss");

  const fila = [
    "", fechaActual, "unread", esp,
    p["tu-nombre-y-apellido"] || "",
    p["tu-lugar-de-trabajo"]  || "",
    p["tu-numero-de-carnet"]  || "",
    email,
    limpiarTelefono(p["tu-whatsapp"] || ""),
    p["tu-estado-ciudad"] || "",
    p["tu-instagram"]     || "",
    p["tu-encuesta"]      || "",
    p["tu-mensaje"]       || ""
  ];

  const vals = hoja.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (normalizarEmail(vals[i][7]) === email) return ContentService.createTextOutput("SUCCESS");
  }
  hoja.appendRow(fila);
  return ContentService.createTextOutput("SUCCESS");
}

function obtenerHojaFormateada(ss, nombre, headers, color) {
  var hoja = ss.getSheetByName(nombre);
  if (!hoja) {
    hoja = ss.insertSheet(nombre);
    hoja.appendRow(headers);
    hoja.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground(color)
        .setFontColor('#ffffff');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function limpiarTelefono(tel) {
  if (!tel) return "";
  let s = tel.toString().replace(/[^0-9]/g, "");
  if (s.startsWith("0")) s = s.substring(1);
  if (!s.startsWith("58")) s = "58" + s;
  return "+" + s;
}

function normalizarEmail(e) { return String(e || "").toLowerCase().trim(); }

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.TEXT);
}
