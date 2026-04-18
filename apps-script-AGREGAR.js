// ═══════════════════════════════════════════════════════════════════
// AGREGAR AL FINAL DEL APPS SCRIPT EXISTENTE
// No borrar nada — pegar esto debajo de todo el código actual
// ═══════════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    var datos = JSON.parse(e.postData.contents);
    var tipo  = datos.tipo || '';
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hoja;

    if (tipo === 'acreditacion') {
      // ── Datos de acreditación del staff ──
      hoja = ss.getSheetByName('ACREDITACION');
      if (!hoja) {
        hoja = ss.insertSheet('ACREDITACION');
        // Encabezados con formato
        var headers = ['Timestamp','Fecha','Día','Hora','Nombre Completo',
                       'Especialidad','Pase','UID Asistente','Puesto Staff'];
        hoja.appendRow(headers);
        hoja.getRange(1,1,1,headers.length)
            .setFontWeight('bold')
            .setBackground('#1A4F7A')
            .setFontColor('#ffffff');
        hoja.setFrozenRows(1);
      }

      // Anti-duplicado: verificar si ya existe este UID para esta fecha
      var datos_hoja = hoja.getDataRange().getValues();
      for (var i = 1; i < datos_hoja.length; i++) {
        var fila_uid   = String(datos_hoja[i][7]); // columna UID
        var fila_fecha = String(datos_hoja[i][1]); // columna Fecha
        if (fila_uid === String(datos.uid) && fila_fecha === String(datos.fecha)) {
          // Ya existe — no duplicar
          return ContentService
            .createTextOutput(JSON.stringify({ ok: true, duplicado: true }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }

      // Nuevo registro
      hoja.appendRow([
        new Date(),           // Timestamp automático
        datos.fecha   || '',  // Fecha del evento (2026-05-07)
        datos.dia     || '',  // Día 1 / Día 2 / Día 3
        datos.hora    || '',  // Hora de acreditación
        datos.nombre  || '',  // Nombre completo
        datos.especialidad || '',
        datos.pase    || '',  // ENFERMERÍA / MÉDICO / ESTUDIANTE
        datos.uid     || '',  // ID único del asistente
        datos.staffId || ''   // Puesto del staff (Staff-01, etc.)
      ]);

    } else if (tipo === 'tracking') {
      // ── Tracking de conferencias (app participantes) ──
      hoja = ss.getSheetByName('ASISTENCIA');
      if (!hoja) {
        hoja = ss.insertSheet('ASISTENCIA');
        var h2 = ['Timestamp','Email','Nombre Completo','Especialidad',
                  'ID Conferencia','Nombre Conferencia','Día','Hora Conferencia',
                  'Total Conferencias','Día Evento'];
        hoja.appendRow(h2);
        hoja.getRange(1,1,1,h2.length)
            .setFontWeight('bold')
            .setBackground('#00897B')
            .setFontColor('#ffffff');
        hoja.setFrozenRows(1);
      }
      hoja.appendRow([
        new Date(),
        datos.email          || '',
        datos.nombre_completo|| '',
        datos.especialidad   || '',
        datos.id_conf        || '',
        datos.nombre_conf    || '',
        datos.dia            || '',
        datos.hora           || '',
        datos.total_conferencias || '',
        datos.dia_evento     || ''
      ]);

    } else if (tipo === 'solicitud_certificado') {
      // ── Solicitudes de certificado ──
      hoja = ss.getSheetByName('CERTIFICADOS_PENDIENTES');
      if (!hoja) {
        hoja = ss.insertSheet('CERTIFICADOS_PENDIENTES');
        var h3 = ['Timestamp','Email','Nombre Completo','Especialidad',
                  'Total Conferencias','IDs Conferencias','Estado'];
        hoja.appendRow(h3);
        hoja.getRange(1,1,1,h3.length)
            .setFontWeight('bold')
            .setBackground('#B7860B')
            .setFontColor('#ffffff');
        hoja.setFrozenRows(1);
      }
      // Anti-duplicado por email
      var cert_data = hoja.getDataRange().getValues();
      for (var j = 1; j < cert_data.length; j++) {
        if (String(cert_data[j][1]) === String(datos.email)) {
          // Ya solicitó — solo actualizar total
          hoja.getRange(j+1, 5).setValue(datos.total_conferencias || '');
          hoja.getRange(j+1, 6).setValue(datos.conferencias || '');
          return ContentService
            .createTextOutput(JSON.stringify({ ok: true, actualizado: true }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
      hoja.appendRow([
        new Date(),
        datos.email              || '',
        datos.nombre_completo    || '',
        datos.especialidad       || '',
        datos.total_conferencias || '',
        datos.conferencias       || '',
        'PENDIENTE'
      ]);

    } else if (tipo === 'incidencia') {
      // ── Incidencias / Derivaciones del staff ──
      hoja = ss.getSheetByName('INCIDENCIAS');
      if (!hoja) {
        hoja = ss.insertSheet('INCIDENCIAS');
        var hi = ['Timestamp','Fecha','Día','Hora','Nombre','Especialidad',
                  'Pase','UID','Motivo','Puesto Staff'];
        hoja.appendRow(hi);
        hoja.getRange(1,1,1,hi.length)
            .setFontWeight('bold')
            .setBackground('#D4700A')
            .setFontColor('#ffffff');
        hoja.setFrozenRows(1);
      }
      hoja.appendRow([
        new Date(),
        datos.fecha        || '',
        datos.dia          || '',
        datos.hora         || '',
        datos.nombre       || '',
        datos.especialidad || '',
        datos.pase         || '',
        datos.uid          || '',
        datos.motivo       || '',
        datos.staffId      || ''
      ]);

    } else if (tipo === 'registro_asistente') {
      // ── Registro inicial del asistente (formulario app) ──
      hoja = ss.getSheetByName('REGISTROS');
      if (!hoja) {
        hoja = ss.insertSheet('REGISTROS');
        var h4 = ['Timestamp','Nombre','Apellido','Nombre Completo',
                  'Teléfono','Email','Especialidad','Días Autorizados'];
        hoja.appendRow(h4);
        hoja.getRange(1,1,1,h4.length)
            .setFontWeight('bold')
            .setBackground('#1565C8')
            .setFontColor('#ffffff');
        hoja.setFrozenRows(1);
      }
      hoja.appendRow([
        new Date(),
        datos.nombre          || '',
        datos.apellido        || '',
        datos.nombre_completo || '',
        datos.telefono        || '',
        datos.email           || '',
        datos.especialidad    || '',
        datos.dias_autorizados|| ''
      ]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Mantener también doGet para compatibilidad con el código anterior
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, mensaje: 'CHP 2026 API activa' }))
    .setMimeType(ContentService.MimeType.JSON);
}
