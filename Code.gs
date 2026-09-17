/**
 * Backend de Apps Script para la app de "Disponibilidad".
 *
 * DEPLOY:
 * 1. En https://script.google.com creá un proyecto NUEVO (standalone, no lo
 *    vincules a ninguna planilla).
 * 2. Pegá este archivo completo reemplazando el contenido de Code.gs.
 * 3. Implementar > Nueva implementación > Tipo: "Aplicación web".
 *      - Ejecutar como: Yo (tu cuenta)
 *      - Quién tiene acceso: Cualquier usuario
 * 4. Copiá la URL que te da ("Web app URL") y pegala en CONFIG.WEB_APP_URL
 *    dentro de index.html.
 * 5. Cada vez que edites este script, tenés que "Implementar > Gestionar
 *    implementaciones > Editar > Nueva versión" para que los cambios se
 *    reflejen en la URL pública.
 */

const SLOT_MINUTES = 30;

// Completá con tu propio mail de Google: solo ese mail va a poder ver el
// panel de "mis eventos" y cerrar/reabrir eventos.
const OWNER_EMAIL = 'PEGAR_TU_EMAIL_DE_GOOGLE_AQUI';

function doPost(e) {
  return handle_(e, true);
}

function doGet(e) {
  return handle_(e, false);
}

function handle_(e, isPost) {
  let action, params;
  try {
    if (isPost) {
      params = JSON.parse(e.postData.contents);
    } else {
      params = e.parameter;
    }
    action = params.action;
    let data;
    switch (action) {
      case 'createEvent':
        data = createEvent_(params);
        break;
      case 'getEvent':
        data = getEvent_(params.sheetId);
        break;
      case 'submitAvailability':
        data = submitAvailability_(params);
        break;
      case 'getAvailability':
        data = getAvailability_(params.sheetId, params.email);
        break;
      case 'deleteMyAvailability':
        data = deleteMyAvailability_(params.sheetId, params.email);
        break;
      case 'listEvents':
        data = listEvents_(params.email);
        break;
      case 'setEventClosed':
        data = setEventClosed_(params.sheetId, params.email, params.closed === true || params.closed === 'true');
        break;
      case 'deleteEvent':
        data = deleteEvent_(params.sheetId, params.email);
        break;
      case 'getSummary':
        data = getSummary_(params.sheetId);
        break;
      case 'listMyEvents':
        data = listMyEvents_(params.email);
        break;
      default:
        throw new Error('Acción desconocida: ' + action);
    }
    return jsonOut_({ ok: true, data: data });
  } catch (err) {
    return jsonOut_({ ok: false, error: err.message });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- Utilidades de fechas/horarios ----------

function dateRange_(startDate, endDate) {
  const days = [];
  let d = new Date(startDate + 'T00:00:00');
  const end = new Date(endDate + 'T00:00:00');
  while (d <= end) {
    days.push(Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    d.setDate(d.getDate() + 1);
  }
  return days;
}

function timeSlots_(startHour, endHour) {
  const slots = [];
  const sh = startHour.split(':').map(Number);
  const eh = endHour.split(':').map(Number);
  let mins = sh[0] * 60 + sh[1];
  const endMins = eh[0] * 60 + eh[1];
  while (mins < endMins) {
    const h = Math.floor(mins / 60);
    const m = mins % 60;
    slots.push(pad2_(h) + ':' + pad2_(m));
    mins += SLOT_MINUTES;
  }
  return slots;
}

function pad2_(n) {
  return (n < 10 ? '0' : '') + n;
}

function requireOwner_(email) {
  if (!email || email.trim().toLowerCase() !== OWNER_EMAIL.trim().toLowerCase()) {
    throw new Error('No autorizado.');
  }
}

// ---------- Índice de eventos (para el panel del dueño) ----------

function getIndexSheet_() {
  const props = PropertiesService.getScriptProperties();
  let indexId = props.getProperty('INDEX_SHEET_ID');
  let ss;
  if (indexId) {
    try {
      ss = SpreadsheetApp.openById(indexId);
      return ss.getSheetByName('Eventos');
    } catch (e) {
      // El índice fue borrado o movido; se vuelve a crear más abajo.
    }
  }
  ss = SpreadsheetApp.create('Índice de eventos — Disponibilidad');
  props.setProperty('INDEX_SHEET_ID', ss.getId());
  const sheet = ss.getSheets()[0];
  sheet.setName('Eventos');
  sheet.getRange('A1:H1').setValues([[
    'SheetId', 'Título', 'FechaInicio', 'FechaFin', 'URL Planilla', 'Link para compartir', 'CreadoEl', 'Cerrado'
  ]]);
  sheet.getRange('A1:H1').setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

function addToIndex_(entry) {
  const sheet = getIndexSheet_();
  sheet.appendRow([
    entry.sheetId, entry.title, entry.startDate, entry.endDate,
    entry.url, entry.shareUrl, entry.createdAt, false
  ]);
}

function listEvents_(email) {
  requireOwner_(email);
  const sheet = getIndexSheet_();
  const data = sheet.getDataRange().getValues();
  const events = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    events.push({
      sheetId: row[0],
      title: row[1],
      startDate: row[2],
      endDate: row[3],
      url: row[4],
      shareUrl: row[5],
      createdAt: row[6],
      closed: row[7] === true
    });
  }
  events.reverse(); // más reciente primero
  return { events: events };
}

function setEventClosed_(sheetId, email, closed) {
  requireOwner_(email);
  if (!sheetId) throw new Error('Falta sheetId.');

  const ss = SpreadsheetApp.openById(sheetId);
  const config = ss.getSheetByName('Config');
  const vals = config.getRange('A1:B7').getValues();
  let closedRow = -1;
  for (let i = 0; i < vals.length; i++) {
    if (vals[i][0] === 'Cerrado') { closedRow = i + 1; break; }
  }
  if (closedRow === -1) {
    config.getRange(7, 1, 1, 2).setValues([['Cerrado', closed]]);
  } else {
    config.getRange(closedRow, 2).setValue(closed);
  }

  // Reflejar el estado en el índice también.
  const sheet = getIndexSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === sheetId) {
      sheet.getRange(i + 1, 8).setValue(closed);
      break;
    }
  }
  return { closed: closed };
}

function deleteEvent_(sheetId, email) {
  requireOwner_(email);
  if (!sheetId) throw new Error('Falta sheetId.');

  try {
    DriveApp.getFileById(sheetId).setTrashed(true);
  } catch (e) {
    // Si ya no existe el archivo, seguimos igual para poder limpiar el índice.
  }

  const sheet = getIndexSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === sheetId) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return { deleted: true };
}

// ---------- Crear evento ----------

function createEvent_(params) {
  requireOwner_(params.email);
  const title = params.title || 'Disponibilidad';
  const startDate = params.startDate;
  const endDate = params.endDate;
  const startHour = params.startHour || '09:00';
  const endHour = params.endHour || '23:00';

  if (!startDate || !endDate) throw new Error('Falta el rango de fechas.');

  const ss = SpreadsheetApp.create(title + ' — Disponibilidad');
  const sheetId = ss.getId();

  // Config
  const config = ss.getSheets()[0];
  config.setName('Config');
  const createdAt = new Date().toISOString();
  config.getRange('A1:B7').setValues([
    ['Título', title],
    ['FechaInicio', startDate],
    ['FechaFin', endDate],
    ['HoraInicio', startHour],
    ['HoraFin', endHour],
    ['CreadoEl', createdAt],
    ['Cerrado', false]
  ]);
  config.autoResizeColumns(1, 2);

  // Respuestas (una fila por persona)
  const resp = ss.insertSheet('Respuestas');
  resp.getRange('A1:D1').setValues([['Nombre', 'Email', 'Slots', 'ÚltimaActualización']]);
  resp.setFrozenRows(1);
  resp.getRange('A1:D1').setFontWeight('bold');

  // Resumen (se completa/actualiza en cada respuesta)
  const resumen = ss.insertSheet('Resumen');
  resumen.getRange('A1').setValue('Todavía no hay respuestas.');

  addToIndex_({
    sheetId: sheetId,
    title: title,
    startDate: startDate,
    endDate: endDate,
    url: ss.getUrl(),
    shareUrl: '', // el frontend arma el link con su propio dominio
    createdAt: createdAt
  });

  return {
    sheetId: sheetId,
    url: ss.getUrl(),
    title: title,
    startDate: startDate,
    endDate: endDate,
    startHour: startHour,
    endHour: endHour
  };
}

// ---------- Leer evento ----------

function getEvent_(sheetId) {
  if (!sheetId) throw new Error('Falta sheetId.');
  const ss = SpreadsheetApp.openById(sheetId);
  const config = ss.getSheetByName('Config');
  const vals = config.getRange('A1:B7').getValues();
  const map = {};
  vals.forEach(function (row) { map[row[0]] = row[1]; });
  return {
    sheetId: sheetId,
    url: ss.getUrl(),
    title: map['Título'],
    startDate: map['FechaInicio'],
    endDate: map['FechaFin'],
    startHour: map['HoraInicio'],
    endHour: map['HoraFin'],
    closed: map['Cerrado'] === true,
    days: dateRange_(map['FechaInicio'], map['FechaFin']),
    slots: timeSlots_(map['HoraInicio'], map['HoraFin'])
  };
}

// ---------- Leer disponibilidad ya guardada de una persona ----------

function getAvailability_(sheetId, email) {
  if (!sheetId) throw new Error('Falta sheetId.');
  if (!email) throw new Error('Falta email.');
  email = email.trim().toLowerCase();

  const ss = SpreadsheetApp.openById(sheetId);
  const resp = ss.getSheetByName('Respuestas');
  const data = resp.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).toLowerCase() === email) {
      const slotsStr = String(data[i][2] || '');
      const slots = slotsStr.split('|').filter(function (s) { return s; });
      return { slots: slots };
    }
  }
  return { slots: [] };
}

// ---------- Guardar disponibilidad ----------

function submitAvailability_(params) {
  const sheetId = params.sheetId;
  const name = (params.name || '').trim();
  const email = (params.email || '').trim().toLowerCase();
  const slots = params.slots || []; // array de "YYYY-MM-DD HH:mm"

  if (!sheetId) throw new Error('Falta sheetId.');
  if (!name || !email) throw new Error('Falta nombre o email.');

  const ss = SpreadsheetApp.openById(sheetId);
  const config = ss.getSheetByName('Config');
  const cvals = config.getRange('A1:B7').getValues();
  const cmap = {};
  cvals.forEach(function (row) { cmap[row[0]] = row[1]; });
  if (cmap['Cerrado'] === true) {
    throw new Error('Este evento ya no acepta respuestas.');
  }

  const resp = ss.getSheetByName('Respuestas');
  const data = resp.getDataRange().getValues();

  const slotsStr = '|' + slots.join('|') + '|';
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).toLowerCase() === email) { rowIndex = i + 1; break; }
  }

  const now = new Date().toISOString();
  if (rowIndex === -1) {
    resp.appendRow([name, email, slotsStr, now]);
  } else {
    resp.getRange(rowIndex, 1, 1, 4).setValues([[name, email, slotsStr, now]]);
  }

  rebuildResumen_(ss);
  return { ok: true };
}

// ---------- Borrar la respuesta de una persona ----------

function deleteMyAvailability_(sheetId, email) {
  if (!sheetId) throw new Error('Falta sheetId.');
  if (!email) throw new Error('Falta email.');
  email = email.trim().toLowerCase();

  const ss = SpreadsheetApp.openById(sheetId);
  const resp = ss.getSheetByName('Respuestas');
  const data = resp.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).toLowerCase() === email) {
      resp.deleteRow(i + 1);
      rebuildResumen_(ss);
      return { deleted: true };
    }
  }
  return { deleted: false };
}

// ---------- Cálculo compartido del resumen ----------

function computeSummary_(ss) {
  const config = ss.getSheetByName('Config');
  const cvals = config.getRange('A1:B6').getValues();
  const map = {};
  cvals.forEach(function (row) { map[row[0]] = row[1]; });
  const days = dateRange_(map['FechaInicio'], map['FechaFin']);
  const slots = timeSlots_(map['HoraInicio'], map['HoraFin']);

  const resp = ss.getSheetByName('Respuestas');
  const data = resp.getDataRange().getValues();
  const people = [];
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    const slotsStr = String(data[i][2] || '');
    if (!name) continue;
    people.push({ name: name, slotsStr: slotsStr });
  }

  const totalPeople = people.length;
  const counts = [];
  const namesGrid = [];
  slots.forEach(function (slot, si) {
    counts.push([]);
    namesGrid.push([]);
    days.forEach(function (day) {
      const key = '|' + day + ' ' + slot + '|';
      const avail = people.filter(function (p) { return p.slotsStr.indexOf(key) !== -1; });
      counts[si].push(avail.length);
      namesGrid[si].push(avail.map(function (p) { return p.name; }).join(', '));
    });
  });

  const flat = [];
  slots.forEach(function (slot, si) {
    days.forEach(function (day, di) {
      flat.push({ day: day, slot: slot, count: counts[si][di], names: namesGrid[si][di] });
    });
  });
  flat.sort(function (a, b) { return b.count - a.count; });
  const top = flat.filter(function (x) { return x.count > 0; }).slice(0, 10);

  return { days: days, slots: slots, totalPeople: totalPeople, counts: counts, namesGrid: namesGrid, top: top };
}

function getSummary_(sheetId) {
  if (!sheetId) throw new Error('Falta sheetId.');
  const ss = SpreadsheetApp.openById(sheetId);
  return computeSummary_(ss);
}

// ---------- Eventos en los que participó una persona ----------

function listMyEvents_(email) {
  if (!email) throw new Error('Falta email.');
  email = email.trim().toLowerCase();

  const indexSheet = getIndexSheet_();
  const idx = indexSheet.getDataRange().getValues();
  const results = [];

  for (let i = 1; i < idx.length; i++) {
    const row = idx[i];
    if (!row[0]) continue;
    try {
      const ss = SpreadsheetApp.openById(row[0]);
      const resp = ss.getSheetByName('Respuestas');
      const data = resp.getDataRange().getValues();
      let mine = null;
      for (let j = 1; j < data.length; j++) {
        if (String(data[j][1]).toLowerCase() === email) { mine = data[j]; break; }
      }
      if (mine) {
        const slots = String(mine[2] || '').split('|').filter(function (s) { return s; });
        results.push({
          sheetId: row[0],
          title: row[1],
          startDate: row[2],
          endDate: row[3],
          closed: row[7] === true,
          mySlotsCount: slots.length
        });
      }
    } catch (e) {
      // La planilla ya no existe (fue borrada); se ignora.
    }
  }
  results.reverse();
  return { events: results };
}

// ---------- Reconstruir la hoja Resumen ----------

function rebuildResumen_(ss) {
  const summary = computeSummary_(ss);
  const days = summary.days;
  const slots = summary.slots;
  const totalPeople = summary.totalPeople;
  const counts = summary.counts;
  const top = summary.top;

  const resumen = ss.getSheetByName('Resumen');
  resumen.clear();

  if (totalPeople === 0) {
    resumen.getRange('A1').setValue('Todavía no hay respuestas.');
    return;
  }

  // Encabezado
  resumen.getRange(1, 1, 1, days.length + 1).setValues([['Horario'].concat(days)]);
  resumen.getRange(1, 1, 1, days.length + 1).setFontWeight('bold');

  // Grilla de conteos
  const grid = slots.map(function (slot, si) { return [slot].concat(counts[si]); });
  resumen.getRange(2, 1, grid.length, days.length + 1).setValues(grid);

  // Escala de color (blanco -> verde) según cuánta gente puede
  const dataRange = resumen.getRange(2, 2, slots.length, days.length);
  const backgrounds = counts.map(function (row) {
    return row.map(function (c) {
      const ratio = totalPeople ? c / totalPeople : 0;
      return interpolateColor_(ratio);
    });
  });
  dataRange.setBackgrounds(backgrounds);
  resumen.setFrozenRows(1);
  resumen.setFrozenColumns(1);

  // Top horarios con quiénes pueden
  const startCol = days.length + 3;
  resumen.getRange(1, startCol).setValue('Top horarios (' + totalPeople + ' respondieron)').setFontWeight('bold');
  resumen.getRange(2, startCol, 1, 3).setValues([['Día y hora', 'Cuántos', 'Quiénes']]);
  resumen.getRange(2, startCol, 1, 3).setFontWeight('bold');
  const topRows = top.map(function (x) {
    return [x.day + ' ' + x.slot, x.count + '/' + totalPeople, x.names];
  });
  if (topRows.length) {
    resumen.getRange(3, startCol, topRows.length, 3).setValues(topRows);
  }
  resumen.autoResizeColumns(1, days.length + 5);
}

function interpolateColor_(ratio) {
  // blanco (255,255,255) -> verde (46,125,50)
  const r = Math.round(255 + (46 - 255) * ratio);
  const g = Math.round(255 + (125 - 255) * ratio);
  const b = Math.round(255 + (50 - 255) * ratio);
  return rgbToHex_(r, g, b);
}

function rgbToHex_(r, g, b) {
  return '#' + [r, g, b].map(function (x) {
    const h = x.toString(16);
    return h.length === 1 ? '0' + h : h;
  }).join('');
}
