/**
 * SISTEMA DE REGISTRO - INACTIVOS HISTORIAS LABORALES 2025
 * Backend para Google Apps Script conectado a Google Sheets.
 *
 * Hoja principal: INACTIVOS
 * Spreadsheet:
 * https://docs.google.com/spreadsheets/d/1bqq7sTHbr6tCk2uHPW-WoOt-4u1kOe8B1UkWN7DzkxU/edit?gid=830378407
 *
 * Instrucciones rápidas:
 * 1. Abre script.google.com y crea un proyecto.
 * 2. Crea el archivo Code.gs y pega este contenido.
 * 3. Crea el archivo index.html y pega el otro archivo entregado.
 * 4. Implementa como Aplicación web:
 *    - Ejecutar como: Tú
 *    - Quién tiene acceso: Cualquiera con el enlace, o solo tu organización si aplica.
 * 5. Copia la URL de implementación y pégala en la constante API_URL del index.html si lo usarás desde VS Code.
 */

const CONFIG = {
  SPREADSHEET_ID: '1bqq7sTHbr6tCk2uHPW-WoOt-4u1kOe8B1UkWN7DzkxU',
  SHEET_NAME: 'INACTIVOS',
  HISTORY_SHEET: 'HISTORIAL',
  DELETE_PASSWORD: 'ELIMINAR',
  DRIVE_URL: 'https://docs.google.com/spreadsheets/d/1bqq7sTHbr6tCk2uHPW-WoOt-4u1kOe8B1UkWN7DzkxU/edit?gid=830378407#gid=830378407',
  HEADERS: [
    'ID',
    'ESTADO',
    'CAJA',
    'CARPETA',
    'IDENTIFICACION',
    'NOMBRES Y APELLIDOS',
    'TOMO',
    'OBSERVACIONES'
  ],
  HISTORY_HEADERS: [
    'FECHA',
    'ACCION',
    'ID',
    'ESTADO',
    'CAJA',
    'CARPETA',
    'IDENTIFICACION',
    'NOMBRES Y APELLIDOS',
    'TOMO',
    'OBSERVACIONES',
    'USUARIO',
    'DETALLE_JSON'
  ]
};

/**
 * Abre la interfaz si visitas la URL del Web App directamente.
 * También funciona como API cuando se recibe ?action=...
 */
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action) {
      return apiResponse_(handleAction_(e));
    }

    return HtmlService
      .createTemplateFromFile('index')
      .evaluate()
      .setTitle('Inactivos Historias Laborales 2025')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return apiResponse_(fail_(err));
  }
}

/**
 * Recibe operaciones desde el index.html.
 */
function doPost(e) {
  try {
    return apiResponse_(handleAction_(e));
  } catch (err) {
    return apiResponse_(fail_(err));
  }
}

/**
 * API principal.
 * Acciones:
 * - list
 * - get
 * - add
 * - update
 * - delete
 * - history
 * - restore
 * - dashboard
 * - ping
 */
function handleAction_(e) {
  ensureSheets_();

  const request = readRequest_(e);
  const action = String(request.action || 'list').toLowerCase();

  switch (action) {
    case 'ping':
      return ok_({
        message: 'Conexión correcta',
        app: 'Inactivos Historias Laborales 2025',
        spreadsheetUrl: CONFIG.DRIVE_URL,
        now: new Date().toISOString()
      });

    case 'list':
      return ok_({
        records: getRecords_(),
        summary: buildDashboard_()
      });

    case 'get':
      return getRecordById_(request.id);

    case 'add':
      return withLock_(function () {
        return addRecord_(request);
      });

    case 'update':
      return withLock_(function () {
        return updateRecord_(request);
      });

    case 'delete':
      return withLock_(function () {
        return deleteRecord_(request);
      });

    case 'history':
      return ok_({
        history: getHistory_(Number(request.limit || 500))
      });

    case 'restore':
      return withLock_(function () {
        return restoreRecord_(request);
      });

    case 'dashboard':
      return ok_(buildDashboard_());

    default:
      return fail_('Acción no reconocida: ' + action);
  }
}

/**
 * Lee parámetros de GET, POST form-urlencoded, FormData o JSON.
 */
function readRequest_(e) {
  const out = {};

  if (e && e.parameter) {
    Object.keys(e.parameter).forEach(function (key) {
      out[key] = e.parameter[key];
    });
  }

  if (e && e.postData && e.postData.contents) {
    const raw = e.postData.contents;
    const contentType = String(e.postData.type || '').toLowerCase();

    if (contentType.indexOf('application/json') !== -1) {
      try {
        const parsed = JSON.parse(raw);
        Object.keys(parsed || {}).forEach(function (key) {
          out[key] = parsed[key];
        });
      } catch (err) {
        // Continúa con lo que haya en e.parameter.
      }
    } else if (out.payload) {
      try {
        const parsedPayload = JSON.parse(out.payload);
        Object.keys(parsedPayload || {}).forEach(function (key) {
          out[key] = parsedPayload[key];
        });
      } catch (err) {
        // Continúa con lo que haya en e.parameter.
      }
    }
  }

  return out;
}

/**
 * Crea las hojas y encabezados necesarios si no existen.
 */
function ensureSheets_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }
  ensureHeaders_(sheet, CONFIG.HEADERS);

  let history = ss.getSheetByName(CONFIG.HISTORY_SHEET);
  if (!history) {
    history = ss.insertSheet(CONFIG.HISTORY_SHEET);
  }
  ensureHeaders_(history, CONFIG.HISTORY_HEADERS);

  return { ss: ss, sheet: sheet, history: history };
}

function ensureHeaders_(sheet, headers) {
  const lastColumn = Math.max(sheet.getLastColumn(), headers.length);
  const current = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  let changed = false;

  for (let i = 0; i < headers.length; i++) {
    if (String(current[i] || '').trim() !== headers[i]) {
      current[i] = headers[i];
      changed = true;
    }
  }

  if (changed || sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1f2937')
      .setFontColor('#ffffff');
    sheet.autoResizeColumns(1, headers.length);
  }
}

/**
 * Devuelve todos los registros de INACTIVOS como objetos.
 */
function getRecords_() {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  return values.slice(1)
    .filter(function (row) {
      return row.some(function (cell) {
        return String(cell).trim() !== '';
      });
    })
    .map(function (row) {
      return rowToRecord_(row);
    });
}

function rowToRecord_(row) {
  const record = {};
  CONFIG.HEADERS.forEach(function (header, index) {
    record[header] = normalizeCell_(row[index]);
  });
  return record;
}

function recordToRow_(record) {
  return CONFIG.HEADERS.map(function (header) {
    return normalizeCell_(record[header]);
  });
}

function normalizeCell_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function nextId_(records) {
  const maxId = records.reduce(function (max, item) {
    const id = Number(item.ID || 0);
    return id > max ? id : max;
  }, 0);
  return maxId + 1;
}

function findRowById_(sheet, id) {
  id = String(id || '').trim();
  if (!id) return -1;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1;

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === id) {
      return i + 2;
    }
  }
  return -1;
}

function getRecordById_(id) {
  const records = getRecords_();
  const found = records.find(function (item) {
    return String(item.ID) === String(id);
  });

  if (!found) {
    return fail_('No se encontró el registro con ID ' + id);
  }

  return ok_({ record: found });
}

function buildRecordFromRequest_(request, current) {
  const base = current || {};

  const record = {
    'ID': normalizeCell_(request.ID || request.id || base.ID),
    'ESTADO': upper_(request.ESTADO || request.estado || base.ESTADO || 'INACTIVO'),
    'CAJA': normalizeCell_(request.CAJA || request.caja || base.CAJA),
    'CARPETA': normalizeCell_(request.CARPETA || request.carpeta || base.CARPETA),
    'IDENTIFICACION': normalizeCell_(request.IDENTIFICACION || request.identificacion || request.documento || base.IDENTIFICACION),
    'NOMBRES Y APELLIDOS': upper_(request['NOMBRES Y APELLIDOS'] || request.nombres || request.nombre || base['NOMBRES Y APELLIDOS']),
    'TOMO': normalizeTomo_(request.TOMO || request.tomo || base.TOMO),
    'OBSERVACIONES': normalizeCell_(request.OBSERVACIONES || request.observaciones || base.OBSERVACIONES)
  };

  return record;
}

function validateRecord_(record, options) {
  options = options || {};
  const required = ['ESTADO', 'CAJA', 'CARPETA', 'IDENTIFICACION', 'NOMBRES Y APELLIDOS'];
  const missing = required.filter(function (field) {
    return !String(record[field] || '').trim();
  });

  if (missing.length) {
    throw new Error('Faltan campos obligatorios: ' + missing.join(', '));
  }

  if (record.ESTADO !== 'ACTIVO' && record.ESTADO !== 'INACTIVO') {
    throw new Error('El estado debe ser ACTIVO o INACTIVO.');
  }

  if (!/^\d+$/.test(String(record.CAJA))) {
    throw new Error('La CAJA debe ser numérica.');
  }

  if (!/^\d+$/.test(String(record.CARPETA))) {
    throw new Error('La CARPETA debe ser numérica.');
  }

  if (!/^\d+$/.test(String(record.IDENTIFICACION))) {
    throw new Error('La IDENTIFICACIÓN debe ser numérica.');
  }

  if (!record.TOMO) {
    record.TOMO = 'NO REGISTRA TOMOS';
  }

  if (!options.allowExistingId && record.ID) {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    if (findRowById_(sheet, record.ID) !== -1) {
      throw new Error('Ya existe un registro con ID ' + record.ID);
    }
  }

  return true;
}

function addRecord_(request) {
  const env = ensureSheets_();
  const records = getRecords_();
  const record = buildRecordFromRequest_(request);

  if (!record.ID) {
    record.ID = String(nextId_(records));
  }

  validateRecord_(record, { allowExistingId: false });

  env.sheet.appendRow(recordToRow_(record));
  logHistory_('AGREGAR', record.ID, record, { nuevo: record });

  return ok_({
    message: 'Registro agregado correctamente.',
    record: record,
    records: getRecords_(),
    summary: buildDashboard_()
  });
}

function updateRecord_(request) {
  const env = ensureSheets_();
  const rowNumber = findRowById_(env.sheet, request.id || request.ID);
  if (rowNumber === -1) {
    return fail_('No se encontró el registro para actualizar.');
  }

  const currentRow = env.sheet.getRange(rowNumber, 1, 1, CONFIG.HEADERS.length).getValues()[0];
  const current = rowToRecord_(currentRow);
  const updated = buildRecordFromRequest_(request, current);
  updated.ID = current.ID;

  validateRecord_(updated, { allowExistingId: true });

  env.sheet.getRange(rowNumber, 1, 1, CONFIG.HEADERS.length).setValues([recordToRow_(updated)]);
  logHistory_('EDITAR', updated.ID, updated, { antes: current, despues: updated });

  return ok_({
    message: 'Registro actualizado correctamente.',
    record: updated,
    records: getRecords_(),
    summary: buildDashboard_()
  });
}

function deleteRecord_(request) {
  requireDeletePassword_(request.clave || request.password);

  const env = ensureSheets_();
  const id = request.id || request.ID;
  const rowNumber = findRowById_(env.sheet, id);

  if (rowNumber === -1) {
    return fail_('No se encontró el registro para eliminar.');
  }

  const row = env.sheet.getRange(rowNumber, 1, 1, CONFIG.HEADERS.length).getValues()[0];
  const record = rowToRecord_(row);

  env.sheet.deleteRow(rowNumber);
  logHistory_('ELIMINAR', record.ID, record, {
    eliminado: record,
    nota: 'Registro eliminado desde la interfaz. Puede restaurarse desde Historial.'
  });

  return ok_({
    message: 'Registro eliminado correctamente.',
    deleted: record,
    records: getRecords_(),
    summary: buildDashboard_()
  });
}

function getHistory_(limit) {
  const history = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.HISTORY_SHEET);
  const values = history.getDataRange().getValues();

  if (values.length <= 1) return [];

  const rows = values.slice(1).filter(function (row) {
    return row.some(function (cell) {
      return String(cell).trim() !== '';
    });
  });

  const mapped = rows.map(function (row) {
    const item = {};
    CONFIG.HISTORY_HEADERS.forEach(function (header, index) {
      let value = row[index];
      if (header === 'FECHA' && value instanceof Date) {
        value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      }
      item[header] = normalizeCell_(value);
    });
    return item;
  }).reverse();

  if (limit && mapped.length > limit) {
    return mapped.slice(0, limit);
  }

  return mapped;
}

function restoreRecord_(request) {
  requireDeletePassword_(request.clave || request.password);

  const id = String(request.id || request.ID || '').trim();
  if (!id) {
    return fail_('Debes indicar el ID que deseas restaurar.');
  }

  const env = ensureSheets_();

  if (findRowById_(env.sheet, id) !== -1) {
    return fail_('El ID ' + id + ' ya existe en la tabla principal. No se restauró para evitar duplicados.');
  }

  const historyRecords = getHistory_(1000);
  const target = historyRecords.find(function (item) {
    return String(item.ID) === id && String(item.ACCION).toUpperCase() === 'ELIMINAR';
  });

  if (!target) {
    return fail_('No se encontró en el historial un registro eliminado con ID ' + id);
  }

  const restored = {
    'ID': target.ID,
    'ESTADO': target.ESTADO,
    'CAJA': target.CAJA,
    'CARPETA': target.CARPETA,
    'IDENTIFICACION': target.IDENTIFICACION,
    'NOMBRES Y APELLIDOS': target['NOMBRES Y APELLIDOS'],
    'TOMO': target.TOMO,
    'OBSERVACIONES': target.OBSERVACIONES
  };

  validateRecord_(restored, { allowExistingId: true });

  env.sheet.appendRow(recordToRow_(restored));
  logHistory_('RESTAURAR', restored.ID, restored, {
    restaurado: restored,
    desdeHistorial: target
  });

  return ok_({
    message: 'Registro restaurado correctamente.',
    record: restored,
    records: getRecords_(),
    summary: buildDashboard_()
  });
}

function logHistory_(action, id, record, detail) {
  const env = ensureSheets_();
  const user = getUserEmail_();
  const row = [
    new Date(),
    action,
    record.ID || id || '',
    record.ESTADO || '',
    record.CAJA || '',
    record.CARPETA || '',
    record.IDENTIFICACION || '',
    record['NOMBRES Y APELLIDOS'] || '',
    record.TOMO || '',
    record.OBSERVACIONES || '',
    user,
    JSON.stringify(detail || {})
  ];

  env.history.appendRow(row);
}

function buildDashboard_() {
  const records = getRecords_();

  const estado = countBy_(records, 'ESTADO');
  const caja = countBy_(records, 'CAJA');
  const tomo = countBy_(records, 'TOMO');

  const duplicadosIdentificacion = buildDuplicates_(records, 'IDENTIFICACION');
  const sinObservacion = records.filter(function (item) {
    return !String(item.OBSERVACIONES || '').trim();
  }).length;

  const cajas = Object.keys(caja).map(Number).filter(function (n) {
    return !isNaN(n);
  });

  return {
    total: records.length,
    activos: estado.ACTIVO || 0,
    inactivos: estado.INACTIVO || 0,
    sinObservacion: sinObservacion,
    estado: estado,
    caja: caja,
    tomo: tomo,
    duplicadosIdentificacion: duplicadosIdentificacion,
    primeraCaja: cajas.length ? Math.min.apply(null, cajas) : '',
    ultimaCaja: cajas.length ? Math.max.apply(null, cajas) : '',
    ultimaActualizacion: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
  };
}

function countBy_(records, field) {
  return records.reduce(function (acc, item) {
    const key = String(item[field] || 'SIN DATO').trim() || 'SIN DATO';
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
}

function buildDuplicates_(records, field) {
  const map = {};
  records.forEach(function (item) {
    const key = String(item[field] || '').trim();
    if (!key) return;
    if (!map[key]) map[key] = [];
    map[key].push(item.ID);
  });

  return Object.keys(map)
    .filter(function (key) {
      return map[key].length > 1;
    })
    .map(function (key) {
      return {
        valor: key,
        ids: map[key],
        cantidad: map[key].length
      };
    });
}

function requireDeletePassword_(clave) {
  if (String(clave || '').trim() !== CONFIG.DELETE_PASSWORD) {
    throw new Error('Clave incorrecta. Para eliminar o restaurar debes escribir la clave autorizada.');
  }
}

function normalizeTomo_(value) {
  value = normalizeCell_(value);
  if (!value) return 'NO REGISTRA TOMOS';
  return upper_(value);
}

function upper_(value) {
  return normalizeCell_(value).toUpperCase();
}

function getUserEmail_() {
  try {
    return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
  } catch (err) {
    return '';
  }
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function ok_(data) {
  return Object.assign({ ok: true }, data || {});
}

function fail_(err) {
  const message = err && err.message ? err.message : String(err || 'Error desconocido');
  return {
    ok: false,
    error: message
  };
}

/**
 * Respuesta JSON.
 * Nota: si usas index.html desde VS Code y tu navegador bloquea CORS,
 * abre directamente la URL de la Aplicación web, porque este Code.gs también sirve la interfaz.
 */
function apiResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
