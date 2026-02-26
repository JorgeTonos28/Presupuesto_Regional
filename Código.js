/***************
 * App Segmentación Presupuesto Regional
 * Versión: 1.0.5
 ***************/
const APP_VERSION = '1.0.5';

const SHEET_CONFIG = 'Config';
const SHEET_USERS = 'Usuarios';
const SHEET_BASE = 'BaseDatos';
const SHEET_SEG = 'Segmentaciones';
const SHEET_SEG_D = 'SegmentacionesDetalle';

const ROLE_ADMIN = 'ADMIN';
const ROLE_COLAB = 'COLAB';

function doGet(e) {
  try {
    const userRes = getSessionUser_();
    if (!userRes.ok) {
      const denied = HtmlService.createTemplateFromFile('Denied');
      denied.message = userRes.message || 'Acceso denegado.';
      return denied.evaluate()
        .setTitle('Acceso denegado')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    const tpl = HtmlService.createTemplateFromFile('Index');
    tpl.APP_VERSION = APP_VERSION;
    tpl.SESSION_USER = userRes.data;

    return tpl.evaluate()
      .setTitle('Segmentación de Presupuesto')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    Logger.log('doGet error: ' + err);
    const denied = HtmlService.createTemplateFromFile('Denied');
    denied.message = 'Error cargando la app: ' + (err && err.message ? err.message : err);
    return denied.evaluate().setTitle('Error');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Ejecutar 1 vez al inicio (y cuando falten pestañas).
 */
function setup() {
  const ss = SpreadsheetApp.getActive();

  // Config
  ensureSheet_(ss, SHEET_CONFIG, ['key', 'value']);
  ensureConfigDefaults_(ss);

  // Usuarios
  ensureSheet_(ss, SHEET_USERS, ['email', 'name', 'role', 'status', 'createdAt']);

  // BaseDatos (Año + estructura dada)
  ensureSheet_(ss, SHEET_BASE, [
    'Año', 'Regional', 'Departamento', 'Estrategia', 'Subprograma', 'Acción Formativa',
    'Modalidad', 'Provincia', 'Horas', 'Acciones', 'Ptes', 'Presupuesto'
  ]);

  // Segmentaciones (cabecera)
  ensureSheet_(ss, SHEET_SEG, [
    'segId', 'createdAt', 'year', 'pct', 'label', 'createdBy', 'status', 'deletedAt', 'deletedBy'
  ]);

  // SegmentacionesDetalle (por regional)
  ensureSheet_(ss, SHEET_SEG_D, [
    'segId', 'year', 'regional', 'remainingBefore', 'amountSegmented', 'remainingAfter', 'status'
  ]);

  // Bootstrap: el usuario que corre setup queda ADMIN/ACTIVE si no existe
  const email = getActiveEmail_();
  if (email) {
    const users = ss.getSheetByName(SHEET_USERS);
    const map = getHeaderMap_(users);
    const existingRow = findRowByValue_(users, map.email, email);
    if (!existingRow) {
      users.appendRow([email, '', ROLE_ADMIN, 'ACTIVE', new Date()]);
    }
  }

  SpreadsheetApp.flush();
}

function apiBootstrap() {
  try {
    const userRes = getSessionUser_();
    if (!userRes) return fail_('Error interno: getSessionUser_ retornó nulo.');
    if (!userRes.ok) return userRes;

    const cfg = getConfig_();
    const years = listYears_();

    // Intentar normalizar URLs
    let logoUrl = '';
    let signatureUrl = '';
    try { logoUrl = normalizeDriveUrl_(cfg.logo_url); } catch(e) { Logger.log('Logo err: ' + e); }
    try { signatureUrl = normalizeDriveUrl_(cfg.signature_url); } catch(e) { Logger.log('Sig err: ' + e); }

    const response = {
      appVersion: APP_VERSION,
      user: userRes.data,
      years,
      config: {
        locale: cfg.locale || 'es-DO',
        currencyCode: cfg.currency_code || '',
        logoUrl: logoUrl,
        signatureUrl: signatureUrl,
        logoWidth: cfg.logo_width || 'auto',
        logoHeight: cfg.logo_height || '40px',
        signatureWidth: cfg.signature_width || 'auto',
        signatureHeight: cfg.signature_height || '40px'
      }
    };

    return ok_(response);
  } catch (err) {
    Logger.log('apiBootstrap error: ' + err);
    return fail_('No se pudo inicializar la app: ' + (err.message || String(err)));
  }
}

function apiGetDashboard(payload) {
  try {
    const userRes = getSessionUser_();
    if (!userRes) return fail_('Error interno: getSessionUser_ retornó nulo.');
    if (!userRes.ok) return userRes;

    const year = parseInt(payload && payload.year, 10);
    if (!year) return fail_('Año inválido.');

    // Cache ligero (se invalida al crear/borrar segmentación)
    const cache = CacheService.getScriptCache();
    const key = 'dash_' + year;
    const cached = cache.get(key);
    if (cached) return ok_(JSON.parse(cached));

    const base = getBaseTotalsByRegional_(year);
    const seg = getSegmentedTotalsByRegional_(year);

    const regionals = Object.keys(base.byRegional).sort().map(r => {
      const baseVal = base.byRegional[r] || 0;
      const segVal = seg.byRegional[r] || 0;
      const remaining = round2_(baseVal - segVal);
      return { regional: r, base: baseVal, segmented: segVal, remaining };
    });

    const totalBase = round2_(base.total);
    const totalSegmented = round2_(seg.total);
    const totalRemaining = round2_(totalBase - totalSegmented);

    const segList = listSegmentations_(year); // incluye totales por segId

    const data = {
      year,
      totals: { base: totalBase, segmented: totalSegmented, remaining: totalRemaining },
      regionals,
      segmentations: segList
    };

    cache.put(key, JSON.stringify(data), 60); // 1 min
    return ok_(data);
  } catch (err) {
    Logger.log('apiGetDashboard error: ' + err);
    return fail_('No se pudo cargar el dashboard: ' + err.message);
  }
}

function apiCreateSegmentation(payload) {
  try {
    const userRes = getSessionUser_();
    if (!userRes) return fail_('Error interno: getSessionUser_ retornó nulo.');
    if (!userRes.ok) return userRes;
    requireAdmin_(userRes.data);

    const year = parseInt(payload && payload.year, 10);
    const pct = parseFloat(payload && payload.pct);
    const label = String((payload && payload.label) || '').trim();

    if (!year) return fail_('Año inválido.');
    if (!(pct > 0 && pct <= 100)) return fail_('Porcentaje debe ser > 0 y <= 100.');

    // Estado actual
    const base = getBaseTotalsByRegional_(year);
    const seg = getSegmentedTotalsByRegional_(year);

    const remainingByRegional = {};
    Object.keys(base.byRegional).forEach(r => {
      remainingByRegional[r] = round2_((base.byRegional[r] || 0) - (seg.byRegional[r] || 0));
    });

    const totalRemaining = round2_(base.total - seg.total);
    if (!(totalRemaining > 0)) return fail_('No hay disponible por presupuestar en este año.');

    const targetTotal = round2_(totalRemaining * (pct / 100));

    // Distribución proporcional por regional sobre su disponible
    const dist = distributeByRemaining_(remainingByRegional, targetTotal);

    const ss = SpreadsheetApp.getActive();
    const shSeg = ss.getSheetByName(SHEET_SEG);
    const shDet = ss.getSheetByName(SHEET_SEG_D);

    const segId = Utilities.getUuid();
    const now = new Date();
    const createdBy = userRes.data.email;

    shSeg.appendRow([segId, now, year, pct, label, createdBy, 'ACTIVE', '', '']);

    // Detalles
    const rows = [];
    Object.keys(dist).sort().forEach(regional => {
      const remainingBefore = remainingByRegional[regional] || 0;
      const amount = dist[regional] || 0;
      const remainingAfter = round2_(remainingBefore - amount);
      rows.push([segId, year, regional, remainingBefore, amount, remainingAfter, 'ACTIVE']);
    });

    if (rows.length) {
      shDet.getRange(shDet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    // Invalida cache
    CacheService.getScriptCache().remove('dash_' + year);

    SpreadsheetApp.flush();
    return ok_({ segId, year });
  } catch (err) {
    Logger.log('apiCreateSegmentation error: ' + err);
    return fail_('No se pudo crear la segmentación: ' + err.message);
  }
}

function apiDeleteSegmentation(payload) {
  try {
    const userRes = getSessionUser_();
    if (!userRes) return fail_('Error interno: getSessionUser_ retornó nulo.');
    if (!userRes.ok) return userRes;
    requireAdmin_(userRes.data);

    const segId = String(payload && payload.segId || '').trim();
    const year = parseInt(payload && payload.year, 10);
    if (!segId) return fail_('segId inválido.');
    if (!year) return fail_('Año inválido.');

    const ss = SpreadsheetApp.getActive();
    const shSeg = ss.getSheetByName(SHEET_SEG);
    const shDet = ss.getSheetByName(SHEET_SEG_D);

    const segMap = getHeaderMap_(shSeg);
    const detMap = getHeaderMap_(shDet);

    const row = findRowByValue_(shSeg, segMap.segId, segId);
    if (!row) return fail_('No existe esa segmentación.');

    // Marcar DELETED
    shSeg.getRange(row, segMap.status + 1).setValue('DELETED');
    shSeg.getRange(row, segMap.deletedAt + 1).setValue(new Date());
    shSeg.getRange(row, segMap.deletedBy + 1).setValue(userRes.data.email);

    // Marcar detalles DELETED
    const last = shDet.getLastRow();
    if (last >= 2) {
      const values = shDet.getRange(2, 1, last - 1, shDet.getLastColumn()).getValues();
      for (let i = 0; i < values.length; i++) {
        const id = String(values[i][detMap.segId] || '');
        if (id === segId) {
          shDet.getRange(i + 2, detMap.status + 1).setValue('DELETED');
        }
      }
    }

    CacheService.getScriptCache().remove('dash_' + year);
    SpreadsheetApp.flush();
    return ok_({ segId, year });
  } catch (err) {
    Logger.log('apiDeleteSegmentation error: ' + err);
    return fail_('No se pudo eliminar la segmentación: ' + err.message);
  }
}

function apiGetRegionalDetail(payload) {
  try {
    const userRes = getSessionUser_();
    if (!userRes) return fail_('Error interno: getSessionUser_ retornó nulo.');
    if (!userRes.ok) return userRes;

    const year = parseInt(payload && payload.year, 10);
    const regional = String(payload && payload.regional || '').trim();
    // Removed pagination params to return full dataset for client-side filtering

    if (!year) return fail_('Año inválido.');
    if (!regional) return fail_('Regional inválida.');

    // Totales regionales
    const base = getBaseTotalsByRegional_(year);
    const seg = getSegmentedTotalsByRegional_(year);

    const baseVal = base.byRegional[regional] || 0;
    const segVal = seg.byRegional[regional] || 0;
    const remaining = round2_(baseVal - segVal);

    // Segmentaciones de esa regional
    const segRows = listRegionalSegmentationRows_(year, regional);

    // Filas base (todas, hasta 5000)
    const baseRows = listBaseRows_(year, regional, 0, 5000);

    return ok_({
      year,
      regional,
      totals: { base: baseVal, segmented: segVal, remaining },
      segmentations: segRows,
      baseTable: baseRows
    });
  } catch (err) {
    Logger.log('apiGetRegionalDetail error: ' + err);
    return fail_('No se pudo cargar detalle regional: ' + err.message);
  }
}

function apiGetSegmentationDetail(payload) {
  try {
    const userRes = getSessionUser_();
    if (!userRes) return fail_('Error interno: getSessionUser_ retornó nulo.');
    if (!userRes.ok) return userRes;

    const segId = String(payload && payload.segId || '').trim();
    if (!segId) return fail_('segId inválido.');

    const ss = SpreadsheetApp.getActive();
    const shDet = ss.getSheetByName(SHEET_SEG_D);
    const map = getHeaderMap_(shDet);
    const last = shDet.getLastRow();
    if (last < 2) return ok_({ rows: [], total: 0 });

    const values = shDet.getRange(2, 1, last - 1, shDet.getLastColumn()).getValues();
    const rows = [];
    let total = 0;

    values.forEach(r => {
        const id = String(r[map.segId] || '');
        if (id !== segId) return;

        const status = String(r[map.status] || '');
        if (status !== 'ACTIVE') return;

        const regional = String(r[map.regional] || '');
        const amount = toNumber_(r[map.amountSegmented]);

        rows.push({ regional, amount });
        total = round2_(total + amount);
    });

    // Sort by regional name
    rows.sort((a, b) => a.regional.localeCompare(b.regional));

    return ok_({ rows, total });
  } catch (err) {
      Logger.log('apiGetSegmentationDetail error: ' + err);
      return fail_('Error: ' + err.message);
  }
}

/***************
 * ADMIN: Usuarios
 ***************/
function apiListUsers() {
  try {
    const userRes = getSessionUser_();
    if (!userRes) return fail_('Error interno: getSessionUser_ retornó nulo.');
    if (!userRes.ok) return userRes;
    requireAdmin_(userRes.data);

    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_USERS);
    const map = getHeaderMap_(sh);
    const last = sh.getLastRow();
    if (last < 2) return ok_([]);

    const vals = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    const users = vals.map(r => ({
      email: String(r[map.email] || ''),
      name: String(r[map.name] || ''),
      role: String(r[map.role] || ''),
      status: String(r[map.status] || ''),
      createdAt: safeDate_(r[map.createdAt])
    })).filter(u => u.email);

    return ok_(users);
  } catch (err) {
    Logger.log('apiListUsers error: ' + err);
    return fail_('No se pudo listar usuarios: ' + err.message);
  }
}

function apiUpsertUser(payload) {
  try {
    const userRes = getSessionUser_();
    if (!userRes) return fail_('Error interno: getSessionUser_ retornó nulo.');
    if (!userRes.ok) return userRes;
    requireAdmin_(userRes.data);

    const email = String(payload && payload.email || '').trim().toLowerCase();
    const name = String(payload && payload.name || '').trim();
    const role = String(payload && payload.role || '').trim().toUpperCase();
    const status = String(payload && payload.status || 'ACTIVE').trim().toUpperCase();

    if (!email || email.indexOf('@') === -1) return fail_('Email inválido.');
    if (![ROLE_ADMIN, ROLE_COLAB].includes(role)) return fail_('Rol inválido.');
    if (!['ACTIVE', 'INACTIVE'].includes(status)) return fail_('Status inválido.');

    // Validación opcional de dominio
    const cfg = getConfig_();
    const allowedDomain = (cfg.allowed_domain || '').trim().toLowerCase();
    if (allowedDomain) {
      const domain = email.split('@')[1] || '';
      if (domain !== allowedDomain) return fail_('El correo no pertenece al dominio permitido: ' + allowedDomain);
    }

    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_USERS);
    const map = getHeaderMap_(sh);
    const row = findRowByValue_(sh, map.email, email);

    if (row) {
      sh.getRange(row, map.name + 1).setValue(name);
      sh.getRange(row, map.role + 1).setValue(role);
      sh.getRange(row, map.status + 1).setValue(status);
    } else {
      sh.appendRow([email, name, role, status, new Date()]);
    }

    SpreadsheetApp.flush();
    return ok_({ email });
  } catch (err) {
    Logger.log('apiUpsertUser error: ' + err);
    return fail_('No se pudo guardar usuario: ' + err.message);
  }
}

/***************
 * Helpers (Auth / Config / Data)
 ***************/
function getSessionUser_() {
  try {
    const email = getActiveEmail_();
    if (!email) return fail_('No se pudo determinar tu correo institucional (inicia sesión con tu cuenta del dominio).');

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET_USERS);
    if (!sh) return fail_('Falta la pestaña Usuarios. Ejecuta setup().');

    const map = getHeaderMap_(sh);
    const row = findRowByValue_(sh, map.email, email);
    if (!row) return fail_('Tu usuario no está autorizado: ' + email);

    const vals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
    const user = {
      email: String(vals[map.email] || ''),
      name: String(vals[map.name] || ''),
      role: String(vals[map.role] || ''),
      status: String(vals[map.status] || '')
    };

    if (user.status !== 'ACTIVE') return fail_('Tu usuario está inactivo: ' + email);

    return ok_(user);
  } catch (err) {
    Logger.log('getSessionUser_ error: ' + err);
    return fail_('Error validando sesión: ' + (err.message || String(err)));
  }
}

function requireAdmin_(user) {
  if (!user || user.role !== ROLE_ADMIN) {
    throw new Error('No tienes permisos de Administrador.');
  }
}

function getActiveEmail_() {
  try {
    return (Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  } catch (e) {
    return '';
  }
}

function getConfig_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_CONFIG);
  if (!sh) return {};
  const last = sh.getLastRow();
  if (last < 2) return {};
  const vals = sh.getRange(2, 1, last - 1, 2).getValues();
  const cfg = {};
  vals.forEach(r => {
    const k = String(r[0] || '').trim();
    const v = String(r[1] || '').trim();
    if (k) cfg[k] = v;
  });
  return cfg;
}

function ensureConfigDefaults_(ss) {
  const sh = ss.getSheetByName(SHEET_CONFIG);
  const existing = getConfig_();
  const defaults = [
    ['allowed_domain', ''],     // ej: midominio.gob.do
    ['locale', 'es-DO'],
    ['currency_code', ''],       // ej: DOP (opcional)
    ['logo_url', ''],
    ['signature_url', ''],
    ['logo_width', 'auto'],
    ['logo_height', '40px'],
    ['signature_width', 'auto'],
    ['signature_height', '40px']
  ];
  defaults.forEach(([k, v]) => {
    if (existing[k] === undefined) sh.appendRow([k, v]);
  });
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  // Encabezados si no existen
  const range = sh.getRange(1, 1, 1, headers.length);
  const current = range.getValues()[0].map(String);
  const empty = current.every(c => !c || !String(c).trim());
  if (empty) {
    range.setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
}

function getHeaderMap_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const map = {};
  headers.forEach((h, i) => { if (h) map[h] = i; });
  return map;
}

function findRowByValue_(sh, colIndex0, value) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, colIndex0 + 1, last - 1, 1).getValues();
  const needle = String(value || '').trim().toLowerCase();
  for (let i = 0; i < vals.length; i++) {
    const v = String(vals[i][0] || '').trim().toLowerCase();
    if (v === needle) return i + 2;
  }
  return 0;
}

function listYears_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_BASE);
  if (!sh) return [new Date().getFullYear()];
  const last = sh.getLastRow();
  if (last < 2) return [new Date().getFullYear()];

  const map = getHeaderMap_(sh);
  const idxYear = map['Año'];
  if (idxYear === undefined) return [new Date().getFullYear()];

  const vals = sh.getRange(2, idxYear + 1, last - 1, 1).getValues();
  const set = {};
  vals.forEach(r => {
    const y = parseInt(r[0], 10);
    if (y) set[y] = true;
  });
  const years = Object.keys(set).map(n => parseInt(n, 10)).sort((a, b) => b - a);
  return years.length ? years : [new Date().getFullYear()];
}

function getBaseTotalsByRegional_(year) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_BASE);
  if (!sh) throw new Error('Falta BaseDatos. Ejecuta setup().');

  const last = sh.getLastRow();
  if (last < 2) return { total: 0, byRegional: {} };

  const map = getHeaderMap_(sh);
  const idxYear = map['Año'];
  const idxRegional = map['Regional'];
  const idxPres = map['Presupuesto'];

  if (idxRegional === undefined || idxPres === undefined) {
    throw new Error('BaseDatos debe tener columnas Regional y Presupuesto.');
  }
  if (idxYear === undefined) {
    throw new Error('BaseDatos debe tener columna Año para el selector.');
  }

  const values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  const byRegional = {};
  let total = 0;

  values.forEach(r => {
    const y = parseInt(r[idxYear], 10);
    if (y !== year) return;

    const reg = String(r[idxRegional] || '').trim();
    if (!reg) return;

    const pres = toNumber_(r[idxPres]);
    if (!byRegional[reg]) byRegional[reg] = 0;
    byRegional[reg] = round2_(byRegional[reg] + pres);
    total = round2_(total + pres);
  });

  return { total, byRegional };
}

function getSegmentedTotalsByRegional_(year) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SEG_D);
  if (!sh) throw new Error('Falta SegmentacionesDetalle. Ejecuta setup().');

  const last = sh.getLastRow();
  if (last < 2) return { total: 0, byRegional: {} };

  const map = getHeaderMap_(sh);
  const idxYear = map.year;
  const idxRegional = map.regional;
  const idxAmt = map.amountSegmented;
  const idxStatus = map.status;

  const values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  const byRegional = {};
  let total = 0;

  values.forEach(r => {
    const y = parseInt(r[idxYear], 10);
    if (y !== year) return;
    const st = String(r[idxStatus] || '').trim();
    if (st !== 'ACTIVE') return;

    const reg = String(r[idxRegional] || '').trim();
    if (!reg) return;

    const amt = toNumber_(r[idxAmt]);
    if (!byRegional[reg]) byRegional[reg] = 0;
    byRegional[reg] = round2_(byRegional[reg] + amt);
    total = round2_(total + amt);
  });

  return { total, byRegional };
}

function listSegmentations_(year) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_SEG);
  const shDet = ss.getSheetByName(SHEET_SEG_D);

  const segLast = sh.getLastRow();
  if (segLast < 2) return [];

  const segMap = getHeaderMap_(sh);
  const segVals = sh.getRange(2, 1, segLast - 1, sh.getLastColumn()).getValues();

  // Totales por segId desde detalle
  const detLast = shDet.getLastRow();
  const totalsById = {};
  if (detLast >= 2) {
    const detMap = getHeaderMap_(shDet);
    const detVals = shDet.getRange(2, 1, detLast - 1, shDet.getLastColumn()).getValues();
    detVals.forEach(r => {
      const y = parseInt(r[detMap.year], 10);
      if (y !== year) return;
      const st = String(r[detMap.status] || '');
      if (st !== 'ACTIVE') return;
      const id = String(r[detMap.segId] || '');
      const amt = toNumber_(r[detMap.amountSegmented]);
      if (!totalsById[id]) totalsById[id] = 0;
      totalsById[id] = round2_(totalsById[id] + amt);
    });
  }

  const list = [];
  segVals.forEach(r => {
    const y = parseInt(r[segMap.year], 10);
    if (y !== year) return;
    const st = String(r[segMap.status] || '');
    if (st !== 'ACTIVE') return;

    const id = String(r[segMap.segId] || '');
    list.push({
      segId: id,
      createdAt: r[segMap.createdAt],
      pct: toNumber_(r[segMap.pct]),
      label: String(r[segMap.label] || ''),
      createdBy: String(r[segMap.createdBy] || ''),
      totalSegmented: totalsById[id] || 0
    });
  });

  // Más reciente primero
  list.sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());

  // Convert Dates to strings for safe serialization
  return list.map(item => ({
    ...item,
    createdAt: safeDate_(item.createdAt)
  }));
}

function listRegionalSegmentationRows_(year, regional) {
  const ss = SpreadsheetApp.getActive();
  const shSeg = ss.getSheetByName(SHEET_SEG);
  const shDet = ss.getSheetByName(SHEET_SEG_D);

  const segMap = getHeaderMap_(shSeg);
  const segLast = shSeg.getLastRow();
  const segById = {};
  if (segLast >= 2) {
    const segVals = shSeg.getRange(2, 1, segLast - 1, shSeg.getLastColumn()).getValues();
    segVals.forEach(r => {
      const st = String(r[segMap.status] || '');
      if (st !== 'ACTIVE') return;
      const y = parseInt(r[segMap.year], 10);
      if (y !== year) return;
      const id = String(r[segMap.segId] || '');
      segById[id] = {
        segId: id,
        createdAt: r[segMap.createdAt],
        pct: toNumber_(r[segMap.pct]),
        label: String(r[segMap.label] || ''),
        createdBy: String(r[segMap.createdBy] || '')
      };
    });
  }

  const detMap = getHeaderMap_(shDet);
  const detLast = shDet.getLastRow();
  if (detLast < 2) return [];

  const detVals = shDet.getRange(2, 1, detLast - 1, shDet.getLastColumn()).getValues();
  const out = [];

  detVals.forEach(r => {
    const y = parseInt(r[detMap.year], 10);
    if (y !== year) return;
    const st = String(r[detMap.status] || '');
    if (st !== 'ACTIVE') return;

    const reg = String(r[detMap.regional] || '').trim();
    if (reg !== regional) return;

    const id = String(r[detMap.segId] || '');
    const meta = segById[id];
    if (!meta) return;

    out.push({
      segId: id,
      createdAt: meta.createdAt,
      pct: meta.pct,
      label: meta.label,
      createdBy: meta.createdBy,
      remainingBefore: toNumber_(r[detMap.remainingBefore]),
      amountSegmented: toNumber_(r[detMap.amountSegmented]),
      remainingAfter: toNumber_(r[detMap.remainingAfter])
    });
  });

  out.sort((a, b) => new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime());

  // Convert Dates to strings
  return out.map(item => ({
    ...item,
    createdAt: safeDate_(item.createdAt)
  }));
}

function listBaseRows_(year, regional, offset, limit) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_BASE);
  const last = sh.getLastRow();
  const cols = sh.getLastColumn();
  if (last < 2) return { columns: [], rows: [], offset, limit, total: 0 };

  const map = getHeaderMap_(sh);
  const idxYear = map['Año'];
  const idxRegional = map['Regional'];

  const all = sh.getRange(1, 1, last, cols).getValues();
  const headers = all[0].map(h => String(h || '').trim());

  const filtered = [];
  for (let i = 1; i < all.length; i++) {
    const r = all[i];
    if (parseInt(r[idxYear], 10) !== year) continue;
    if (String(r[idxRegional] || '').trim() !== regional) continue;
    filtered.push(r);
  }

  const total = filtered.length;
  const page = filtered.slice(offset, offset + limit);

  return {
    columns: headers,
    rows: page,
    offset,
    limit,
    total
  };
}

/***************
 * Distribución y utilidades
 ***************/
function distributeByRemaining_(remainingByRegional, targetTotal) {
  const regs = Object.keys(remainingByRegional).filter(r => (remainingByRegional[r] || 0) > 0);
  const sumRemaining = regs.reduce((acc, r) => acc + (remainingByRegional[r] || 0), 0);

  const dist = {};
  if (sumRemaining <= 0) return dist;

  // Primero proporcional, redondeo a 2 decimales
  let allocated = 0;
  regs.forEach(r => {
    const share = (remainingByRegional[r] / sumRemaining) * targetTotal;
    const amt = round2_(share);
    dist[r] = amt;
    allocated = round2_(allocated + amt);
  });

  // Ajuste por diferencias de redondeo (centavos)
  let diff = round2_(targetTotal - allocated);
  if (diff !== 0 && regs.length) {
    // Ajustar al regional con mayor disponible
    regs.sort((a, b) => (remainingByRegional[b] || 0) - (remainingByRegional[a] || 0));
    const top = regs[0];
    dist[top] = round2_((dist[top] || 0) + diff);

    // No permitir exceder su disponible
    if (dist[top] > remainingByRegional[top]) dist[top] = round2_(remainingByRegional[top]);
    if (dist[top] < 0) dist[top] = 0;
  }

  // Asegura no exceder disponible por regional
  Object.keys(dist).forEach(r => {
    const max = remainingByRegional[r] || 0;
    if (dist[r] > max) dist[r] = round2_(max);
    if (dist[r] < 0) dist[r] = 0;
  });

  return dist;
}

function toNumber_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const s = String(v || '').trim();
  if (!s) return 0;
  // soporta "1,234.56" o "1.234,56" (simple)
  const normalized = s.replace(/\s/g, '')
    .replace(/\./g, '')
    .replace(',', '.')
    .replace(/[^0-9.\-]/g, '');
  const n = parseFloat(normalized);
  return isFinite(n) ? n : 0;
}

function round2_(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

function ok_(data) {
  return { ok: true, data, message: '' };
}
function fail_(message) {
  return { ok: false, data: null, message: message || 'Error' };
}

function normalizeDriveUrl_(value) {
  if (!value) return '';
  const val = String(value).trim();
  if (!val) return '';
  // Si ya es URL, devolver tal cual
  if (/^https?:\/\//i.test(val)) return val;
  // Asumir que es un ID de archivo de Drive
  // Si la API Drive está habilitada, usarla para obtener webContentLink o thumbnailLink
  try {
      if (typeof Drive !== 'undefined') {
          // Intentar obtener el archivo
          const file = Drive.Files.get(val);
          if (file) {
              if (file.thumbnailLink) {
                 // Hack para mejorar la resolución del thumbnailLink
                 return file.thumbnailLink.replace('=s220', '=s1000');
              }
              if (file.webContentLink) {
                 return file.webContentLink;
              }
          }
      }
  } catch (e) {
      Logger.log('Error accediendo a Drive API para ' + val + ': ' + e);
  }
  return 'https://drive.google.com/uc?export=view&id=' + val;
}

function safeDate_(val) {
  if (val instanceof Date) return val.toISOString();
  // Si ya es string, devolver
  if (typeof val === 'string') return val;
  return '';
}