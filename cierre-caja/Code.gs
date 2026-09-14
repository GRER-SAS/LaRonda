/**
 * CIERRE DE CAJA - GRER S.A.S.
 * Google Apps Script + Google Sheets.
 */

const CONFIG = {
  SPREADSHEET_ID: '1Z63A-EUS4IecezgEmr5v_JG2Da0Hdo2xFf1oAmYNt8g',
  SHEET_NAME: 'Cierres de Caja',
  HOTELS: ['Azul de la Plaza', 'Chat Noir', 'La Ronda', 'Magdalena', 'Sanchez'],
  SHIFTS: ['Mañana', 'Tarde', 'Noche'],
  BILL_DENOMINATIONS: [100, 50, 20, 10, 5, 1],
  TIME_ZONE: 'America/Guayaquil',
  TOLERANCE_CENTS: 0
};

function doGet(e) {
  const page = e && e.parameter && e.parameter.page === 'dashboard' ? 'Dashboard' : 'Index';
  return HtmlService.createTemplateFromFile(page)
    .evaluate()
    .setTitle(page === 'Dashboard' ? 'Panel de Caja | GRER S.A.S.' : 'Cierre de Caja | GRER S.A.S.')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setup() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME);

  const headers = [
    'ID Cierre', 'Fecha', 'Hora registro', 'Recepcionista', 'Hotel', 'Turno',
    'Caja contada', 'Saldo anterior', 'Ingresos efectivo', 'Transferencias',
    'Tarjetas', 'Egresos', 'Efectivo esperado', 'Diferencia', 'Estado', 'Observaciones'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange('G:N').setNumberFormat('$0.00');
    sheet.autoResizeColumns(1, headers.length);
  }

  return `Hoja ${CONFIG.SHEET_NAME} lista.`;
}

function getInitialData() {
  setup();
  const today = new Date();
  const dates = [];
  for (let offset = -90; offset <= 30; offset++) {
    const d = new Date(today);
    d.setDate(today.getDate() + offset);
    dates.push(Utilities.formatDate(d, CONFIG.TIME_ZONE, 'yyyy-MM-dd'));
  }
  return {
    hotels: CONFIG.HOTELS,
    shifts: CONFIG.SHIFTS,
    denominations: CONFIG.BILL_DENOMINATIONS,
    dates,
    today: Utilities.formatDate(today, CONFIG.TIME_ZONE, 'yyyy-MM-dd'),
    sheetName: CONFIG.SHEET_NAME
  };
}

function getPreviousBalance(hotel, fecha) {
  setup();
  if (!CONFIG.HOTELS.includes(hotel)) return 0;
  if (!/^\\d{4}-\\d{2}-\\d{2}$/.test(String(fecha))) return 0;

  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return 0;

  let best = null;
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowFecha = String(row[1] || '');
    const rowHotel = String(row[4] || '');
    if (rowHotel !== hotel || rowFecha > fecha) continue;
    if (String(row[14] || '') !== 'CERRADO') continue;
    const candidate = {
      fecha: rowFecha,
      hora: String(row[2] || ''),
      cajaContada: Number(row[6] || 0)
    };
    if (!best || candidate.fecha > best.fecha || (candidate.fecha === best.fecha && candidate.hora > best.hora)) {
      best = candidate;
    }
  }

  return best ? roundMoney_(best.cajaContada) : 0;
}

function checkDuplicate(hotel, fecha, turno) {
  setup();
  const result = findDuplicate_(String(hotel || '').trim(), String(fecha || '').trim(), String(turno || '').trim());
  return { exists: !!result, message: result ? `Ya existe un cierre para ${hotel} · ${fecha} · ${turno}.` : '' };
}

function registrarCierre(payload) {
  const data = normalizePayload_(payload);
  validatePayload_(data);

  const expectedCents = data.saldoAnteriorCents + data.efectivoCents - data.egresosCents;
  const differenceCents = data.cajaContadaCents - expectedCents;

  if (Math.abs(differenceCents) > CONFIG.TOLERANCE_CENTS) {
    throw new Error(
      `No se puede cerrar la caja. La diferencia es de ${formatMoney_(differenceCents)}. ` +
      'La caja debe coincidir exactamente, incluso por $0,01.'
    );
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      setup();
      sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    }

    // Repetimos la validación dentro del bloqueo para evitar doble registro simultáneo.
    if (findDuplicate_(data.hotel, data.fecha, data.turno)) {
      throw new Error(`Ya existe un cierre para ${data.hotel} · ${data.fecha} · ${data.turno}.`);
    }

    const now = new Date();
    const id = Utilities.getUuid();
    const hour = Utilities.formatDate(now, CONFIG.TIME_ZONE, 'HH:mm:ss');

    sheet.appendRow([
      id, data.fecha, hour, data.recepcionista, data.hotel, data.turno,
      centsToNumber_(data.cajaContadaCents), centsToNumber_(data.saldoAnteriorCents),
      centsToNumber_(data.efectivoCents), centsToNumber_(data.transferenciaCents),
      centsToNumber_(data.tarjetaCents), centsToNumber_(data.egresosCents),
      centsToNumber_(expectedCents), centsToNumber_(differenceCents), 'CERRADO', data.observaciones
    ]);

    return {
      ok: true, id,
      expected: centsToNumber_(expectedCents),
      difference: centsToNumber_(differenceCents)
    };
  } finally {
    lock.releaseLock();
  }
}

function getDashboardData(filters) {
  setup();
  filters = filters || {};
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[0]) continue;
    const item = {
      id: String(r[0]), fecha: String(r[1]), hora: String(r[2]), recepcionista: String(r[3]),
      hotel: String(r[4]), turno: String(r[5]), cajaContada: Number(r[6] || 0),
      saldoAnterior: Number(r[7] || 0), efectivo: Number(r[8] || 0), transferencia: Number(r[9] || 0),
      tarjeta: Number(r[10] || 0), egresos: Number(r[11] || 0), esperado: Number(r[12] || 0),
      diferencia: Number(r[13] || 0), estado: String(r[14] || ''), observaciones: String(r[15] || '')
    };
    if (filters.hotel && filters.hotel !== 'Todos' && item.hotel !== filters.hotel) continue;
    if (filters.fecha && filters.fecha !== 'Todas' && item.fecha !== filters.fecha) continue;
    if (filters.turno && filters.turno !== 'Todos' && item.turno !== filters.turno) continue;
    rows.push(item);
  }

  rows.sort((a, b) => (b.fecha + ' ' + b.hora).localeCompare(a.fecha + ' ' + a.hora));

  const totalIngresos = rows.reduce((s, r) => s + r.efectivo + r.transferencia + r.tarjeta, 0);
  const totalEfectivo = rows.reduce((s, r) => s + r.efectivo, 0);
  const totalTransferencias = rows.reduce((s, r) => s + r.transferencia, 0);
  const totalTarjetas = rows.reduce((s, r) => s + r.tarjeta, 0);
  const totalEgresos = rows.reduce((s, r) => s + r.egresos, 0);
  const cierresCorrectos = rows.filter(r => Math.round(r.diferencia * 100) === 0).length;

  return {
    hotels: CONFIG.HOTELS,
    shifts: CONFIG.SHIFTS,
    rows,
    summary: {
      cierres: rows.length,
      cierresCorrectos,
      totalIngresos: roundMoney_(totalIngresos),
      totalEfectivo: roundMoney_(totalEfectivo),
      totalTransferencias: roundMoney_(totalTransferencias),
      totalTarjetas: roundMoney_(totalTarjetas),
      totalEgresos: roundMoney_(totalEgresos)
    }
  };
}

function findDuplicate_(hotel, fecha, turno) {
  if (!hotel || !fecha || !turno) return false;
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) return false;
  const values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 5).getValues();
  return values.some(r => String(r[0]) === fecha && String(r[3]) === hotel && String(r[4]) === turno);
}

function normalizePayload_(p) {
  p = p || {};
  const billCounts = p.billCounts || {};
  let billsCents = 0;
  CONFIG.BILL_DENOMINATIONS.forEach(d => {
    const qty = parseInteger_(billCounts[String(d)] || 0);
    billsCents += d * 100 * qty;
  });
  const coinsCents = parseMoneyCents_(p.coins || 0);
  return {
    fecha: String(p.fecha || '').trim(),
    recepcionista: String(p.recepcionista || '').trim(),
    hotel: String(p.hotel || '').trim(),
    turno: String(p.turno || '').trim(),
    cajaContadaCents: billsCents + coinsCents,
    saldoAnteriorCents: parseMoneyCents_(p.saldoAnterior),
    efectivoCents: parseMoneyCents_(p.efectivo),
    transferenciaCents: parseMoneyCents_(p.transferencia),
    tarjetaCents: parseMoneyCents_(p.tarjeta),
    egresosCents: parseMoneyCents_(p.egresos),
    observaciones: String(p.observaciones || '').trim().slice(0, 500),
    billCounts,
    coinsCents
  };
}

function validatePayload_(d) {
  if (!/^\\d{4}-\\d{2}-\\d{2}$/.test(d.fecha)) throw new Error('Fecha inválida.');
  if (!d.recepcionista) throw new Error('Ingrese el nombre del recepcionista.');
  if (!CONFIG.HOTELS.includes(d.hotel)) throw new Error('Seleccione un hotel válido.');
  if (!CONFIG.SHIFTS.includes(d.turno)) throw new Error('Seleccione un turno válido.');
  if (!Number.isInteger(d.cajaContadaCents) || d.cajaContadaCents < 0) throw new Error('El conteo físico de caja es inválido.');
  [d.saldoAnteriorCents, d.efectivoCents, d.transferenciaCents, d.tarjetaCents, d.egresosCents].forEach(v => {
    if (!Number.isInteger(v) || v < 0) throw new Error('Los valores monetarios deben ser números positivos o cero.');
  });
  CONFIG.BILL_DENOMINATIONS.forEach(denom => {
    const qty = parseInteger_(d.billCounts[String(denom)] || 0);
    if (qty < 0) throw new Error('La cantidad de billetes no puede ser negativa.');
  });
}

function parseInteger_(value) {
  if (value === null || value === undefined || value === '') return 0;
  const text = String(value).trim();
  if (!/^\\d+$/.test(text)) throw new Error(`Cantidad inválida: ${value}`);
  return Number(text);
}

function parseMoneyCents_(value) {
  if (value === null || value === undefined || value === '') return 0;
  const text = String(value).replace(',', '.').trim();
  if (!/^\\d+(?:\\.\\d{1,2})?$/.test(text)) throw new Error(`Valor monetario inválido: ${value}`);
  const [whole, fraction = ''] = text.split('.');
  return Number(whole) * 100 + Number((fraction + '00').slice(0, 2));
}

function centsToNumber_(cents) { return cents / 100; }
function roundMoney_(value) { return Math.round(Number(value) * 100) / 100; }
function formatMoney_(cents) {
  const sign = cents < 0 ? '-' : '';
  return `${sign}$${(Math.abs(cents) / 100).toFixed(2)}`;
}
