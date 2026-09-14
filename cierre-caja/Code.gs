/**
 * CIERRE DE CAJA - GRER S.A.S.
 * Backend para Google Apps Script + Google Sheets.
 *
 * La caja solo puede cerrarse cuando:
 * caja contada = saldo anterior + ingresos efectivo - egresos
 * La comparación se realiza en centavos enteros para evitar errores de punto flotante.
 */

const CONFIG = {
  SPREADSHEET_ID: '1Z63A-EUS4IecezgEmr5v_JG2Da0Hdo2xFf1oAmYNt8g',
  SHEET_NAME: 'Cierres de Caja',
  HOTELS: ['Azul de la Plaza', 'Chat Noir', 'La Ronda', 'Magdalena', 'Sanchez'],
  TIME_ZONE: 'America/Guayaquil',
  TOLERANCE_CENTS: 0
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Cierre de Caja | GRER S.A.S.')
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

  // Permite cierres históricos y fechas futuras razonables.
  for (let offset = -90; offset <= 30; offset++) {
    const d = new Date(today);
    d.setDate(today.getDate() + offset);
    dates.push(Utilities.formatDate(d, CONFIG.TIME_ZONE, 'yyyy-MM-dd'));
  }

  return {
    hotels: CONFIG.HOTELS,
    dates,
    today: Utilities.formatDate(today, CONFIG.TIME_ZONE, 'yyyy-MM-dd'),
    sheetName: CONFIG.SHEET_NAME
  };
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

    const now = new Date();
    const id = Utilities.getUuid();
    const hour = Utilities.formatDate(now, CONFIG.TIME_ZONE, 'HH:mm:ss');

    sheet.appendRow([
      id,
      data.fecha,
      hour,
      data.recepcionista,
      data.hotel,
      data.turno,
      centsToNumber_(data.cajaContadaCents),
      centsToNumber_(data.saldoAnteriorCents),
      centsToNumber_(data.efectivoCents),
      centsToNumber_(data.transferenciaCents),
      centsToNumber_(data.tarjetaCents),
      centsToNumber_(data.egresosCents),
      centsToNumber_(expectedCents),
      centsToNumber_(differenceCents),
      'CERRADO',
      data.observaciones
    ]);

    return {
      ok: true,
      id,
      expected: centsToNumber_(expectedCents),
      difference: centsToNumber_(differenceCents)
    };
  } finally {
    lock.releaseLock();
  }
}

function normalizePayload_(p) {
  p = p || {};
  return {
    fecha: String(p.fecha || '').trim(),
    recepcionista: String(p.recepcionista || '').trim(),
    hotel: String(p.hotel || '').trim(),
    turno: String(p.turno || '').trim(),
    cajaContadaCents: parseMoneyCents_(p.cajaContada),
    saldoAnteriorCents: parseMoneyCents_(p.saldoAnterior),
    efectivoCents: parseMoneyCents_(p.efectivo),
    transferenciaCents: parseMoneyCents_(p.transferencia),
    tarjetaCents: parseMoneyCents_(p.tarjeta),
    egresosCents: parseMoneyCents_(p.egresos),
    observaciones: String(p.observaciones || '').trim().slice(0, 500)
  };
}

function validatePayload_(d) {
  if (!/^\\d{4}-\\d{2}-\\d{2}$/.test(d.fecha)) throw new Error('Fecha inválida.');
  if (!d.recepcionista) throw new Error('Ingrese el nombre del recepcionista.');
  if (!CONFIG.HOTELS.includes(d.hotel)) throw new Error('Seleccione un hotel válido.');
  if (!['Mañana', 'Tarde', 'Noche'].includes(d.turno)) throw new Error('Seleccione un turno válido.');

  [d.cajaContadaCents, d.saldoAnteriorCents, d.efectivoCents,
   d.transferenciaCents, d.tarjetaCents, d.egresosCents].forEach((v) => {
    if (!Number.isInteger(v) || v < 0) throw new Error('Los valores monetarios deben ser números positivos o cero.');
  });
}

function parseMoneyCents_(value) {
  if (value === null || value === undefined || value === '') return 0;
  const text = String(value).replace(',', '.').trim();
  if (!/^\\d+(?:\\.\\d{1,2})?$/.test(text)) throw new Error(`Valor monetario inválido: ${value}`);
  const [whole, fraction = ''] = text.split('.');
  return Number(whole) * 100 + Number((fraction + '00').slice(0, 2));
}

function centsToNumber_(cents) {
  return cents / 100;
}

function formatMoney_(cents) {
  const sign = cents < 0 ? '-' : '';
  const abs = Math.abs(cents);
  return `${sign}$${(abs / 100).toFixed(2)}`;
}
