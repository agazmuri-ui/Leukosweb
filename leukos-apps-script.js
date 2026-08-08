// ============================================================
//  LEUKÓS · Backend (Google Apps Script)
//  Pegar en: Google Sheet → Extensiones → Apps Script
//  Luego: Implementar → Nueva implementación → Aplicación web
//    · Ejecutar como: Yo
//    · Quién accede: Cualquier usuario
// ============================================================

const SPREADSHEET_ID = '1lFqHulZQ5utE7vjPBB6dQ-C6UcVTp0eIsCbpe3TvblI';
const MP_ACCESS_TOKEN = 'APP_USR-1555794936589450-071909-9e6baf0da808b4d9b71b11c440f64796-3554180428';
const LEUKOS_URL = 'https://leukos.cl';

// Precios por tamaño
const PRICES = {
  'pequeña':  9000,
  'pequeña+': 12000,
  'mediana':  16000,
  'grande':   20000
};

// ── GET: devuelve inventario disponible ──────────────────────
function doGet(e) {
  try {
    const products = getInventory();
    return jsonResponse({ ok: true, products });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ── POST: crea preferencia MP y devuelve JSON ────────────────
function doPost(e) {
  try {
    const cartJson = e.parameter.cart;
    if (!cartJson) throw new Error('Sin datos de carrito');
    const items = JSON.parse(cartJson);
    const initPoint = createMPPreference(items);
    return jsonResponse({ ok: true, init_point: initPoint });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ── Lee hoja "Terminados" y devuelve productos con stock ─────
function getInventory() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Terminados');
  if (!sheet) throw new Error('Hoja "Terminados" no encontrada');

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  // Headers: Key | Linea | Variante | Cantidad | Actualizado
  const products = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const key      = String(row[0] || '').trim();
    const variante = String(row[2] || '').trim();
    const stock    = parseInt(row[3]) || 0;

    if (!key || !variante) continue;

    // Variante: "Coco · P Negro (180g-340g) · Lavanda"
    const parts = variante.split(' · ');
    if (parts.length < 3) continue;

    const frascoFull = parts[1].trim();  // "P Negro (180g-340g)"
    const aroma      = parts[2].trim();  // "Lavanda"

    // Detectar tamaño por prefijo del frasco
    const fu = frascoFull.toUpperCase();
    let size, price;
    if      (fu.startsWith('P+ ')) { size = 'pequeña+'; price = PRICES['pequeña+']; }
    else if (fu.startsWith('P '))  { size = 'pequeña';  price = PRICES['pequeña'];  }
    else if (fu.startsWith('M '))  { size = 'mediana';  price = PRICES['mediana'];  }
    else if (fu.startsWith('G '))  { size = 'grande';   price = PRICES['grande'];   }
    else continue;

    // Color del frasco: quitar prefijo de tamaño y paréntesis de peso
    const frascoColor = frascoFull
      .replace(/^[A-Z\+]+\s+/, '')           // "P+ ", "M ", etc.
      .replace(/\s*\([\d\-g]+\)\s*$/, '')    // "(180g-340g)"
      .trim();

    products.push({ key, size, frasco: frascoColor, aroma, price, stock });
  }

  return products;
}

// ── Crea preferencia en Mercado Pago ────────────────────────
function createMPPreference(items) {
  const mpItems = items.map(item => ({
    title: `Vela Leukós ${capitalize(item.size)} · ${item.frasco} · ${item.aroma}`,
    quantity: 1,
    unit_price: item.price,
    currency_id: 'CLP'
  }));

  const payload = {
    items: mpItems,
    back_urls: {
      success: LEUKOS_URL + '/?pago=exitoso',
      failure: LEUKOS_URL + '/?pago=fallido',
      pending: LEUKOS_URL + '/?pago=pendiente'
    },
    auto_return: 'approved',
    statement_descriptor: 'LEUKOS VELAS'
  };

  const response = UrlFetchApp.fetch('https://api.mercadopago.com/checkout/preferences', {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + MP_ACCESS_TOKEN,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());
  if (!result.init_point) throw new Error('MP no devolvió init_point: ' + response.getContentText());
  return result.init_point;
}

// ── Helpers ──────────────────────────────────────────────────
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function capitalize(str) {
  return str.charAt(0).toUpperCase() + str.slice(1);
}

function testPermiso() {
  const r = UrlFetchApp.fetch('https://www.google.com');
  Logger.log(r.getResponseCode());
}
