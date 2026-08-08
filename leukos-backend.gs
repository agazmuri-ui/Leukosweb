// ============================================================
//  LEUKÓS · Backend (Google Apps Script)
//  Pegar en: Google Sheet → Extensiones → Apps Script
//  Luego: Implementar → Nueva implementación → Aplicación web
//    · Ejecutar como: Yo
//    · Quién accede: Cualquier usuario
// ============================================================

const SPREADSHEET_ID = '1lFqHulZQ5utE7vjPBB6dQ-C6UcVTp0eIsCbpe3TvblI';
const MP_ACCESS_TOKEN = 'AGREGA_TU_ACCESS_TOKEN_AQUI'; // ← pega tu token productivo aquí
const LEUKOS_URL = 'https://leukos.cl';

// Correos que recibirán cada pedido
const ORDER_EMAILS = ['agazmuri@colegiosanbenito.org', 'velasleukos@gmail.com'];

// Precios por tamaño
const PRICES = {
  'Pequeña':  9000,
  'Pequeña+': 12000,
  'Mediana':  16000,
  'Grande':   20000
};

// Etiquetas de envío
const SHIPPING_LABELS = {
  'retiro-vitacura':  'Retiro en Vitacura (gratis)',
  'retiro-chicureo':  'Retiro en Chicureo (gratis)',
  'despacho-local':   'Despacho a Vitacura / Las Condes / Chicureo (+$3.500)',
  'starken':          'Starken (por pagar en destino)',
  'blue-express':     'Blue Express (por pagar en destino)'
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

// ── POST: recibe pedido, envía email y redirige a MP ─────────
function doPost(e) {
  try {
    const cartJson    = e.parameter.cart;
    const customerJson = e.parameter.customer;
    if (!cartJson) throw new Error('Sin datos de carrito');

    const items    = JSON.parse(cartJson);
    const customer = customerJson ? JSON.parse(customerJson) : {};

    // Enviar email de pedido
    try { sendOrderEmail(items, customer); } catch(mailErr) {
      Logger.log('Error enviando email: ' + mailErr.message);
    }

    // Crear preferencia MP (incluye despacho si corresponde)
    const initPoint = createMPPreference(items, customer);

    return HtmlService.createHtmlOutput(
      `<html><head><script>window.top.location.replace("${initPoint}");</script></head>` +
      `<body style="font-family:sans-serif;padding:40px;text-align:center">` +
      `Redirigiendo a Mercado Pago...</body></html>`
    );
  } catch (err) {
    return HtmlService.createHtmlOutput(
      `<html><body style="font-family:sans-serif;padding:40px;text-align:center">` +
      `<p>Error al procesar el pago: ${err.message}</p>` +
      `<a href="${LEUKOS_URL}">← Volver a Leukós</a></body></html>`
    );
  }
}

// ── Envía email con los detalles del pedido ──────────────────
function sendOrderEmail(items, customer) {
  const subtotal = items.reduce((sum, i) => sum + (i.price || 0), 0);
  const shippingPrice = customer.shippingPrice || 0;
  const total = subtotal + shippingPrice;

  const shippingLabel = SHIPPING_LABELS[customer.envio] || customer.envio || '—';

  // Líneas del pedido
  const itemsHtml = items.map(i =>
    `<tr>
      <td style="padding:8px 12px;">${i.size} · ${i.frasco} · ${i.aroma}</td>
      <td style="padding:8px 12px; text-align:right;">$${(i.price||0).toLocaleString()}</td>
    </tr>`
  ).join('');

  const html = `
    <div style="font-family:sans-serif; max-width:560px; margin:0 auto; color:#2c2c2c;">
      <h2 style="background:#2c2c2c; color:#f5f0e8; padding:20px 24px; margin:0; letter-spacing:2px;">
        NUEVO PEDIDO · LEUKÓS
      </h2>
      <div style="padding:24px;">
        <h3 style="margin-top:0;">Datos del cliente</h3>
        <p><strong>Nombre:</strong> ${customer.nombre || '—'}</p>
        <p><strong>Teléfono:</strong> ${customer.telefono || '—'}</p>
        <p><strong>Envío:</strong> ${shippingLabel}</p>
        ${customer.direccion ? `<p><strong>Dirección:</strong> ${customer.direccion}</p>` : ''}

        <h3 style="margin-top:28px;">Productos</h3>
        <table style="width:100%; border-collapse:collapse; border:1px solid #e5ddd0;">
          <thead>
            <tr style="background:#f5f0e8;">
              <th style="padding:8px 12px; text-align:left;">Producto</th>
              <th style="padding:8px 12px; text-align:right;">Precio</th>
            </tr>
          </thead>
          <tbody>${itemsHtml}</tbody>
          <tfoot>
            ${shippingPrice > 0 ? `
            <tr>
              <td style="padding:8px 12px;">Despacho</td>
              <td style="padding:8px 12px; text-align:right;">$${shippingPrice.toLocaleString()}</td>
            </tr>` : ''}
            <tr style="background:#f5f0e8; font-weight:bold;">
              <td style="padding:10px 12px;">TOTAL</td>
              <td style="padding:10px 12px; text-align:right;">$${total.toLocaleString()}</td>
            </tr>
          </tfoot>
        </table>

        <p style="margin-top:24px; font-size:12px; color:#888;">
          Pedido recibido · ${new Date().toLocaleString('es-CL', {timeZone:'America/Santiago'})}
        </p>
      </div>
    </div>
  `;

  const subject = `🕯️ Nuevo pedido Leukós · ${customer.nombre || 'Cliente'} · $${total.toLocaleString()}`;

  ORDER_EMAILS.forEach(email => {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: html
    });
  });
}

// ── Lee hoja "Terminados" y devuelve productos con stock ─────
function getInventory() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Terminados');
  if (!sheet) throw new Error('Hoja "Terminados" no encontrada');

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  // Headers: Key | Linea | Variante | Cantidad
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

    const frascoFull = parts[1].trim();
    const aroma      = parts[2].trim();

    const fu = frascoFull.toUpperCase();
    let size, price;
    if      (fu.startsWith('P+ ')) { size = 'pequeña+'; price = PRICES['Pequeña+']; }
    else if (fu.startsWith('P '))  { size = 'pequeña';  price = PRICES['Pequeña'];  }
    else if (fu.startsWith('M '))  { size = 'mediana';  price = PRICES['Mediana'];  }
    else if (fu.startsWith('G '))  { size = 'grande';   price = PRICES['Grande'];   }
    else continue;

    const frascoColor = frascoFull
      .replace(/^[A-Z\+]+\s+/, '')
      .replace(/\s*\([\d\-g]+\)\s*$/, '')
      .trim();

    products.push({ key, size, frasco: frascoColor, aroma, price, stock });
  }

  return products;
}

// ── Crea preferencia en Mercado Pago ────────────────────────
function createMPPreference(items, customer) {
  const mpItems = items.map(item => ({
    title: `Vela Leukós ${item.size} · ${item.frasco} · ${item.aroma}`,
    quantity: 1,
    unit_price: item.price,
    currency_id: 'CLP'
  }));

  // Agregar despacho si tiene costo
  const shippingPrice = customer.shippingPrice || 0;
  if (shippingPrice > 0) {
    mpItems.push({
      title: 'Despacho',
      quantity: 1,
      unit_price: shippingPrice,
      currency_id: 'CLP'
    });
  }

  const payload = {
    items: mpItems,
    back_urls: {
      success: LEUKOS_URL + '/?pago=exitoso',
      failure: LEUKOS_URL + '/?pago=fallido',
      pending: LEUKOS_URL + '/?pago=pendiente'
    },
    auto_return: 'approved',
    statement_descriptor: 'LEUKOS VELAS',
    payer: customer.nombre ? {
      name: customer.nombre,
      phone: { area_code: '56', number: customer.telefono || '' }
    } : undefined
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

// ── Helpers ─────────────────────────────────────────────────
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
