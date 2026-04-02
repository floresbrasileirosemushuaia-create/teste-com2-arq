const SS_ID = '1kfQ5lIQo9cLtu0SlXBa6eodUruP2HmYPrli4_jBjqXw';

function doGet(e) {
  if (e && e.parameter && e.parameter.api === '1') {
    return handleApiGet_(e);
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Catálogo de excursiones')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  if (e && e.parameter && e.parameter.api === '1') {
    return handleApiPost_(e);
  }

  return HtmlService.createHtmlOutput('OK');
}

function handleApiGet_(e) {
  var action = String(e.parameter.action || '').trim();
  var callback = sanitizeCallback_(e.parameter.callback);

  try {
    if (action === 'catalog') {
      return jsonpResponse_(callback, {
        ok: true,
        data: getCatalogoPublicoData()
      });
    }

    return jsonpResponse_(callback, {
      ok: false,
      error: 'Acción GET inválida.'
    });
  } catch (err) {
    return jsonpResponse_(callback, {
      ok: false,
      error: err && err.message ? err.message : String(err)
    });
  }
}

function handleApiPost_(e) {
  var action = String(e.parameter.action || '').trim();

  try {
    if (action === 'submit') {
      var payload = JSON.parse(String(e.parameter.payload || '{}'));
      var result = submitPedidoAndBuildWhatsAppUrl(payload);

      return postMessageResponse_({
        source: 'perla-api-submit',
        ok: true,
        data: result
      });
    }

    return postMessageResponse_({
      source: 'perla-api-submit',
      ok: false,
      error: 'Acción POST inválida.'
    });
  } catch (err) {
    return postMessageResponse_({
      source: 'perla-api-submit',
      ok: false,
      error: err && err.message ? err.message : String(err)
    });
  }
}

function sanitizeCallback_(callback) {
  var cb = String(callback || 'callback').replace(/[^\w$.]/g, '');
  return cb || 'callback';
}

function jsonpResponse_(callback, payload) {
  var body = callback + '(' + JSON.stringify(payload) + ');';
  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function postMessageResponse_(payload) {
  var safeJson = JSON.stringify(payload).replace(/</g, '\\u003c');
  var html = ''
    + '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body>'
    + '<script>'
    + '(function(){'
    + 'var data=' + safeJson + ';'
    + 'if (window.parent) { window.parent.postMessage(data, "*"); }'
    + 'else if (window.top) { window.top.postMessage(data, "*"); }'
    + '})();'
    + '<\/script>'
    + '</body></html>';

  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function normalizeText_(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^\w\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeLogoUrl(raw) {
  if (!raw) return '';
  var s = String(raw).trim();
  if (!s) return '';

  var match = s.match(/\/file\/d\/([^/]+)\//);
  if (match && match[1]) {
    return 'https://drive.google.com/uc?export=view&id=' + match[1];
  }

  if (!/^https?:\/\//i.test(s) && /^[A-Za-z0-9_-]{15,}$/.test(s)) {
    return 'https://drive.google.com/uc?export=view&id=' + s;
  }

  return s;
}

function getSheetByPossibleNames_(ss, names) {
  for (var i = 0; i < names.length; i++) {
    var sh = ss.getSheetByName(names[i]);
    if (sh) return sh;
  }
  return null;
}

function findHeaderIndex_(headers, aliases) {
  var normalizedHeaders = headers.map(normalizeText_);
  for (var i = 0; i < aliases.length; i++) {
    var idx = normalizedHeaders.indexOf(normalizeText_(aliases[i]));
    if (idx > -1) return idx;
  }
  return -1;
}

function formatCurrencyAr_(num) {
  var n = Number(num || 0);
  var rounded = Math.round(n);
  return 'AR$ ' + rounded.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

function parseMoneyToNumber_(value) {
  var s = String(value || '').replace(/[^\d.-]/g, '');
  return Number(s || 0);
}

function getSetupMap_(ss) {
  var setup = getSheetByPossibleNames_(ss, ['setup', 'Setup', 'SETUP']);
  if (!setup) return {};

  var lastRow = Math.max(setup.getLastRow(), 1);
  var values = setup.getRange(1, 1, lastRow, 2).getDisplayValues();
  var map = {};

  values.forEach(function(row) {
    var key = normalizeText_(row[0]);
    var value = String(row[1] || '').trim();
    if (key) map[key] = value;
  });

  return map;
}

function buildItemLabel_(item) {
  var tipo = String(item.tipo || '').trim();
  var categoria = String(item.categoria || '').trim();
  var actividad = String(item.actividad || '').trim();
  var title = String(item.title || '').trim();

  if (tipo && categoria) return tipo + ' - ' + categoria;
  if (tipo && actividad) return tipo + ' - ' + actividad;
  if (title) return title;
  if (actividad) return actividad;
  return tipo || 'Servicio';
}

function getCatalogoPublicoData() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = getSheetByPossibleNames_(ss, ['catalogo', 'Catalogo', 'CATALOGO']);

  if (!sheet) {
    throw new Error('No encontré la hoja "catalogo".');
  }

  var values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return {
      logoUrl: '',
      categories: [],
      items: [],
      whatsappNumber: ''
    };
  }

  var headers = values[0];
  var rows = values.slice(1);

  var idxTipo = findHeaderIndex_(headers, ['TIPO']);
  var idxActividad = findHeaderIndex_(headers, ['Actividad']);
  var idxCategoria = findHeaderIndex_(headers, ['Categoria', 'Categoría']);
  var idxPrecio = findHeaderIndex_(headers, ['Precio']);
  var idxObs = findHeaderIndex_(headers, ['obs', 'OBS']);

  if (idxTipo === -1 || idxActividad === -1 || idxCategoria === -1 || idxPrecio === -1 || idxObs === -1) {
    throw new Error('Faltan columnas obligatorias en catalogo: TIPO, Actividad, Categoria, Precio, obs.');
  }

  var setupMap = getSetupMap_(ss);
  var logoUrl = normalizeLogoUrl(setupMap['logo'] || '');
  var whatsappNumber = String(setupMap['whatsapp'] || '').trim();

  var categoryMap = {};
  var items = [];

  rows.forEach(function(row, i) {
    var tipo = String(row[idxTipo] || '').trim();
    var actividad = String(row[idxActividad] || '').trim();
    var categoria = String(row[idxCategoria] || '').trim();
    var precioText = String(row[idxPrecio] || '').trim();
    var obs = String(row[idxObs] || '').trim();

    if (!tipo && !actividad && !categoria && !precioText && !obs) return;

    var title = actividad || tipo || 'Servicio';
    if (tipo && actividad && normalizeText_(actividad).indexOf(normalizeText_(tipo)) === -1) {
      title = tipo + ' - ' + actividad;
    }

    items.push({
      rowIndex: i + 2,
      tipo: tipo || 'Sin categoría',
      actividad: actividad,
      categoria: categoria,
      title: title,
      priceText: precioText,
      priceNumber: parseMoneyToNumber_(precioText),
      obs: obs
    });

    if (tipo) categoryMap[tipo] = true;
  });

  return {
    logoUrl: logoUrl,
    whatsappNumber: whatsappNumber,
    categories: Object.keys(categoryMap).sort(function(a, b) {
      return a.localeCompare(b, 'es', { sensitivity: 'base' });
    }),
    items: items
  };
}

function getOrCreatePedidosSheet_(ss) {
  var sheet = ss.getSheetByName('Pedidos');
  if (!sheet) {
    sheet = ss.insertSheet('Pedidos');
    sheet.appendRow([
      'FechaHora',
      'Nombre',
      'WhatsApp',
      'Hospedaje',
      'PlataformaOrigen',
      'ReferidoNombre',
      'OtroOrigen',
      'Items',
      'Total',
      'MensajeWhatsApp',
      'Estado'
    ]);
  }
  return sheet;
}

function sanitizePhoneForWa_(value) {
  return String(value || '').replace(/\D/g, '');
}

function buildItemsSummary_(items) {
  return items.map(function(item) {
    var lineTotal = Number(item.priceNumber || 0) * Number(item.qty || 0);
    return [
      item.qty + 'x',
      buildItemLabel_(item),
      '| Fecha: ' + (item.desiredDate || 'Sin fecha definida'),
      '| ' + formatCurrencyAr_(lineTotal)
    ].join(' ').replace(/\s+/g, ' ').trim();
  }).join(' || ');
}

function buildWhatsAppMessage_(payload, totalFormatted) {
  var lines = [];

  lines.push('Perla Andina');
  lines.push('Hola, soy ' + payload.customer.name + ' quiero solicitar estas excursiones:');
  lines.push('');

  payload.items.forEach(function(item) {
    var lineTotal = Number(item.priceNumber || 0) * Number(item.qty || 0);
    lines.push('• ' + item.qty + 'x ' + buildItemLabel_(item));
    lines.push(' | Fecha: ' + (item.desiredDate || 'Sin fecha definida') + ' | ' + formatCurrencyAr_(lineTotal));
    lines.push('');
  });

  lines.push('Nombre: ' + payload.customer.name);
  lines.push('WhatsApp: ' + (payload.customer.whatsapp || '-'));
  lines.push('Hospedaje: ' + (payload.customer.hotel || '-'));

  return lines.join('\n').replace(/\n{3,}/g, '\n\n');
}

function submitPedidoAndBuildWhatsAppUrl(payload) {
  if (!payload || !payload.customer || !payload.source || !payload.items) {
    throw new Error('Datos incompletos del pedido.');
  }

  var ss = SpreadsheetApp.openById(SS_ID);
  var setupMap = getSetupMap_(ss);
  var waNumber = sanitizePhoneForWa_(setupMap['whatsapp'] || '');

  if (!waNumber) {
    throw new Error('No encontré el número de WhatsApp en setup.');
  }

  var customerName = String(payload.customer.name || '').trim();
  if (!customerName) {
    throw new Error('El nombre completo es obligatorio.');
  }

  if (!Array.isArray(payload.items) || payload.items.length === 0) {
    throw new Error('El carrito está vacío.');
  }

  var total = payload.items.reduce(function(sum, item) {
    return sum + (Number(item.priceNumber || 0) * Number(item.qty || 0));
  }, 0);

  var totalFormatted = formatCurrencyAr_(total);
  var message = buildWhatsAppMessage_(payload, totalFormatted);
  var itemsSummary = buildItemsSummary_(payload.items);

  var pedidosSheet = getOrCreatePedidosSheet_(ss);
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');

  pedidosSheet.appendRow([
    timestamp,
    customerName,
    String(payload.customer.whatsapp || '').trim(),
    String(payload.customer.hotel || '').trim(),
    String(payload.source.platform || '').trim(),
    String(payload.source.referredName || '').trim(),
    String(payload.source.otherOrigin || '').trim(),
    itemsSummary,
    totalFormatted,
    message,
    'NUEVO'
  ]);

  return {
    whatsappUrl: 'https://wa.me/' + waNumber + '?text=' + encodeURIComponent(message),
    whatsappNumber: waNumber,
    message: message,
    total: totalFormatted
  };
}
