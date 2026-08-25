/**
 * Google Apps Script API untuk ticketing event.
 *
 * Sebelum deploy:
 * 1. Buat Google Sheet kosong.
 * 2. Extensions > Apps Script, tempel file ini.
 * 3. Sesuaikan harga, kuota, dan instruksi pembayaran di CONFIG.
 * 4. Jalankan setupSheet() sekali dan izinkan permission.
 */

const CONFIG = {
  SHEET_NAME: 'Orders',
  TIMEZONE: Session.getScriptTimeZone() || 'Asia/Jakarta',
  RESERVE_PENDING_STOCK: true,
  PAYMENT_PROOF_FOLDER: 'Payment Proofs',
  PAYMENT_INSTRUCTION: 'QRIS tersedia di halaman pembayaran. Untuk transfer bank, gunakan BCA 3850962005 a.n. MUHAMMAD ZAIDAN. Kirim bukti pembayaran dengan Order ID Anda.',
  TICKETS: {
    'Presale': { price: 55000, quota: 500 }
  }
};

const HEADERS = [
  'Timestamp',
  'Order ID',
  'Nama Lengkap',
  'Email',
  'Nomor WhatsApp',
  'Kategori Tiket',
  'Jumlah Tiket',
  'Total Harga',
  'Status Pembayaran',
  'E-Ticket Code / QR Code Unique String',
  'Metode Pembayaran',
  'Bukti Pembayaran',
  'E-Ticket Email Sent At'
];

/** Run once after binding the script to the spreadsheet. */
function setupSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  } else {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }

  return jsonResponse({ status: 'success', message: 'Sheet siap digunakan.' });
}

/** Rewrites only row 1 so every status/email column has the correct position. */
function perbaikiHeaderSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Tab Orders belum ditemukan.');
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  return 'Header Orders sudah diperbaiki. Status Pembayaran berada di kolom I dan Email Sent At di kolom M.';
}

/**
 * GET /exec?action=stock
 * Returns the remaining quota for every configured ticket category.
 */
function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || 'stock').toLowerCase();

    if (action !== 'stock') {
      return jsonResponse({ status: 'failed', message: 'Action GET tidak dikenal.' });
    }

    return jsonResponse({ status: 'success', data: getStock() });
  } catch (error) {
    return jsonResponse({ status: 'failed', message: error.message });
  }
}

/**
 * POST /exec
 * Expected JSON: { nama, email, wa, kategori, jumlah }
 * Content-Type should be text/plain to avoid a browser CORS preflight.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  let lockAcquired = false;

  try {
    lock.waitLock(15000);
    lockAcquired = true;

    const payload = parsePayload(e);
    const name = cleanText(payload.nama || payload.name);
    const email = cleanText(payload.email);
    const whatsapp = cleanText(payload.wa || payload.whatsapp);
    const category = cleanText(payload.kategori || payload.category);
    const quantity = parseQuantity(payload.jumlah || payload.quantity);
    const paymentMethod = cleanText(payload.metodePembayaran || payload.paymentMethod);
    const paymentProof = payload.buktiPembayaran || payload.paymentProof;
    const ticket = CONFIG.TICKETS[category];

    validateOrder(name, email, whatsapp, category, quantity, ticket, paymentMethod, paymentProof);

    const stock = getStock()[category];
    if (stock.remaining < quantity) {
      throw new Error('Kuota ' + category + ' tidak mencukupi. Sisa: ' + stock.remaining + '.');
    }

    const total = ticket.price * quantity;
    const orderId = createOrderId();
    const qrCode = 'ET-' + Utilities.getUuid().replace(/-/g, '').toUpperCase();
    const proofUrl = savePaymentProof(paymentProof, orderId);
    const sheet = getOrdersSheet();

    sheet.appendRow([
      new Date(),
      orderId,
      name,
      email,
      whatsapp,
      category,
      quantity,
      total,
      'PENDING',
      qrCode,
      paymentMethod,
      proofUrl,
      ''
    ]);

    return jsonResponse({
      status: 'success',
      orderId: orderId,
      totalBayar: total,
      instruksiPembayaran: CONFIG.PAYMENT_INSTRUCTION,
      eTicketCode: qrCode
    });
  } catch (error) {
    return jsonResponse({ status: 'failed', message: error.message });
  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}

function getStock() {
  const used = {};
  Object.keys(CONFIG.TICKETS).forEach(function(category) {
    used[category] = 0;
  });

  const sheet = getOrdersSheet();
  const rowCount = sheet.getLastRow() - 1;
  if (rowCount > 0) {
    const rows = sheet.getRange(2, 6, rowCount, 4).getValues();
    rows.forEach(function(row) {
      const category = String(row[0]);
      const quantity = Number(row[1]) || 0;
      const paymentStatus = String(row[3]).toUpperCase();
      const shouldReserve = CONFIG.RESERVE_PENDING_STOCK
        ? paymentStatus !== 'CANCELLED'
        : paymentStatus === 'SUCCESS';

      if (Object.prototype.hasOwnProperty.call(used, category) && shouldReserve) {
        used[category] += quantity;
      }
    });
  }

  const result = {};
  Object.keys(CONFIG.TICKETS).forEach(function(category) {
    const quota = CONFIG.TICKETS[category].quota;
    result[category] = {
      price: CONFIG.TICKETS[category].price,
      quota: quota,
      terjual: used[category],
      remaining: Math.max(0, quota - used[category])
    };
  });
  return result;
}

/**
 * Installable trigger: sends the e-ticket when Status Pembayaran becomes
 * LUNAS or SUCCESS. The email is sent only once per order.
 */
function onEdit(e) {
  if (!e || !e.range || !e.value) return;

  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== CONFIG.SHEET_NAME || range.getRow() < 2) return;

  const statusColumn = findHeaderColumn(sheet, 'Status Pembayaran');
  if (range.getColumn() !== statusColumn) return;

  const status = normalizeStatus(e.value);
  if (!isPaidStatus(status)) return;

  const sentColumn = findHeaderColumn(sheet, 'E-Ticket Email Sent At');
  if (isEmailSentMarker(sheet.getRange(range.getRow(), sentColumn).getValue())) return;

  sendETicketEmail(sheet, range.getRow(), sentColumn);
}

/** Run once to create the authorized spreadsheet edit trigger. */
function pasangTriggerEmail() {
  const spreadsheet = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();

  return 'Trigger email berhasil dipasang.';
}

/** Run with the order row selected to resend its e-ticket once. */
function kirimETicketBarisAktif() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const activeRange = sheet ? sheet.getActiveRange() : null;
  const rowNumber = activeRange ? activeRange.getRow() : 0;
  if (!sheet || sheet.getName() !== CONFIG.SHEET_NAME || rowNumber < 2) {
    throw new Error('Buka tab Orders, klik salah satu sel pada baris order, lalu jalankan lagi.');
  }

  const sentColumn = HEADERS.indexOf('E-Ticket Email Sent At') + 1;
  sendETicketEmail(sheet, rowNumber, sentColumn);
  return 'E-ticket berhasil dikirim ulang.';
}

/** Sends every unpaid-email e-ticket that is already marked LUNAS or SUCCESS. */
function kirimSemuaETicketLunas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Tab Orders belum ditemukan.');

  const statusColumn = findHeaderColumn(sheet, 'Status Pembayaran');
  const sentColumn = findHeaderColumn(sheet, 'E-Ticket Email Sent At');
  const lastRow = sheet.getLastRow();
  let sentCount = 0;
  let eligibleCount = 0;
  let paidCount = 0;

  for (let rowNumber = 2; rowNumber <= lastRow; rowNumber++) {
    const rawStatus = sheet.getRange(rowNumber, statusColumn).getValue();
    const status = normalizeStatus(rawStatus);
    const sentAt = sheet.getRange(rowNumber, sentColumn).getValue();
    const emailAlreadySent = isEmailSentMarker(sentAt);
    if (isPaidStatus(status)) paidCount++;
    if (isPaidStatus(status) && !emailAlreadySent) {
      eligibleCount++;
      try {
        sendETicketEmail(sheet, rowNumber, sentColumn);
        sentCount++;
      } catch (error) {
        Logger.log('Gagal baris ' + rowNumber + ': ' + error.message);
      }
    } else if (rawStatus && (!isPaidStatus(status) || emailAlreadySent)) {
      Logger.log('Baris ' + rowNumber + ' dilewati: status terbaca sebagai [' + rawStatus + '], normalisasi [' + status + '], sentAt [' + sentAt + '].');
    }
  }

  const result = sentCount + ' e-ticket berhasil dikirim dari ' + eligibleCount + ' order yang memenuhi syarat. Status lunas terbaca: ' + paidCount + '.';
  Logger.log(result);
  return result;
}

/** Logs the exact email/status fields to diagnose an order that was not sent. */
function cekOrderLunas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Tab Orders belum ditemukan.');

  const lastRow = sheet.getLastRow();
  const statusColumn = findHeaderColumn(sheet, 'Status Pembayaran');
  const sentColumn = findHeaderColumn(sheet, 'E-Ticket Email Sent At');
  for (let rowNumber = 2; rowNumber <= lastRow; rowNumber++) {
    const row = sheet.getRange(rowNumber, 1, 1, HEADERS.length).getValues()[0];
    const rawStatus = sheet.getRange(rowNumber, statusColumn).getValue();
    const status = normalizeStatus(rawStatus);
    if (isPaidStatus(status)) {
      Logger.log('Baris %s | status=%s | email=%s | order=%s | kode=%s | sentAt=%s',
        rowNumber, status, row[3], row[1], row[9], row[sentColumn - 1]);
    } else if (rawStatus) {
      Logger.log('Baris %s status tidak dikenali: [%s] -> [%s]', rowNumber, rawStatus, status);
    }
  }
  Logger.log('Sisa kuota kirim Gmail hari ini: ' + MailApp.getRemainingDailyQuota());
}

/** Tests the external QR service used by the e-ticket email. */
function tesQR() {
  const response = UrlFetchApp.fetch('https://quickchart.io/qr?size=120&text=HOB-TEST', {
    muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Layanan QR mengembalikan HTTP ' + code + '.');
  }
  Logger.log('QR service OK. HTTP ' + code + ', ukuran ' + response.getBlob().getBytes().length + ' bytes.');
}

function findHeaderColumn(sheet, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const column = headers.findIndex(function(header) {
    return String(header).trim().toLowerCase() === headerName.toLowerCase();
  }) + 1;
  if (!column) throw new Error('Kolom ' + headerName + ' tidak ditemukan. Jalankan setupSheet().');
  return column;
}

function normalizeStatus(value) {
  return String(value == null ? '' : value).trim().toUpperCase().replace(/[^A-Z]/g, '');
}

function isPaidStatus(status) {
  return status === 'LUNAS' || status === 'SUCCESS' || status === 'PAID';
}

function isEmailSentMarker(value) {
  if (!value || isPaidStatus(normalizeStatus(value))) return false;
  return value instanceof Date || String(value).trim() !== '';
}

function sendETicketEmail(sheet, rowNumber, sentColumn) {
  if (!sheet || !rowNumber) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
    const activeRange = spreadsheet.getActiveSheet().getActiveRange();
    rowNumber = activeRange ? activeRange.getRow() : 0;
    sentColumn = HEADERS.indexOf('E-Ticket Email Sent At') + 1;
  }
  if (!sheet || sheet.getName() !== CONFIG.SHEET_NAME || rowNumber < 2) {
    throw new Error('Buka tab Orders dan klik sel pada baris order, atau jalankan kirimSemuaETicketLunas.');
  }

  const row = sheet.getRange(rowNumber, 1, 1, HEADERS.length).getValues()[0];
  const orderId = String(row[1]);
  const name = String(row[2]);
  const email = String(row[3]).trim();
  const whatsapp = String(row[4]);
  const category = String(row[5]);
  const quantity = Number(row[6]) || 0;
  const total = Number(row[7]) || 0;
  const ticketCode = String(row[9]);

  if (!email || !ticketCode || !orderId) {
    throw new Error('Email, Order ID, atau kode e-ticket belum tersedia.');
  }

  const qrResponse = UrlFetchApp.fetch(
    'https://quickchart.io/qr?size=300&margin=2&text=' + encodeURIComponent(ticketCode),
    { muteHttpExceptions: true }
  );
  const qrStatus = qrResponse.getResponseCode();
  if (qrStatus < 200 || qrStatus >= 300) {
    throw new Error('QR code gagal dibuat. Layanan QR mengembalikan HTTP ' + qrStatus + '.');
  }
  const qrBlob = qrResponse.getBlob().setName(orderId + '-QR.png');
  const formattedTotal = formatCurrency(total);
  const subject = 'E-Ticket ' + orderId + ' - HOB';
  const htmlBody = '<div style="margin:0;background:#f1f1ef;padding:28px 12px;font-family:Arial,Helvetica,sans-serif;color:#171717">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="max-width:620px;margin:0 auto">' +
    '<tr><td style="background:#111;padding:22px 26px;border-bottom:5px solid #ff5400">' +
    '<span style="font-size:32px;line-height:1;font-weight:900;letter-spacing:2px;color:#ff5400">HOB</span>' +
    '<span style="float:right;color:#f4f0e8;font-size:11px;letter-spacing:2px;font-weight:bold;padding-top:9px">OFFICIAL E-TICKET</span></td></tr>' +
    '<tr><td style="background:#fff;padding:32px 28px 24px">' +
    '<p style="margin:0;color:#ff5400;font-size:12px;font-weight:bold;letter-spacing:2px">PAYMENT CONFIRMED</p>' +
    '<h1 style="margin:10px 0 8px;font-size:30px;line-height:1.1;color:#111">Tiket kamu sudah siap.</h1>' +
    '<p style="margin:0;color:#666;font-size:15px;line-height:1.6">Halo ' + escapeHtml(name) + ', pembayaran kamu telah dikonfirmasi. Tunjukkan QR code ini saat masuk.</p>' +
    '<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="margin:26px 0 0;border:1px solid #e3e3e3">' +
    '<tr><td style="padding:14px 16px;border-bottom:1px solid #e3e3e3;color:#777;font-size:12px">ORDER ID</td><td style="padding:14px 16px;border-bottom:1px solid #e3e3e3;text-align:right;font-weight:bold;font-size:13px">' + escapeHtml(orderId) + '</td></tr>' +
    '<tr><td style="padding:14px 16px;border-bottom:1px solid #e3e3e3;color:#777;font-size:12px">NAMA</td><td style="padding:14px 16px;border-bottom:1px solid #e3e3e3;text-align:right;font-weight:bold;font-size:13px">' + escapeHtml(name) + '</td></tr>' +
    '<tr><td style="padding:14px 16px;border-bottom:1px solid #e3e3e3;color:#777;font-size:12px">EMAIL</td><td style="padding:14px 16px;border-bottom:1px solid #e3e3e3;text-align:right;font-size:13px">' + escapeHtml(email) + '</td></tr>' +
    '<tr><td style="padding:14px 16px;border-bottom:1px solid #e3e3e3;color:#777;font-size:12px">WHATSAPP</td><td style="padding:14px 16px;border-bottom:1px solid #e3e3e3;text-align:right;font-size:13px">' + escapeHtml(whatsapp) + '</td></tr>' +
    '<tr><td style="padding:14px 16px;border-bottom:1px solid #e3e3e3;color:#777;font-size:12px">JUMLAH TIKET</td><td style="padding:14px 16px;border-bottom:1px solid #e3e3e3;text-align:right;font-weight:bold;font-size:13px">' + quantity + ' tiket</td></tr>' +
    '<tr><td style="padding:14px 16px;color:#777;font-size:12px">TOTAL</td><td style="padding:14px 16px;text-align:right;color:#ff5400;font-weight:bold;font-size:16px">' + formattedTotal + '</td></tr></table>' +
    '<div style="margin:26px 0 0;padding:24px;text-align:center;background:#f7f7f5;border:2px dashed #d8d8d4">' +
    '<p style="margin:0 0 14px;color:#555;font-size:12px;font-weight:bold;letter-spacing:1px">SCAN QR CODE SAAT CHECK-IN</p>' +
    '<img src="cid:ticketQr" alt="QR Code e-ticket" width="260" height="260" style="display:block;width:260px;height:260px;margin:0 auto;background:#fff">' +
    '<p style="margin:14px 0 0;color:#111;font-size:13px;font-weight:bold;letter-spacing:1px">' + escapeHtml(ticketCode) + '</p></div>' +
    '<p style="margin:24px 0 0;color:#777;font-size:12px;line-height:1.6;text-align:center">Simpan email ini. QR code ini bersifat pribadi dan hanya dapat digunakan sesuai jumlah tiket yang dibeli.</p>' +
    '</td></tr><tr><td style="padding:18px 26px;background:#111;color:#888;font-size:11px;text-align:center">HOB / OFFICIAL TICKETING</td></tr></table></div>';
  const plainBody = 'Pembayaran berhasil. Order ID: ' + orderId +
    '. Kode e-ticket: ' + ticketCode + '. Total: ' + formattedTotal + '.';

  GmailApp.sendEmail(email, subject, plainBody, {
    htmlBody: htmlBody,
    inlineImages: { ticketQr: qrBlob }
  });
  sheet.getRange(rowNumber, sentColumn).setValue(new Date());
}

function formatCurrency(value) {
  return 'Rp' + Math.round(value).toLocaleString('id-ID');
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, function(character) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[character];
  });
}

function getOrdersSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet || sheet.getLastRow() === 0) {
    throw new Error('Sheet Orders belum siap. Jalankan setupSheet() terlebih dahulu.');
  }
  return sheet;
}

function parsePayload(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('Payload JSON tidak ditemukan.');
  }

  try {
    return JSON.parse(e.postData.contents);
  } catch (error) {
    throw new Error('Payload harus berupa JSON yang valid.');
  }
}

function validateOrder(name, email, whatsapp, category, quantity, ticket, paymentMethod, paymentProof) {
  if (!name || name.length < 2) throw new Error('Nama lengkap wajib diisi.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Format email tidak valid.');
  if (!whatsapp || whatsapp.length < 8) throw new Error('Nomor WhatsApp tidak valid.');
  if (!ticket) throw new Error('Kategori tiket tidak tersedia.');
  if (!Number.isInteger(quantity) || quantity < 1 || quantity > 10) {
    throw new Error('Jumlah tiket harus berupa angka 1 sampai 10.');
  }
  if (paymentMethod !== 'QRIS' && paymentMethod !== 'TRANSFER_BANK') {
    throw new Error('Metode pembayaran tidak valid.');
  }
  if (!paymentProof || !paymentProof.data || !paymentProof.namaFile) {
    throw new Error('Bukti pembayaran wajib diunggah.');
  }
  if (Number(paymentProof.ukuran) > 5000000) {
    throw new Error('Ukuran bukti pembayaran maksimal 5 MB.');
  }
}

function parseQuantity(value) {
  const quantity = Number(value);
  return Number.isInteger(quantity) ? quantity : NaN;
}

function cleanText(value) {
  return String(value == null ? '' : value).trim().replace(/[\r\n]/g, ' ');
}

function createOrderId() {
  const date = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd');
  const suffix = Utilities.getUuid().replace(/-/g, '').substring(0, 4).toUpperCase();
  return 'GIGS-' + date + '-' + suffix;
}

function savePaymentProof(paymentProof, orderId) {
  const match = String(paymentProof.data).match(/^data:([^;]+);base64,(.+)$/);
  if (!match) throw new Error('Format bukti pembayaran tidak valid.');

  const folderIterator = DriveApp.getFoldersByName(CONFIG.PAYMENT_PROOF_FOLDER);
  const folder = folderIterator.hasNext()
    ? folderIterator.next()
    : DriveApp.createFolder(CONFIG.PAYMENT_PROOF_FOLDER);
  const extension = String(paymentProof.namaFile).split('.').pop().replace(/[^a-z0-9]/gi, '').toLowerCase() || 'bin';
  const blob = Utilities.newBlob(Utilities.base64Decode(match[2]), match[1], orderId + '.' + extension);
  return folder.createFile(blob).getUrl();
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}