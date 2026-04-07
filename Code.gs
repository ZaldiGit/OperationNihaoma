const CONFIG = {
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID',
  ROOT_FOLDER_ID: 'YOUR_ROOT_FOLDER_ID',
  WRITE_TOKEN: 'CHANGE_ME_TO_A_LONG_RANDOM_STRING',
  TIMEZONE: 'Asia/Jakarta',
};

const SHEETS = {
  students: 'students_master',
  progress: 'student_progress_log',
  documents: 'student_documents',
  invoices: 'student_invoices',
  payments: 'invoice_payment_log',
  programPrices: 'program_prices',
  statusProses: 'status_proses_ref',
  statusVerifikasi: 'status_verifikasi_ref',
  statusPelunasan: 'status_pelunasan_ref',
  statusPengiriman: 'status_pengiriman_ref',
  prioritas: 'prioritas_ref',
  nextAction: 'next_action_ref',
  requiredDocTypes: 'required_doc_types_ref',
  metodePembayaran: 'metode_pembayaran_ref',
};

const HEADERS = {
  students: [
    'student_id','tanggal_input','nama_lengkap','nama_panggilan','jenis_kelamin',
    'tanggal_lahir','kewarganegaraan','no_whatsapp','email','alamat',
    'no_paspor_atau_nik','program_diminati','kampus_tujuan','kota_tujuan',
    'negara_tujuan','intake','durasi_program','estimasi_biaya','sumber_leads',
    'pic_admin','status_proses','tanggal_follow_up_terakhir','next_action',
    'tanggal_next_action','prioritas','catatan_admin','is_active'
  ],
  progress: [
    'log_id','student_id','tanggal_update','updated_by','status_lama',
    'status_baru','catatan','next_action','tanggal_next_action'
  ],
  documents: [
    'doc_id','student_id','nama_mahasiswa','jenis_dokumen','nama_file',
    'link_file','storage_path','tanggal_upload','uploaded_by',
    'status_verifikasi','verified_by','tanggal_verifikasi','catatan_verifikasi',
    'versi_dokumen'
  ],
  invoices: [
    'invoice_id','student_id','nama_mahasiswa','kode_invoice','tanggal_invoice',
    'program','deskripsi_biaya','mata_uang','harga_program','sudah_dibayar',
    'sisa_tagihan','status_pelunasan','status_pengiriman','tanggal_kirim',
    'bukti_pembayaran_link','catatan_invoice'
  ],
  payments: [
    'payment_id','invoice_id','student_id','tanggal_pembayaran',
    'jumlah_pembayaran','metode_pembayaran','bukti_pembayaran_link',
    'dicatat_oleh','catatan'
  ],
};

function doGet(e) {
  try {
    validateToken_(e.parameter.token);
    const action = e.parameter.action || 'bootstrap';
    if (action === 'bootstrap') {
      return json_({ ok: true, ...getBootstrap_() });
    }
    return json_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json_({ ok: false, error: err.message || String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const payload = JSON.parse((e.postData && e.postData.contents) || '{}');
    validateToken_(payload.token);
    const action = payload.action;
    let result = {};
    switch (action) {
      case 'add_student':
        result = addStudent_(payload);
        break;
      case 'update_progress':
        result = updateProgress_(payload);
        break;
      case 'upload_document':
        result = uploadDocument_(payload);
        break;
      case 'create_invoice':
        result = createInvoice_(payload);
        break;
      case 'record_payment':
        result = recordPayment_(payload);
        break;
      default:
        throw new Error('Unknown action');
    }
    return json_({ ok: true, ...result });
  } catch (err) {
    return json_({ ok: false, error: err.message || String(err) });
  } finally {
    lock.releaseLock();
  }
}

function setupNihaomaSheets() {
  const ss = getSpreadsheet_();
  ensureSheet_(ss, SHEETS.students, HEADERS.students);
  ensureSheet_(ss, SHEETS.progress, HEADERS.progress);
  ensureSheet_(ss, SHEETS.documents, HEADERS.documents);
  ensureSheet_(ss, SHEETS.invoices, HEADERS.invoices);
  ensureSheet_(ss, SHEETS.payments, HEADERS.payments);

  writeSingleColumnRef_(ss, SHEETS.statusProses, [
    'New Lead','Follow Up','Interested','Dokumen Awal Masuk','Siap Daftar',
    'Sudah Daftar','Menunggu Pembayaran','Proses Visa','Siap Berangkat',
    'Aktif','Selesai','Cancel'
  ]);
  writeSingleColumnRef_(ss, SHEETS.statusVerifikasi, ['Belum Dicek','Valid','Revisi','Tidak Berlaku']);
  writeSingleColumnRef_(ss, SHEETS.statusPelunasan, ['Belum Lunas','Sebagian','Lunas']);
  writeSingleColumnRef_(ss, SHEETS.statusPengiriman, ['Belum Dikirim','Sudah Dikirim']);
  writeSingleColumnRef_(ss, SHEETS.prioritas, ['Tinggi','Sedang','Rendah']);
  writeSingleColumnRef_(ss, SHEETS.nextAction, [
    'Follow up WA','Follow up Email','Minta dokumen','Review dokumen',
    'Jadwalkan konsultasi','Buat invoice','Kirim reminder pembayaran','Update status'
  ]);
  writeSingleColumnRef_(ss, SHEETS.requiredDocTypes, [
    'Paspor','KTP / ID','Ijazah','Transkrip','Sertifikat Bahasa',
    'Foto Formal','Bukti Pembayaran','Surat Pernyataan'
  ]);
  writeSingleColumnRef_(ss, SHEETS.metodePembayaran, ['Transfer','Cash','QRIS','Lainnya']);

  const priceSheet = ensureSheet_(ss, SHEETS.programPrices, ['program_diminati', 'estimasi_biaya']);
  if (priceSheet.getLastRow() === 1) {
    priceSheet.getRange(2, 1, 5, 2).setValues([
      ['Chinese Language Program', 15000000],
      ['Bachelor Program', 35000000],
      ['Master Program', 45000000],
      ['PhD Program', 55000000],
      ['Short Course', 8000000],
    ]);
  }
  return 'Setup selesai.';
}

function getBootstrap_() {
  const ss = getSpreadsheet_();
  const studentsSheet = ensureSheet_(ss, SHEETS.students, HEADERS.students);
  const progressSheet = ensureSheet_(ss, SHEETS.progress, HEADERS.progress);
  const documentsSheet = ensureSheet_(ss, SHEETS.documents, HEADERS.documents);
  const invoicesSheet = ensureSheet_(ss, SHEETS.invoices, HEADERS.invoices);
  const paymentsSheet = ensureSheet_(ss, SHEETS.payments, HEADERS.payments);

  return {
    students: getObjects_(studentsSheet),
    progress: getObjects_(progressSheet),
    documents: getObjects_(documentsSheet),
    invoices: getObjects_(invoicesSheet),
    payments: getObjects_(paymentsSheet),
    references: {
      status_proses: getSingleColumnRef_(ss, SHEETS.statusProses),
      status_verifikasi: getSingleColumnRef_(ss, SHEETS.statusVerifikasi),
      status_pelunasan: getSingleColumnRef_(ss, SHEETS.statusPelunasan),
      status_pengiriman: getSingleColumnRef_(ss, SHEETS.statusPengiriman),
      prioritas: getSingleColumnRef_(ss, SHEETS.prioritas),
      next_action: getSingleColumnRef_(ss, SHEETS.nextAction),
      required_doc_types: getSingleColumnRef_(ss, SHEETS.requiredDocTypes),
      metode_pembayaran: getSingleColumnRef_(ss, SHEETS.metodePembayaran),
      program_prices: getProgramPrices_(ss),
    },
    meta: {
      spreadsheet_name: ss.getName(),
      spreadsheet_id: ss.getId(),
      generated_at: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
    },
  };
}

function addStudent_(payload) {
  const ss = getSpreadsheet_();
  const sheet = ensureSheet_(ss, SHEETS.students, HEADERS.students);
  const progressSheet = ensureSheet_(ss, SHEETS.progress, HEADERS.progress);

  const studentId = nextId_(sheet, 'STD-', 4, 'student_id');
  const estimasiBiaya = lookupProgramPrice_(ss, payload.program_diminati);
  const row = {
    student_id: studentId,
    tanggal_input: today_(),
    nama_lengkap: payload.nama_lengkap || '',
    nama_panggilan: payload.nama_panggilan || '',
    jenis_kelamin: payload.jenis_kelamin || '',
    tanggal_lahir: payload.tanggal_lahir || '',
    kewarganegaraan: payload.kewarganegaraan || '',
    no_whatsapp: payload.no_whatsapp || '',
    email: payload.email || '',
    alamat: payload.alamat || '',
    no_paspor_atau_nik: payload.no_paspor_atau_nik || '',
    program_diminati: payload.program_diminati || '',
    kampus_tujuan: payload.kampus_tujuan || '',
    kota_tujuan: payload.kota_tujuan || '',
    negara_tujuan: payload.negara_tujuan || '',
    intake: payload.intake || '',
    durasi_program: payload.durasi_program || '',
    estimasi_biaya: estimasiBiaya,
    sumber_leads: payload.sumber_leads || '',
    pic_admin: payload.pic_admin || '',
    status_proses: payload.status_proses || 'New Lead',
    tanggal_follow_up_terakhir: '',
    next_action: payload.next_action || '',
    tanggal_next_action: payload.tanggal_next_action || '',
    prioritas: payload.prioritas || 'Sedang',
    catatan_admin: payload.catatan_admin || '',
    is_active: payload.is_active || 'TRUE',
  };
  appendObject_(sheet, row, HEADERS.students);

  const logId = nextId_(progressSheet, 'LOG-', 4, 'log_id');
  appendObject_(progressSheet, {
    log_id: logId,
    student_id: studentId,
    tanggal_update: now_(),
    updated_by: payload.pic_admin || 'Admin',
    status_lama: '',
    status_baru: row.status_proses,
    catatan: payload.catatan_admin || '',
    next_action: row.next_action,
    tanggal_next_action: row.tanggal_next_action,
  }, HEADERS.progress);

  return { student_id: studentId, estimasi_biaya: estimasiBiaya };
}

function updateProgress_(payload) {
  const ss = getSpreadsheet_();
  const studentsSheet = ensureSheet_(ss, SHEETS.students, HEADERS.students);
  const progressSheet = ensureSheet_(ss, SHEETS.progress, HEADERS.progress);

  const students = getObjects_(studentsSheet);
  const studentRowIndex = findRowIndexByValue_(studentsSheet, 'student_id', payload.student_id);
  if (studentRowIndex < 2) throw new Error('Student tidak ditemukan');

  const current = students.find(r => String(r.student_id) === String(payload.student_id)) || {};
  updateCellByHeader_(studentsSheet, studentRowIndex, 'status_proses', payload.status_baru || current.status_proses || '');
  updateCellByHeader_(studentsSheet, studentRowIndex, 'next_action', payload.next_action || '');
  updateCellByHeader_(studentsSheet, studentRowIndex, 'tanggal_next_action', payload.tanggal_next_action || '');
  updateCellByHeader_(studentsSheet, studentRowIndex, 'tanggal_follow_up_terakhir', now_());
  if (payload.catatan) {
    const existing = String(current.catatan_admin || '');
    const combined = existing ? existing + '\n' + payload.catatan : payload.catatan;
    updateCellByHeader_(studentsSheet, studentRowIndex, 'catatan_admin', combined);
  }

  const logId = nextId_(progressSheet, 'LOG-', 4, 'log_id');
  appendObject_(progressSheet, {
    log_id: logId,
    student_id: payload.student_id,
    tanggal_update: now_(),
    updated_by: payload.updated_by || 'Admin',
    status_lama: current.status_proses || '',
    status_baru: payload.status_baru || '',
    catatan: payload.catatan || '',
    next_action: payload.next_action || '',
    tanggal_next_action: payload.tanggal_next_action || '',
  }, HEADERS.progress);

  return { updated: true };
}

function uploadDocument_(payload) {
  const ss = getSpreadsheet_();
  const sheet = ensureSheet_(ss, SHEETS.documents, HEADERS.documents);
  const docId = nextId_(sheet, 'DOC-', 4, 'doc_id');

  const bytes = Utilities.base64Decode(payload.file_base64);
  const blob = Utilities.newBlob(bytes, payload.mime_type || 'application/octet-stream', payload.nama_file || 'document.bin');
  const folder = getOrCreateStudentFolder_(payload.student_id);
  const file = folder.createFile(blob);
  const url = file.getUrl();
  const storagePath = folder.getName() + '/' + file.getName();

  appendObject_(sheet, {
    doc_id: docId,
    student_id: payload.student_id || '',
    nama_mahasiswa: payload.nama_mahasiswa || '',
    jenis_dokumen: payload.jenis_dokumen || '',
    nama_file: payload.nama_file || '',
    link_file: url,
    storage_path: storagePath,
    tanggal_upload: now_(),
    uploaded_by: payload.uploaded_by || 'Admin',
    status_verifikasi: payload.status_verifikasi || 'Belum Dicek',
    verified_by: '',
    tanggal_verifikasi: '',
    catatan_verifikasi: payload.catatan_verifikasi || '',
    versi_dokumen: payload.versi_dokumen || 'v1',
  }, HEADERS.documents);

  return { doc_id: docId, link_file: url };
}

function createInvoice_(payload) {
  const ss = getSpreadsheet_();
  const sheet = ensureSheet_(ss, SHEETS.invoices, HEADERS.invoices);
  const invoiceId = nextId_(sheet, 'INVROW-', 4, 'invoice_id');
  const kodeInvoice = nextId_(sheet, 'NHEC-', 4, 'kode_invoice');

  const harga = Number(payload.harga_program || 0);
  appendObject_(sheet, {
    invoice_id: invoiceId,
    student_id: payload.student_id || '',
    nama_mahasiswa: payload.nama_mahasiswa || '',
    kode_invoice: kodeInvoice,
    tanggal_invoice: payload.tanggal_invoice || today_(),
    program: payload.program || '',
    deskripsi_biaya: payload.deskripsi_biaya || '',
    mata_uang: payload.mata_uang || 'IDR',
    harga_program: harga,
    sudah_dibayar: 0,
    sisa_tagihan: harga,
    status_pelunasan: 'Belum Lunas',
    status_pengiriman: payload.status_pengiriman || 'Belum Dikirim',
    tanggal_kirim: payload.tanggal_kirim || '',
    bukti_pembayaran_link: '',
    catatan_invoice: payload.catatan_invoice || '',
  }, HEADERS.invoices);

  return { invoice_id: invoiceId, kode_invoice: kodeInvoice };
}

function recordPayment_(payload) {
  const ss = getSpreadsheet_();
  const paymentsSheet = ensureSheet_(ss, SHEETS.payments, HEADERS.payments);
  const invoicesSheet = ensureSheet_(ss, SHEETS.invoices, HEADERS.invoices);

  const paymentId = nextId_(paymentsSheet, 'PAY-', 4, 'payment_id');
  appendObject_(paymentsSheet, {
    payment_id: paymentId,
    invoice_id: payload.invoice_id || '',
    student_id: payload.student_id || '',
    tanggal_pembayaran: payload.tanggal_pembayaran || today_(),
    jumlah_pembayaran: Number(payload.jumlah_pembayaran || 0),
    metode_pembayaran: payload.metode_pembayaran || 'Transfer',
    bukti_pembayaran_link: payload.bukti_pembayaran_link || '',
    dicatat_oleh: payload.dicatat_oleh || 'Finance',
    catatan: payload.catatan || '',
  }, HEADERS.payments);

  refreshInvoicePaymentSummary_(payload.invoice_id);
  return { payment_id: paymentId };
}

function refreshInvoicePaymentSummary_(invoiceId) {
  const ss = getSpreadsheet_();
  const invoicesSheet = ensureSheet_(ss, SHEETS.invoices, HEADERS.invoices);
  const paymentsSheet = ensureSheet_(ss, SHEETS.payments, HEADERS.payments);

  const invoiceRowIndex = findRowIndexByValue_(invoicesSheet, 'invoice_id', invoiceId);
  if (invoiceRowIndex < 2) throw new Error('Invoice tidak ditemukan');

  const invoices = getObjects_(invoicesSheet);
  const invoice = invoices.find(r => String(r.invoice_id) === String(invoiceId));
  const payments = getObjects_(paymentsSheet).filter(r => String(r.invoice_id) === String(invoiceId));
  const totalPaid = payments.reduce((sum, r) => sum + Number(r.jumlah_pembayaran || 0), 0);
  const harga = Number(invoice.harga_program || 0);
  const sisa = Math.max(harga - totalPaid, 0);
  const status = sisa <= 0 && harga > 0 ? 'Lunas' : totalPaid > 0 ? 'Sebagian' : 'Belum Lunas';
  const latestProof = payments.length ? payments[payments.length - 1].bukti_pembayaran_link : '';

  updateCellByHeader_(invoicesSheet, invoiceRowIndex, 'sudah_dibayar', totalPaid);
  updateCellByHeader_(invoicesSheet, invoiceRowIndex, 'sisa_tagihan', sisa);
  updateCellByHeader_(invoicesSheet, invoiceRowIndex, 'status_pelunasan', status);
  updateCellByHeader_(invoicesSheet, invoiceRowIndex, 'bukti_pembayaran_link', latestProof);
}

function getSpreadsheet_() {
  if (!CONFIG.SPREADSHEET_ID || CONFIG.SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID') {
    throw new Error('CONFIG.SPREADSHEET_ID belum diisi.');
  }
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else {
    const currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
    const needReset = headers.some((h, i) => String(currentHeaders[i] || '') !== h);
    if (needReset) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function getObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values[0];
  return values.slice(1).filter(row => row.some(v => String(v) !== '')).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      const val = row[i];
      obj[h] = val instanceof Date
        ? Utilities.formatDate(val, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss')
        : val;
    });
    return obj;
  });
}

function appendObject_(sheet, obj, headers) {
  const row = headers.map(h => obj[h] !== undefined ? obj[h] : '');
  sheet.appendRow(row);
}

function nextId_(sheet, prefix, width, headerName) {
  const rows = getObjects_(sheet);
  let maxNum = 0;
  rows.forEach(r => {
    const val = String(r[headerName] || '');
    const m = val.match(/(\d+)$/);
    if (m) maxNum = Math.max(maxNum, Number(m[1]));
  });
  const nextNum = maxNum + 1;
  return prefix + String(nextNum).padStart(width, '0');
}

function findRowIndexByValue_(sheet, headerName, value) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf(headerName) + 1;
  if (colIndex < 1) throw new Error('Header tidak ditemukan: ' + headerName);
  const values = sheet.getRange(2, colIndex, Math.max(sheet.getLastRow() - 1, 0), 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(value)) return i + 2;
  }
  return -1;
}

function updateCellByHeader_(sheet, rowIndex, headerName, value) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf(headerName) + 1;
  if (colIndex < 1) throw new Error('Header tidak ditemukan: ' + headerName);
  sheet.getRange(rowIndex, colIndex).setValue(value);
}

function getSingleColumnRef_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 1) return [];
  return sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().flat().map(String).filter(Boolean);
}

function writeSingleColumnRef_(ss, sheetName, values) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clearContents();
  sheet.getRange(1, 1, values.length, 1).setValues(values.map(v => [v]));
  sheet.hideSheet();
  return sheet;
}

function getProgramPrices_(ss) {
  const sheet = ensureSheet_(ss, SHEETS.programPrices, ['program_diminati', 'estimasi_biaya']);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  return values.slice(1).filter(r => String(r[0]) !== '').map(r => ({
    program_diminati: String(r[0]),
    estimasi_biaya: Number(r[1] || 0),
  }));
}

function lookupProgramPrice_(ss, programName) {
  const rows = getProgramPrices_(ss);
  const found = rows.find(r => String(r.program_diminati) === String(programName));
  return found ? Number(found.estimasi_biaya || 0) : 0;
}

function getOrCreateStudentFolder_(studentId) {
  const baseFolder = (!CONFIG.ROOT_FOLDER_ID || CONFIG.ROOT_FOLDER_ID === 'YOUR_ROOT_FOLDER_ID')
    ? DriveApp.getRootFolder()
    : DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);

  const folderName = String(studentId) + ' - Documents';
  const it = baseFolder.getFoldersByName(folderName);
  if (it.hasNext()) return it.next();
  return baseFolder.createFolder(folderName);
}

function today_() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function now_() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}

function validateToken_(token) {
  if (!CONFIG.WRITE_TOKEN || CONFIG.WRITE_TOKEN === 'CHANGE_ME_TO_A_LONG_RANDOM_STRING') {
    throw new Error('CONFIG.WRITE_TOKEN belum diisi.');
  }
  if (String(token || '') !== String(CONFIG.WRITE_TOKEN)) {
    throw new Error('Token tidak valid.');
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
