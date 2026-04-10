const CONFIG = {
  SPREADSHEET_ID: '1vfR-3sN9kSK6kh2R5AO2wmBjvd2bAEQCPG8LjQaTrw8',
  ROOT_FOLDER_ID: '1qNRyyDJHmfaJ4cAQEuCZNt6cUB06mRgT', // isi jika ingin file dokumen masuk ke folder Drive tertentu
  WRITE_TOKEN: 'nihaoma-live-7f4c9b2e1a8d6f3c5e9a2b7d1c4f8a6',
  TIMEZONE: 'Asia/Jakarta',
  REFERENCES_SHEET: 'REFERENCES',
  FORM_RESPONSES_SHEET: 'Form Responses 1',
};

const SHEETS = {
  students: 'students_master',
  progress: 'student_progress_log',
  documents: 'student_documents',
  invoices: 'student_invoices',
  payments: 'invoice_payment_log',
  programPrices: 'program_prices',
};

const HEADERS = {
  students: [
    'student_id','tanggal_input','nama_lengkap','nama_panggilan','jenis_kelamin',
    'tanggal_lahir','kewarganegaraan','no_whatsapp','email','alamat',
    'no_paspor_atau_nik','program_diminati','kampus_tujuan','kota_tujuan',
    'negara_tujuan','intake','durasi_program','estimasi_biaya','sumber_leads',
    'pic_admin','status_proses','tanggal_follow_up_terakhir','next_action',
    'tanggal_next_action','prioritas','catatan_admin','is_active','source','updated_at'
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

    if (action === 'bootstrap') return json_({ ok: true, ...getBootstrap_() });
    if (action === 'get_students') return json_({ ok: true, data: getStudents_() });
    if (action === 'get_refs') return json_({ ok: true, data: getRefsForApp_() });
    if (action === 'setup') return json_({ ok: true, message: setupNihaomaSheets() });

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
      case 'update_student':
        result = updateStudent_(payload);
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
      case 'sync_form_response':
        result = syncFormResponse_(payload.payload || payload);
        break;
      case 'delete_student':
        result = deleteStudent_(payload);
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
  ensureSheet_(ss, SHEETS.programPrices, ['program_diminati', 'estimasi_biaya']);

  const refSheet = ensureReferencesSheet_(ss);
  if (refSheet.getLastRow() <= 1) {
    refSheet.getRange(2, 1, 12, 15).setValues([
      ['Follow up WhatsApp', 'S1-S3', 28880000, 'New Lead', 'Belum Dicek', 'Belum Lunas', 'Belum Dikirim', 'Tinggi', 'Paspor', 'Transfer', 'TRUE', 'Pria', 'Syifa', 'Februari', 'GForm'],
      ['Kirim brosur', 'Program Bahasa', 15000000, 'Follow Up', 'Valid', 'Sebagian', 'Sudah Dikirim', 'Sedang', 'KTP / ID', 'Cash', 'FALSE', 'Wanita', 'Aldi', 'September', 'Instagram'],
      ['Jadwalkan konsultasi', 'D3', 21000000, 'Interested', 'Revisi', 'Lunas', '', 'Rendah', 'Ijazah', 'QRIS', '', '', 'Rizal', 'Juli', 'Website'],
      ['Minta dokumen', 'D3 + Intern', 23000000, 'Dokumen Awal Masuk', 'Tidak Berlaku', '', '', '', 'Transkrip', 'EDC', '', '', '', 'Maret', 'Referral'],
      ['Review dokumen', 'Camp Program', 12000000, 'Siap Daftar', '', '', '', '', 'Sertifikat Bahasa', 'Lainnya', '', '', '', '', 'Walk In'],
      ['Buat invoice', 'Study Trip', 12000000, 'Sudah Daftar', '', '', '', '', 'Foto Formal', '', '', '', '', '', ''],
      ['Follow up pembayaran', '', '', 'Menunggu Pembayaran', '', '', '', '', 'Bukti Pembayaran', '', '', '', '', '', ''],
      ['Proses visa', '', '', 'Proses Visa', '', '', '', '', 'Surat Pernyataan', '', '', '', '', '', ''],
      ['Final check', '', '', 'Siap Berangkat', '', '', '', '', 'Form Aplikasi', '', '', '', '', '', ''],
      ['Closing / Done', '', '', 'Aktif', '', '', '', '', 'Dokumen Visa', '', '', '', '', '', ''],
      ['', '', '', 'Selesai', '', '', '', '', 'Lainnya', '', '', '', '', '', ''],
      ['', '', '', 'Cancel', '', '', '', '', '', '', '', '', '', '', ''],
    ]);
  }

  const priceSheet = ss.getSheetByName(SHEETS.programPrices);
  if (priceSheet.getLastRow() <= 1) {
    const programPrices = getProgramPricesFromReferences_(ss);
    if (programPrices.length) {
      priceSheet.getRange(2, 1, programPrices.length, 2)
        .setValues(programPrices.map(r => [r.program_diminati, r.estimasi_biaya]));
    }
  }

  return 'Setup selesai di spreadsheet master.';
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
    references: getRefsForApp_(),
    meta: {
      spreadsheet_name: ss.getName(),
      spreadsheet_id: ss.getId(),
      generated_at: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
    },
  };
}

function getStudents_() {
  const ss = getSpreadsheet_();
  const sheet = ensureSheet_(ss, SHEETS.students, HEADERS.students);
  return getObjects_(sheet);
}

function getRefsForApp_() {
  const ss = getSpreadsheet_();
  const ref = getReferencesMap_(ss);
  return {
    status_proses: ref.status_proses,
    status_verifikasi: ref.status_verifikasi,
    status_pelunasan: ref.status_pelunasan,
    status_pengiriman: ref.status_pengiriman,
    prioritas: ref.prioritas,
    next_action: ref.next_action,
    required_doc_types: ref.required_doc_types,
    metode_pembayaran: ref.metode_pembayaran,
    program: ref.program_list,
    program_diminati: ref.program_list,
    pic_admin: ref.pic_admin,
    intake: ref.intake,
    sumber_leads: ref.sumber_leads,
    jenis_kelamin: ref.gender_list,
    yesno_list: ref.yesno_list,
    program_prices: getProgramPrices_(ss),
  };
}

function addStudent_(payload) {
  const ss = getSpreadsheet_();
  const sheet = ensureSheet_(ss, SHEETS.students, HEADERS.students);
  const progressSheet = ensureSheet_(ss, SHEETS.progress, HEADERS.progress);

  const duplicateRow = findDuplicateStudentRow_(sheet, payload.email, payload.no_whatsapp);
  if (duplicateRow > 0) {
    const existingId = getCellValueByHeader_(sheet, duplicateRow, 'student_id');
    return {
      student_id: existingId,
      duplicate: true,
      message: 'Data mahasiswa sudah ada, tidak dibuat ulang.'
    };
  }

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
    sumber_leads: payload.sumber_leads || payload.source || 'streamlit',
    pic_admin: payload.pic_admin || '',
    status_proses: payload.status_proses || 'New Lead',
    tanggal_follow_up_terakhir: '',
    next_action: payload.next_action || '',
    tanggal_next_action: payload.tanggal_next_action || '',
    prioritas: payload.prioritas || 'Sedang',
    catatan_admin: payload.catatan_admin || '',
    is_active: payload.is_active || 'TRUE',
    source: payload.source || 'streamlit',
    updated_at: now_(),
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
    catatan: payload.catatan_admin || 'Data dibuat',
    next_action: row.next_action,
    tanggal_next_action: row.tanggal_next_action,
  }, HEADERS.progress);

  return { student_id: studentId, estimasi_biaya: estimasiBiaya, duplicate: false };
}

function updateStudent_(payload) {
  const ss = getSpreadsheet_();
  const studentsSheet = ensureSheet_(ss, SHEETS.students, HEADERS.students);
  const progressSheet = ensureSheet_(ss, SHEETS.progress, HEADERS.progress);

  const studentId = payload.student_id;
  if (!studentId) throw new Error('student_id wajib diisi');

  const rowIndex = findRowIndexByValue_(studentsSheet, 'student_id', studentId);
  if (rowIndex < 2) throw new Error('Student tidak ditemukan');

  const current = getRowObjectByIndex_(studentsSheet, rowIndex);
  const allowedFields = HEADERS.students.filter(h => ['student_id','tanggal_input'].indexOf(h) === -1);
  const updates = payload.payload || payload.updates || {};

  allowedFields.forEach(header => {
    if (updates[header] !== undefined) {
      updateCellByHeader_(studentsSheet, rowIndex, header, updates[header]);
    }
  });

  if (updates.program_diminati !== undefined) {
    const estimasiBiaya = lookupProgramPrice_(ss, updates.program_diminati);
    updateCellByHeader_(studentsSheet, rowIndex, 'estimasi_biaya', estimasiBiaya);
  }

  updateCellByHeader_(studentsSheet, rowIndex, 'updated_at', now_());

  const oldStatus = String(current.status_proses || '');
  const newStatus = updates.status_proses !== undefined ? String(updates.status_proses || '') : oldStatus;
  const statusChanged = oldStatus !== newStatus;

  if (statusChanged || payload.catatan_progress) {
    const logId = nextId_(progressSheet, 'LOG-', 4, 'log_id');
    appendObject_(progressSheet, {
      log_id: logId,
      student_id: studentId,
      tanggal_update: now_(),
      updated_by: payload.updated_by || updates.pic_admin || 'Admin',
      status_lama: oldStatus,
      status_baru: newStatus,
      catatan: payload.catatan_progress || updates.catatan_admin || 'Update data mahasiswa',
      next_action: updates.next_action !== undefined ? updates.next_action : current.next_action || '',
      tanggal_next_action: updates.tanggal_next_action !== undefined ? updates.tanggal_next_action : current.tanggal_next_action || '',
    }, HEADERS.progress);
  }

  return { updated: true, student_id: studentId };
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
  updateCellByHeader_(studentsSheet, studentRowIndex, 'updated_at', now_());

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
  ensureSheet_(ss, SHEETS.invoices, HEADERS.invoices);

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

function onFormSubmit(e) {
  return syncFormResponseFromEvent_(e);
}

function syncAllFormResponses() {
  const ss = getSpreadsheet_();
  const formSheet = ss.getSheetByName(CONFIG.FORM_RESPONSES_SHEET);
  if (!formSheet || formSheet.getLastRow() < 2) {
    return 'Sheet Form Responses 1 belum ada atau masih kosong.';
  }

  const values = formSheet.getDataRange().getValues();
  const headers = values[0].map(String);
  let processed = 0;

  for (let i = 1; i < values.length; i++) {
    const rowObj = {};
    headers.forEach((h, idx) => rowObj[h] = values[i][idx]);
    syncFormResponse_(rowObj);
    processed++;
  }

  return 'Sinkron selesai. Total baris diproses: ' + processed;
}

function syncFormResponseFromEvent_(e) {
  if (!e || !e.namedValues) throw new Error('Event form tidak valid');
  return syncFormResponse_(e.namedValues);
}

function syncFormResponse_(sourceData) {
  const ss = getSpreadsheet_();
  const studentsSheet = ensureSheet_(ss, SHEETS.students, HEADERS.students);
  const progressSheet = ensureSheet_(ss, SHEETS.progress, HEADERS.progress);

  const payload = normalizeFormPayload_(sourceData);
  if (!payload.nama_lengkap && !payload.email && !payload.no_whatsapp) {
    return { skipped: true, reason: 'Data form tidak cukup' };
  }

  const existingRow = findDuplicateStudentRowByPayload_(studentsSheet, payload);

  if (existingRow > 0) {
    const current = getRowObjectByIndex_(studentsSheet, existingRow);
    const fieldsToUpdate = {
      nama_lengkap: firstNonEmpty_(payload.nama_lengkap, current.nama_lengkap),
      nama_panggilan: firstNonEmpty_(payload.nama_panggilan, current.nama_panggilan),
      jenis_kelamin: firstNonEmpty_(payload.jenis_kelamin, current.jenis_kelamin),
      tanggal_lahir: firstNonEmpty_(payload.tanggal_lahir, current.tanggal_lahir),
      kewarganegaraan: firstNonEmpty_(payload.kewarganegaraan, current.kewarganegaraan),
      no_whatsapp: firstNonEmpty_(payload.no_whatsapp, current.no_whatsapp),
      email: firstNonEmpty_(payload.email, current.email),
      alamat: firstNonEmpty_(payload.alamat, current.alamat),
      no_paspor_atau_nik: firstNonEmpty_(payload.no_paspor_atau_nik, current.no_paspor_atau_nik),
      program_diminati: firstNonEmpty_(payload.program_diminati, current.program_diminati),
      kampus_tujuan: firstNonEmpty_(payload.kampus_tujuan, current.kampus_tujuan),
      kota_tujuan: firstNonEmpty_(payload.kota_tujuan, current.kota_tujuan),
      negara_tujuan: firstNonEmpty_(payload.negara_tujuan, current.negara_tujuan),
      intake: firstNonEmpty_(payload.intake, current.intake),
      durasi_program: firstNonEmpty_(payload.durasi_program, current.durasi_program),
      sumber_leads: firstNonEmpty_(payload.sumber_leads, current.sumber_leads, 'gform'),
      status_proses: current.status_proses || 'New Lead',
      source: 'gform',
      updated_at: now_(),
    };

    Object.keys(fieldsToUpdate).forEach(header => updateCellByHeader_(studentsSheet, existingRow, header, fieldsToUpdate[header]));
    updateCellByHeader_(studentsSheet, existingRow, 'estimasi_biaya', lookupProgramPrice_(ss, fieldsToUpdate.program_diminati));

    const logId = nextId_(progressSheet, 'LOG-', 4, 'log_id');
    appendObject_(progressSheet, {
      log_id: logId,
      student_id: current.student_id,
      tanggal_update: now_(),
      updated_by: 'Google Form',
      status_lama: current.status_proses || '',
      status_baru: current.status_proses || 'New Lead',
      catatan: 'Data dari Google Form diperbarui ke master.',
      next_action: current.next_action || '',
      tanggal_next_action: current.tanggal_next_action || '',
    }, HEADERS.progress);

    return { updated: true, student_id: current.student_id };
  }

  const result = addStudent_({
    nama_lengkap: payload.nama_lengkap,
    nama_panggilan: payload.nama_panggilan,
    jenis_kelamin: payload.jenis_kelamin,
    tanggal_lahir: payload.tanggal_lahir,
    kewarganegaraan: payload.kewarganegaraan,
    no_whatsapp: payload.no_whatsapp,
    email: payload.email,
    alamat: payload.alamat,
    no_paspor_atau_nik: payload.no_paspor_atau_nik,
    program_diminati: payload.program_diminati,
    kampus_tujuan: payload.kampus_tujuan,
    kota_tujuan: payload.kota_tujuan,
    negara_tujuan: payload.negara_tujuan,
    intake: payload.intake,
    durasi_program: payload.durasi_program,
    sumber_leads: payload.sumber_leads || 'gform',
    pic_admin: '',
    status_proses: 'New Lead',
    next_action: 'Follow up WhatsApp',
    prioritas: 'Sedang',
    catatan_admin: 'Masuk dari Google Form',
    is_active: 'TRUE',
    source: 'gform',
  });

  return { inserted: true, student_id: result.student_id };
}

function normalizeFormPayload_(sourceData) {
  const named = normalizeNamedValues_(sourceData);

  const pilihanJurusan = getFirstNamedValue_(named, [
    'pilihan jurusan maksimal 3',
    'pilihan jurusan',
    'jurusan pilihan',
    'program diminati',
    'program',
  ]);

  const jurusanLainnya = getFirstNamedValue_(named, [
    'jurusan lainnya jika memilih lainnya',
    'jurusan lainnya',
    'jurusan lain',
  ]);

  const pilihanUniversitas = getFirstNamedValue_(named, [
    'pilihan universitas di china',
    'pilihan universitas',
    'universitas tujuan',
    'kampus tujuan',
  ]);

  const pilihanKota = getFirstNamedValue_(named, [
    'pilihan kota di china',
    'kota tujuan',
    'city',
  ]);

  const pilihanNegara = getFirstNamedValue_(named, [
    'negara tujuan',
    'country',
  ]);

  const asalSekolah = getFirstNamedValue_(named, [
    'asal sekolah sekolah sebelumnya',
    'asal sekolah',
    'sekolah sebelumnya',
  ]);

  const email = getFirstNamedValue_(named, [
    'email',
    'alamat email',
    'email aktif',
  ]);

  const whatsapp = getFirstNamedValue_(named, [
    'no whatsapp',
    'nomor whatsapp',
    'no wa',
    'nomor wa',
    'nomor whatsapp aktif',
    'whatsapp',
    'phone number',
  ]);

  const jurusanGabungan = combineProgramChoice_(pilihanJurusan, jurusanLainnya);

  return {
    nama_lengkap: getFirstNamedValue_(named, [
      'nama lengkap',
      'nama lengkap sesuai paspor',
      'nama sesuai paspor',
      'nama_lengkap',
      'full name',
      'nama',
    ]),
    nama_panggilan: getFirstNamedValue_(named, [
      'nama panggilan',
      'nama_panggilan',
      'nickname',
    ]),
    jenis_kelamin: getFirstNamedValue_(named, [
      'jenis kelamin',
      'gender',
    ]),
    tanggal_lahir: getFirstNamedValue_(named, [
      'tanggal lahir',
      'tgl lahir',
      'date of birth',
    ]),
    kewarganegaraan: firstNonEmpty_(
      getFirstNamedValue_(named, ['kewarganegaraan', 'nationality']),
      'Indonesia'
    ),
    no_whatsapp: whatsapp,
    email: email,
    alamat: firstNonEmpty_(
      getFirstNamedValue_(named, ['alamat', 'address']),
      asalSekolah
    ),
    no_paspor_atau_nik: getFirstNamedValue_(named, [
      'no paspor atau nik',
      'nomor paspor atau nik',
      'no paspor',
      'nomor paspor',
      'nik',
      'passport id number',
    ]),
    program_diminati: jurusanGabungan,
    kampus_tujuan: pilihanUniversitas,
    kota_tujuan: pilihanKota,
    negara_tujuan: pilihanNegara,
    intake: getFirstNamedValue_(named, [
      'intake',
      'periode intake',
      'periode keberangkatan',
    ]),
    durasi_program: getFirstNamedValue_(named, [
      'durasi program',
      'durasi',
    ]),
    sumber_leads: firstNonEmpty_(
      getFirstNamedValue_(named, [
        'sumber leads',
        'source',
        'darimana mengetahui kami',
        'sumber informasi',
      ]),
      'gform'
    ),
  };
}

function normalizeNamedValues_(sourceData) {(sourceData) {
  const out = {};
  Object.keys(sourceData || {}).forEach(key => {
    const normalized = normalizeHeader_(key);
    let value = sourceData[key];
    if (Array.isArray(value)) value = value[0];
    out[normalized] = String(value || '').trim();
  });
  return out;
}

function getFirstNamedValue_(normalizedMap, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const key = normalizeHeader_(aliases[i]);
    if (normalizedMap[key]) return normalizedMap[key];
  }
  return '';
}

function combineProgramChoice_(pilihanJurusan, jurusanLainnya) {
  const base = String(pilihanJurusan || '').trim();
  const other = String(jurusanLainnya || '').trim();

  if (!base && !other) return '';
  if (!base) return other;

  if (/lainnya/i.test(base) && other) {
    return base.replace(/lainnya/ig, other);
  }

  if (other && base.toLowerCase().indexOf(other.toLowerCase()) === -1) {
    return base + ' | ' + other;
  }
  return base;
}

function findDuplicateStudentRowByPayload_(sheet, payload) {
  let row = findDuplicateStudentRow_(sheet, payload.email, payload.no_whatsapp);
  if (row > 0) return row;

  const students = getObjects_(sheet);
  const nameNorm = normalizeLooseText_(payload.nama_lengkap);
  const dobNorm = normalizeDateLike_(payload.tanggal_lahir);
  const campusNorm = normalizeLooseText_(payload.kampus_tujuan);

  for (let i = 0; i < students.length; i++) {
    const current = students[i];
    const currentName = normalizeLooseText_(current.nama_lengkap);
    const currentDob = normalizeDateLike_(current.tanggal_lahir);
    const currentCampus = normalizeLooseText_(current.kampus_tujuan);

    if (nameNorm && dobNorm && currentName === nameNorm && currentDob === dobNorm) {
      return i + 2;
    }

    if (nameNorm && campusNorm && currentName === nameNorm && currentCampus === campusNorm) {
      return i + 2;
    }
  }
  return -1;
}

function normalizeDateLike_(value) {
  if (!value) return '';
  const d = new Date(value);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  }
  return String(value || '').trim().toLowerCase();
}

function findDuplicateStudentRow_(sheet, email, noWhatsapp) {
  const students = getObjects_(sheet);
  const emailNorm = normalizeLooseText_(email);
  const waNorm = normalizePhone_(noWhatsapp);

  for (let i = 0; i < students.length; i++) {
    const row = students[i];
    if (emailNorm && normalizeLooseText_(row.email) === emailNorm) return i + 2;
    if (waNorm && normalizePhone_(row.no_whatsapp) === waNorm) return i + 2;
  }
  return -1;
}

function normalizeLooseText_(value) {
  return String(value || '').trim().toLowerCase();
}

function normalizePhone_(value) {
  return String(value || '').replace(/[^0-9]/g, '');
}

function firstNonEmpty_() {
  for (let i = 0; i < arguments.length; i++) {
    if (String(arguments[i] || '').trim() !== '') return arguments[i];
  }
  return '';
}

function getSpreadsheet_() {
  if (!CONFIG.SPREADSHEET_ID) throw new Error('CONFIG.SPREADSHEET_ID belum diisi.');
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else {
    const width = Math.max(sheet.getLastColumn(), headers.length);
    const currentHeaders = sheet.getRange(1, 1, 1, width).getValues()[0];
    const needReset = headers.some((h, i) => String(currentHeaders[i] || '') !== h);
    if (needReset) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function ensureReferencesSheet_(ss) {
  const headers = [
    'next_action_list',
    'program_list',
    'price_list',
    'status_proses_list',
    'status_verifikasi_list',
    'status_pelunasan_list',
    'status_pengiriman_list',
    'prioritas_list',
    'doc_type_list',
    'payment_method_list',
    'yesno_list',
    'gender_list',
    'pic_admin_list',
    'intake_list',
    'lead_source_list'
  ];

  let sheet = ss.getSheetByName(CONFIG.REFERENCES_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.REFERENCES_SHEET);

  const existing = sheet.getLastRow() > 0 ? sheet.getRange(1, 1, 1, headers.length).getValues()[0] : [];
  const normalizedExisting = existing.map(normalizeHeader_);
  const normalizedTarget = headers.map(normalizeHeader_);
  const headersMismatch = normalizedTarget.some((h, i) => normalizedExisting[i] !== h);

  if (sheet.getLastRow() === 0 || headersMismatch) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values[0];
  return values
    .slice(1)
    .filter(row => row.some(v => String(v) !== ''))
    .map(row => {
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

function getRowObjectByIndex_(sheet, rowIndex) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const obj = {};
  headers.forEach((h, i) => {
    obj[h] = row[i] instanceof Date
      ? Utilities.formatDate(row[i], CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss')
      : row[i];
  });
  return obj;
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
  const numRows = Math.max(sheet.getLastRow() - 1, 0);
  if (numRows === 0) return -1;
  const values = sheet.getRange(2, colIndex, numRows, 1).getValues();
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

function getCellValueByHeader_(sheet, rowIndex, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf(headerName) + 1;
  if (colIndex < 1) throw new Error('Header tidak ditemukan: ' + headerName);
  return sheet.getRange(rowIndex, colIndex).getValue();
}

function getReferencesMap_(ss) {
  const sheet = ensureReferencesSheet_(ss);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return emptyReferences_();

  const headers = values[0].map(String);
  const rows = values.slice(1);

  return {
    next_action: getColumnByCandidates_(headers, rows, ['next_action_list', 'nextactionlist'], 0),
    program_list: getColumnByCandidates_(headers, rows, ['program_list', 'programlist'], 1),
    price_list: getNumberColumnByCandidates_(headers, rows, ['price_list', 'pricelist'], 2),
    status_proses: getColumnByCandidates_(headers, rows, ['status_proses_list', 'statusproseslist'], 3),
    status_verifikasi: getColumnByCandidates_(headers, rows, ['status_verifikasi_list', 'statusverifikasilist', 'status_ve'], 4),
    status_pelunasan: getColumnByCandidates_(headers, rows, ['status_pelunasan_list', 'statuspelunasanlist'], 5),
    status_pengiriman: getColumnByCandidates_(headers, rows, ['status_pengiriman_list', 'statuspengirimanlist'], 6),
    prioritas: getColumnByCandidates_(headers, rows, ['prioritas_list', 'prioritaslist'], 7),
    required_doc_types: getColumnByCandidates_(headers, rows, ['doc_type_list', 'doctypelist'], 8),
    metode_pembayaran: getColumnByCandidates_(headers, rows, ['payment_method_list', 'paymentmethodlist', 'payment_yesno_list'], 9),
    yesno_list: getColumnByCandidates_(headers, rows, ['yesno_list', 'yesnolist'], 10),
    gender_list: getColumnByCandidates_(headers, rows, ['gender_list', 'genderlist'], 11),
    pic_admin: getColumnByCandidates_(headers, rows, ['pic_admin_list', 'picadminlist'], 12),
    intake: getColumnByCandidates_(headers, rows, ['intake_list', 'intakelist'], 13),
    sumber_leads: getColumnByCandidates_(headers, rows, ['lead_source_list', 'leadsource_list'], 14),
  };
}

function emptyReferences_() {
  return {
    next_action: [],
    program_list: [],
    price_list: [],
    status_proses: [],
    status_verifikasi: [],
    status_pelunasan: [],
    status_pengiriman: [],
    prioritas: [],
    required_doc_types: [],
    metode_pembayaran: [],
    yesno_list: [],
    gender_list: [],
    pic_admin: [],
    intake: [],
    sumber_leads: [],
  };
}

function getColumnByCandidates_(headers, rows, candidates, fallbackIndex) {
  const normalizedHeaders = headers.map(normalizeHeader_);
  let idx = -1;

  for (let i = 0; i < candidates.length; i++) {
    const target = normalizeHeader_(candidates[i]);
    idx = normalizedHeaders.indexOf(target);
    if (idx !== -1) break;
  }

  if (idx === -1 && fallbackIndex >= 0 && fallbackIndex < headers.length) idx = fallbackIndex;
  if (idx === -1) return [];

  const out = [];
  for (let r = 0; r < rows.length; r++) {
    const v = rows[r][idx];
    if (String(v || '').trim() !== '') out.push(String(v).trim());
  }
  return out;
}

function getNumberColumnByCandidates_(headers, rows, candidates, fallbackIndex) {
  const textValues = getColumnByCandidates_(headers, rows, candidates, fallbackIndex);
  return textValues.map(v => Number(v || 0));
}

function normalizeHeader_(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function getProgramPrices_(ss) {
  const sheet = ss.getSheetByName(SHEETS.programPrices);
  if (sheet && sheet.getLastRow() > 1) {
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const programIdx = headers.indexOf('program_diminati');
    const priceIdx = headers.indexOf('estimasi_biaya');
    if (programIdx !== -1 && priceIdx !== -1) {
      return values.slice(1)
        .filter(r => String(r[programIdx] || '').trim() !== '')
        .map(r => ({
          program_diminati: String(r[programIdx]).trim(),
          estimasi_biaya: Number(r[priceIdx] || 0),
        }));
    }
  }
  return getProgramPricesFromReferences_(ss);
}

function getProgramPricesFromReferences_(ss) {
  const ref = getReferencesMap_(ss);
  const programs = ref.program_list || [];
  const prices = ref.price_list || [];
  const out = [];
  for (let i = 0; i < Math.max(programs.length, prices.length); i++) {
    const program = String(programs[i] || '').trim();
    if (!program) continue;
    out.push({
      program_diminati: program,
      estimasi_biaya: Number(prices[i] || 0),
    });
  }
  return out;
}

function lookupProgramPrice_(ss, programName) {
  const rows = getProgramPrices_(ss);
  const found = rows.find(r => String(r.program_diminati) === String(programName));
  return found ? Number(found.estimasi_biaya || 0) : 0;
}

function getOrCreateStudentFolder_(studentId) {
  const baseFolder = CONFIG.ROOT_FOLDER_ID
    ? DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID)
    : DriveApp.getRootFolder();

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
  if (String(token || '') !== String(CONFIG.WRITE_TOKEN)) {
    throw new Error('Token tidak valid.');
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function deleteStudent_(payload) {
  const ss = getSpreadsheet_();
  const studentId = String(payload.student_id || '').trim();
  if (!studentId) throw new Error('student_id wajib diisi');

  const studentsSheet = ensureSheet_(ss, SHEETS.students, HEADERS.students);
  const progressSheet = ensureSheet_(ss, SHEETS.progress, HEADERS.progress);
  const documentsSheet = ensureSheet_(ss, SHEETS.documents, HEADERS.documents);
  const invoicesSheet = ensureSheet_(ss, SHEETS.invoices, HEADERS.invoices);
  const paymentsSheet = ensureSheet_(ss, SHEETS.payments, HEADERS.payments);
  const formSheet = ss.getSheetByName(CONFIG.FORM_RESPONSES_SHEET);

  const studentRow = findRowIndexByValue_(studentsSheet, 'student_id', studentId);
  if (studentRow < 2) throw new Error('Student tidak ditemukan');

  const student = getRowObjectByIndex_(studentsSheet, studentRow);
  const deletedStudent = {
    student_id: student.student_id || '',
    nama_lengkap: student.nama_lengkap || '',
    email: student.email || '',
    no_whatsapp: student.no_whatsapp || ''
  };

  // Hapus sheet turunan dulu supaya referensi masih ada.
  const deletedProgress = deleteRowsByFieldValue_(progressSheet, 'student_id', studentId);
  const deletedDocuments = deleteRowsByFieldValue_(documentsSheet, 'student_id', studentId);
  const deletedInvoices = deleteRowsByFieldValue_(invoicesSheet, 'student_id', studentId);

  // Ambil invoice id yang sudah terhapus untuk ikut hapus payment.
  let deletedPayments = 0;
  if (deletedInvoices.deletedRows && deletedInvoices.deletedRows.length) {
    const invoiceIds = deletedInvoices.deletedRows
      .map(function(r) { return String(r.invoice_id || '').trim(); })
      .filter(function(v) { return v; });
    invoiceIds.forEach(function(invoiceId) {
      const delPay = deleteRowsByFieldValue_(paymentsSheet, 'invoice_id', invoiceId);
      deletedPayments += delPay.deletedCount;
    });
  }

  // Hapus dari students_master paling akhir.
  studentsSheet.deleteRow(studentRow);

  // Opsional: hapus baris terkait di Form Responses 1.
  // Karena form response tidak punya student_id, kita cocokkan berdasarkan nama/email/wa.
  let deletedFormResponses = 0;
  if (formSheet) {
    deletedFormResponses = deleteMatchingFormResponses_(formSheet, deletedStudent);
  }

  return {
    deleted: true,
    student_id: studentId,
    deleted_progress: deletedProgress.deletedCount,
    deleted_documents: deletedDocuments.deletedCount,
    deleted_invoices: deletedInvoices.deletedCount,
    deleted_payments: deletedPayments,
    deleted_form_responses: deletedFormResponses
  };
}

function deleteRowsByFieldValue_(sheet, headerName, targetValue) {
  if (!sheet || sheet.getLastRow() < 2) {
    return { deletedCount: 0, deletedRows: [] };
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf(headerName);
  if (colIndex === -1) {
    return { deletedCount: 0, deletedRows: [] };
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const rowsToDelete = [];
  const deletedRows = [];

  for (var i = 0; i < values.length; i++) {
    const rowValue = String(values[i][colIndex] || '').trim();
    if (rowValue === String(targetValue || '').trim()) {
      rowsToDelete.push(i + 2);
      deletedRows.push(objectFromRow_(headers, values[i]));
    }
  }

  rowsToDelete.reverse().forEach(function(rowIndex) {
    sheet.deleteRow(rowIndex);
  });

  return { deletedCount: rowsToDelete.length, deletedRows: deletedRows };
}

function deleteMatchingFormResponses_(formSheet, student) {
  if (!formSheet || formSheet.getLastRow() < 2) return 0;

  const headers = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];
  const values = formSheet.getRange(2, 1, formSheet.getLastRow() - 1, formSheet.getLastColumn()).getValues();

  const normalizedHeaders = headers.map(function(h) { return normalizeText_(h); });
  const nameIdx = findHeaderIndexByAliases_(normalizedHeaders, [
    'nama lengkap', 'nama lengkap sesuai paspor', 'full name', 'nama'
  ]);
  const emailIdx = findHeaderIndexByAliases_(normalizedHeaders, [
    'email', 'alamat email', 'email aktif'
  ]);
  const waIdx = findHeaderIndexByAliases_(normalizedHeaders, [
    'no whatsapp', 'nomor whatsapp', 'nomor whatsapp aktif', 'no wa', 'nomor wa', 'whatsapp'
  ]);

  const studentName = normalizeText_(student.nama_lengkap || '');
  const studentEmail = normalizeText_(student.email || '');
  const studentWa = normalizePhone_(student.no_whatsapp || '');

  const rowsToDelete = [];

  for (var i = 0; i < values.length; i++) {
    const rowName = nameIdx > -1 ? normalizeText_(values[i][nameIdx]) : '';
    const rowEmail = emailIdx > -1 ? normalizeText_(values[i][emailIdx]) : '';
    const rowWa = waIdx > -1 ? normalizePhone_(values[i][waIdx]) : '';

    var matched = false;

    if (studentEmail && rowEmail && studentEmail === rowEmail) matched = true;
    if (!matched && studentWa && rowWa && studentWa === rowWa) matched = true;
    if (!matched && studentName && rowName && studentName === rowName) matched = true;

    if (matched) {
      rowsToDelete.push(i + 2);
    }
  }

  rowsToDelete.reverse().forEach(function(rowIndex) {
    formSheet.deleteRow(rowIndex);
  });

  return rowsToDelete.length;
}

function objectFromRow_(headers, row) {
  var obj = {};
  headers.forEach(function(h, i) {
    obj[h] = row[i];
  });
  return obj;
}

function findHeaderIndexByAliases_(normalizedHeaders, aliases) {
  for (var a = 0; a < aliases.length; a++) {
    var wanted = normalizeText_(aliases[a]);
    for (var i = 0; i < normalizedHeaders.length; i++) {
      if (normalizedHeaders[i] === wanted) return i;
      if (normalizedHeaders[i].indexOf(wanted) !== -1) return i;
      if (wanted.indexOf(normalizedHeaders[i]) !== -1) return i;
    }
  }
  return -1;
}

function normalizePhone_(value) {
  return String(value || '').replace(/\D/g, '');
}
