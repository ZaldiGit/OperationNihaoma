/*
Tambahkan patch ini ke code.gs yang sekarang.
1) Tambah case 'delete_student' pada doPost.
2) Tempel semua function di bawah ini.
*/

// ---- 1. Tambahkan di switch(action) pada doPost ----
// case 'delete_student':
//   result = deleteStudent_(payload);
//   break;


// ---- 2. Function delete student ----
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
