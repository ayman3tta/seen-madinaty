/**
 * كود ربط سيستم مدينتي بجوجل شيت (Google Apps Script)
 * يتضمن: المزامنة الفورية + سجل النشاط الكامل + إمكانية التراجع (UNDO) من الشيت أو التطبيق
 * -----------------------------------------------------
 * انسخ هذا الكود بالكامل وضعه في:
 * Google Sheets -> Extensions (التوسيعات) -> Apps Script -> ثم اضغط Deploy (نشر) كـ Web App
 */

const COLS = [
  { key: 'id', title: 'الكود (ID)' },
  { key: 'kind', title: 'القسم' },
  { key: 'addr', title: 'العنوان' },
  { key: 'phase', title: 'المرحلة' },
  { key: 'model', title: 'النموذج والمساحة' },
  { key: 'finish', title: 'التشطيب' },
  { key: 'price', title: 'السعر المطلوب' },
  { key: 'money', title: 'الموقف المالي' },
  { key: 'owner', title: 'اسم المالك' },
  { key: 'name', title: 'اسم العميل' },
  { key: 'phone', title: 'رقم الهاتف' },
  { key: 'status', title: 'الحالة' },
  { key: 'stage', title: 'حالة العميل' },
  { key: 'budget', title: 'الميزانية' },
  { key: 'purpose', title: 'الغرض' },
  { key: 'last', title: 'تاريخ آخر تواصل' },
  { key: 'next', title: 'تاريخ المتابعة القادمة' },
  { key: 'summary', title: 'ملخص آخر إجراء' },
  { key: 'notes', title: 'ملاحظات' }
];

const LOG_COLS = [
  'كود النشاط (Log ID)',
  'التاريخ والوقت',
  'نوع الإجراء',
  'القسم',
  'العنصر',
  'كود العنصر (ID)',
  'البيانات السابقة',
  'البيانات الجديدة',
  'تراجع (Undo)',
  'حالة التراجع'
];

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('البيانات');
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName('البيانات');
  }
  initHeaders(sheet);
  return sheet;
}

function getLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('سجل_النشاط');
  if (!sheet) {
    sheet = ss.insertSheet('سجل_النشاط');
  }
  initLogHeaders(sheet);
  return sheet;
}

function initHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    const titles = COLS.map(c => c.title);
    sheet.appendRow(titles);
    sheet.getRange(1, 1, 1, titles.length)
      .setFontWeight('bold')
      .setBackground('#2563eb')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
}

function initLogHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(LOG_COLS);
    sheet.getRange(1, 1, 1, LOG_COLS.length)
      .setFontWeight('bold')
      .setBackground('#1e293b')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 130);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 90);
    sheet.setColumnWidth(5, 140);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(7, 200);
    sheet.setColumnWidth(8, 200);
    sheet.setColumnWidth(9, 110);
    sheet.setColumnWidth(10, 130);
  }
}

function rowToObject(row) {
  const obj = {};
  COLS.forEach((c, i) => {
    let val = row[i];
    if (val instanceof Date) {
      val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    obj[c.key] = val != null ? String(val) : '';
  });
  return obj;
}

function objectToRow(obj) {
  return COLS.map(c => {
    let val = obj[c.key] != null ? obj[c.key] : '';
    if (c.key === 'phone' && val && String(val).startsWith('0')) {
      return "'" + val;
    }
    return val;
  });
}

function appendLog(logSheet, log) {
  if (!logSheet || !log) return;
  initLogHeaders(logSheet);

  const rowValues = [
    log.logId || ('log_' + Date.now()),
    log.time || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    log.actionName || log.action || 'إجراء',
    log.targetKind || '',
    log.targetTitle || '',
    log.targetId || '',
    log.before ? JSON.stringify(log.before) : '',
    log.after ? JSON.stringify(log.after) : '',
    false,
    log.isUndone ? 'تم التراجع ✅' : 'نشط'
  ];

  // إدراج السطر الجديد دائماً في الأعلى (السطر 2) ليبقى أحدث نشاط أولاً
  logSheet.insertRowAfter(1);
  logSheet.getRange(2, 1, 1, rowValues.length).setValues([rowValues]);

  // تعيين مربع اختيار (Checkbox) لخانة التراجع
  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  logSheet.getRange(2, 9).setDataValidation(checkboxRule);
}

function undoRow(logSheet, dataSheet, rowNum) {
  if (!logSheet || !dataSheet || rowNum <= 1) return false;
  const values = logSheet.getRange(rowNum, 1, 1, LOG_COLS.length).getValues()[0];
  const action = String(values[2] || '');
  const targetId = String(values[5] || '');
  const beforeRaw = values[6];
  const status = String(values[9] || '');

  if (status.includes('تم التراجع')) {
    return false; // تم التراجع عنه مسبقاً
  }

  if (action.includes('إضافة')) {
    // إذا كان إضافة -> التراجع هو حذف العنصر
    deleteById(dataSheet, targetId);
  } else if (action.includes('حذف') || action.includes('تعديل')) {
    // إذا كان حذف أو تعديل -> التراجع هو استعادة البيانات السابقة
    if (beforeRaw) {
      try {
        const prevItem = JSON.parse(beforeRaw);
        saveOrUpdate(dataSheet, prevItem);
      } catch (e) {
        console.error('JSON parse error in undo:', e);
      }
    }
  }

  // تحديث حالة السجل إلى "تم التراجع"
  logSheet.getRange(rowNum, 9).setValue(false);
  logSheet.getRange(rowNum, 10).setValue('تم التراجع بنجاح ✅');
  return true;
}

/**
 * عند تعديل الشيت يدوياً (تفعيل مربع اختيار التراجع)
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== 'سجل_النشاط') return;

    const col = e.range.getColumn();
    const row = e.range.getRow();

    // العمود 9 هو مربع الاختيار الخاص بالتراجع
    if (col === 9 && row > 1 && e.value === 'TRUE') {
      const dataSheet = getSheet();
      const done = undoRow(sheet, dataSheet, row);
      if (done) {
        SpreadsheetApp.getActiveSpreadsheet().toast('تم التراجع عن العملية واستعادة البيانات بنجاح ✅', 'تراجع');
      }
    }
  } catch (err) {
    console.error('onEdit error:', err);
  }
}

/**
 * إنشاء قائمة مخصصة في أعلى جوجل شيت
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('سيستم مدينتي 🏠')
      .addItem('↩️ تراجع عن آخر نشاط (Undo)', 'menuUndoLast')
      .addItem('📋 فحص سجل النشاط', 'menuCheckLogs')
      .addToUi();
  } catch (e) {
    console.log(e);
  }
}

function menuUndoLast() {
  const logSheet = getLogSheet();
  const dataSheet = getSheet();
  const lastRow = logSheet.getLastRow();
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('سجل النشاط فارغ، لا توجد عمليات بعد.');
    return;
  }

  for (let r = 2; r <= lastRow; r++) {
    const status = String(logSheet.getRange(r, 10).getValue());
    if (!status.includes('تم التراجع')) {
      const ok = undoRow(logSheet, dataSheet, r);
      if (ok) {
        SpreadsheetApp.getUi().alert('تم التراجع عن آخر عملية بنجاح واستعادة البيانات إلى الشيت ✅');
        return;
      }
    }
  }
  SpreadsheetApp.getUi().alert('جميع العمليات السابقة تم التراجع عنها بالفعل.');
}

function menuCheckLogs() {
  const logSheet = getLogSheet();
  SpreadsheetApp.setActiveSheet(logSheet);
}

function doGet(e) {
  try {
    const sheet = getSheet();
    const logSheet = getLogSheet();
    const action = e && e.parameter ? e.parameter.action : 'getAll';
    
    if (action === 'getAll') {
      const lastRow = sheet.getLastRow();
      let data = [];
      if (lastRow > 1) {
        const rows = sheet.getRange(2, 1, lastRow - 1, COLS.length).getValues();
        data = rows.map(rowToObject).filter(x => x.id);
      }
      return jsonResponse({ success: true, data: data });
    }

    if (action === 'getLogs') {
      const lastRow = logSheet.getLastRow();
      let logs = [];
      if (lastRow > 1) {
        const rows = logSheet.getRange(2, 1, lastRow - 1, LOG_COLS.length).getValues();
        logs = rows.map(r => {
          let b = null, a = null;
          try { b = r[6] ? JSON.parse(r[6]) : null; } catch(_) {}
          try { a = r[7] ? JSON.parse(r[7]) : null; } catch(_) {}
          return {
            logId: r[0],
            time: r[1],
            actionName: r[2],
            targetKind: r[3],
            targetTitle: r[4],
            targetId: r[5],
            before: b,
            after: a,
            isUndone: String(r[9]).includes('تم التراجع')
          };
        });
      }
      return jsonResponse({ success: true, logs: logs });
    }

    if (action === 'delLog') {
      const logId = e.parameter.logId;
      if (logId) deleteLogById(logSheet, logId);
      return jsonResponse({ success: true });
    }
    
    if (action === 'del') {
      const id = e.parameter.id;
      deleteById(sheet, id);
      return jsonResponse({ success: true });
    }

    return jsonResponse({ success: true, message: 'Ready' });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function doPost(e) {
  try {
    const sheet = getSheet();
    const logSheet = getLogSheet();
    let body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else if (e && e.parameter && e.parameter.data) {
      body = JSON.parse(e.parameter.data);
    }
    
    const action = body.action || (e && e.parameter ? e.parameter.action : 'put');

    if (action === 'put') {
      saveOrUpdate(sheet, body.item);
      if (body.log) appendLog(logSheet, body.log);
      return jsonResponse({ success: true });
    }

    if (action === 'del') {
      deleteById(sheet, body.id || (e && e.parameter ? e.parameter.id : ''));
      if (body.log) appendLog(logSheet, body.log);
      return jsonResponse({ success: true });
    }

    if (action === 'log') {
      if (body.log) appendLog(logSheet, body.log);
      return jsonResponse({ success: true });
    }

    if (action === 'delLog') {
      const logId = body.logId || (e && e.parameter ? e.parameter.logId : '');
      if (logId) deleteLogById(logSheet, logId);
      return jsonResponse({ success: true });
    }

    if (action === 'undo') {
      const logId = body.logId;
      const lastRow = logSheet.getLastRow();
      let done = false;
      if (lastRow > 1) {
        const ids = logSheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0]));
        const idx = ids.indexOf(String(logId));
        if (idx !== -1) {
          done = undoRow(logSheet, sheet, idx + 2);
        }
      }
      return jsonResponse({ success: done, message: done ? 'Undone' : 'Already undone or not found' });
    }

    if (action === 'bulkPut') {
      const list = body.items || [];
      list.forEach(item => saveOrUpdate(sheet, item));
      return jsonResponse({ success: true, count: list.length });
    }

    return jsonResponse({ success: false, error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function saveOrUpdate(sheet, item) {
  if (!item || !item.id) return;
  const lastRow = sheet.getLastRow();
  const newRowValues = objectToRow(item);

  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0]));
    const index = ids.indexOf(String(item.id));
    if (index !== -1) {
      // تعديل سطر موجود
      sheet.getRange(index + 2, 1, 1, COLS.length).setValues([newRowValues]);
      return;
    }
  }
  // إضافة سطر جديد
  sheet.appendRow(newRowValues);
}

function deleteById(sheet, id) {
  if (!id) return;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0]));
  const index = ids.indexOf(String(id));
  if (index !== -1) {
    sheet.deleteRow(index + 2);
  }
}

function deleteLogById(logSheet, logId) {
  if (!logSheet || !logId) return;
  const lastRow = logSheet.getLastRow();
  if (lastRow <= 1) return;
  const ids = logSheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0]));
  const index = ids.indexOf(String(logId));
  if (index !== -1) {
    logSheet.deleteRow(index + 2);
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
