/**
 * كود ربط سيستم مدينتي بجوجل شيت (Google Apps Script)
 * -----------------------------------------------------
 * انسخ هذا الكود بالكامل وضعه في:
 * Google Sheets -> Extensions (التوسيعات) -> Apps Script
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

function initHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    const titles = COLS.map(c => c.title);
    sheet.appendRow(titles);
    sheet.getRange(1, 1, 1, titles.length).setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
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
  return COLS.map(c => obj[c.key] != null ? obj[c.key] : '');
}

function doGet(e) {
  try {
    const sheet = getSheet();
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
    let body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else if (e && e.parameter && e.parameter.data) {
      body = JSON.parse(e.parameter.data);
    }
    
    const action = body.action || (e && e.parameter ? e.parameter.action : 'put');

    if (action === 'put') {
      saveOrUpdate(sheet, body.item);
      return jsonResponse({ success: true });
    }

    if (action === 'del') {
      deleteById(sheet, body.id || (e && e.parameter ? e.parameter.id : ''));
      return jsonResponse({ success: true });
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

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
