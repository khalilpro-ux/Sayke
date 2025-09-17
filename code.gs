// تعريف معرف المصنف الثابت
const SPREADSHEET_ID = '15kTC0uwJWY7p2KKNzyAojQpQ6WKYLjyphUfw-3jBub8';

// تعريف أسماء الأوراق ورؤوسها
const SHEET_CONFIG = {
  'لايكي - داتا': ['الاسم', 'الإيدي', 'تاريخ الانضمام', 'أيام العمل', 'التارغت', 'الساعات', 'آخر تحديث'],
  'سايا - داتا': ['الاسم', 'إيدي سايا', 'تاريخ الانضمام', 'الساعات', 'التارغت', 'آخر تحديث'],
  'معلومات المذيعين - لايكي': ['الاسم', 'الإيدي', 'الاسم الثلاثي', 'المدينة', 'المكتب', 'رقم الهاتف', 'الراتب الأساسي', 'الشحن', 'صافي الراتب', 'تحويل الشحن إلى', 'ملاحظات'],
  'معلومات المذيعين - سايا': ['الاسم', 'الإيدي', 'الاسم الثلاثي', 'المدينة', 'المكتب', 'رقم الهاتف', 'الراتب الأساسي', 'الشحن', 'صافي الراتب', 'تحويل الشحن إلى', 'ملاحظات'],
  'الأرشيف': ['المنصة', 'الاسم', 'الإيدي', 'التارغت', 'الساعات', 'آخر تحديث', 'تاريخ الأرشفة'],
  'الشحن': ['المنصة', 'الاسم', 'الإيدي', 'قيمة الشحن', 'تاريخ الشحن', 'ملاحظات'],
  'داتا راتب - لايكي': ['الإيدي', 'الراتب'],
  'داتا راتب - سايا': ['الإيدي', 'الراتب']
};

// دالة لإنشاء الأوراق تلقائيًا إذا لم تكن موجودة
function createSheetsIfNotExists() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const existingSheets = spreadsheet.getSheets().map(sheet => sheet.getName());

  for (const [sheetName, headers] of Object.entries(SHEET_CONFIG)) {
    if (!existingSheets.includes(sheetName)) {
      console.log(`جاري إنشاء الورقة: ${sheetName}`);
      const newSheet = spreadsheet.insertSheet(sheetName);
      // إضافة الرؤوس
      if (headers && headers.length > 0) {
        const headerRange = newSheet.getRange(1, 1, 1, headers.length);
        headerRange.setValues([headers]);
      }
    }
  }
  return { success: true, message: "تم التحقق من الأوراق وإنشاؤها إن لزم الأمر." };
}

// دالة لقراءة البيانات من ورقة
function readSheetData(sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, message: `الورقة "${sheetName}" غير موجودة.` };
    }

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (lastRow === 0 || lastColumn === 0) {
      return { success: true, headers: [], rows: [], message: "الورقة فارغة." };
    }

    const range = sheet.getRange(1, 1, lastRow, lastColumn);
    const values = range.getValues();

    const headers = values[0];
    const rows = values.slice(1); // استبعاد صف الرؤوس

    return { success: true, headers: headers, rows: rows };
  } catch (error) {
    console.error("Error in readSheetData:", error);
    return { success: false, message: error.toString() };
  }
}

// دالة لكتابة البيانات إلى ورقة (لإضافة صف جديد أو تحديث)
function writeSheetData(sheetName, data, rowIndex = null) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, message: `الورقة "${sheetName}" غير موجودة.` };
    }

    if (!Array.isArray(data) || data.length === 0) {
      return { success: false, message: "البيانات المقدمة غير صحيحة." };
    }

    let range;
    if (rowIndex !== null && rowIndex > 0) {
      // تحديث صف موجود
      range = sheet.getRange(rowIndex, 1, 1, data.length);
    } else {
      // إضافة صف جديد في نهاية الورقة
      const lastRow = sheet.getLastRow() + 1;
      range = sheet.getRange(lastRow, 1, 1, data.length);
    }
    range.setValues([data]); // setValues يتوقع مصفوفة ثنائية الأبعاد

    return { success: true, message: "تم حفظ البيانات بنجاح." };
  } catch (error) {
    console.error("Error in writeSheetData:", error);
    return { success: false, message: error.toString() };
  }
}

// دالة لحذف صف من ورقة
function deleteSheetRow(sheetName, rowIndex) {
  try {
    if (rowIndex <= 1) { // لا يمكن حذف صف الرؤوس أو صف غير صالح
      return { success: false, message: "لا يمكن حذف هذا الصف." };
    }
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, message: `الورقة "${sheetName}" غير موجودة.` };
    }
    sheet.deleteRow(rowIndex);
    return { success: true, message: "تم حذف الصف بنجاح." };
  } catch (error) {
    console.error("Error in deleteSheetRow:", error);
    return { success: false, message: error.toString() };
  }
}

// دالة لتهيئة التطبيق (تُستدعى عند فتح التطبيق لأول مرة)
function initializeApp() {
  createSheetsIfNotExists();
  return { success: true, message: "تمت التهيئة بنجاح." };
}

// دالة نقطة الدخول لتطبيق الويب (Web App)
function doGet(e) {
  // تهيئة التطبيق وإنشاء الأوراق عند أول تحميل
  initializeApp();
  // عرض الواجهة الأمامية
  return HtmlService.createHtmlOutputFromFile('index');
}

function updateSheetCell(sheetName, row, column, value) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error("الورقة '" + sheetName + "' غير موجودة.");
    }
    sheet.getRange(row, column).setValue(value);
    return { success: true };
  } catch (error) {
    console.error("Error in updateSheetCell:", error);
    return { success: false, message: error.toString() };
  }
}
