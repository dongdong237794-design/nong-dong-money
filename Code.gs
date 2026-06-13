/**
 * ระบบน้องด็องทวงตังค์ - Google Apps Script Backend (code.gs)
 * ทำหน้าที่เป็นเว็บแอปและจัดการฐานข้อมูลผ่าน Google Sheets โดยตรง
 */

// ============================================================
// ⚙️ ตั้งค่า: ใส่ Google Drive Folder ID ที่ต้องการเก็บไฟล์
// วิธีหา ID: เปิดโฟลเดอร์ใน Drive → ดู URL → ส่วนหลัง /folders/
// ============================================================
var DRIVE_FOLDER_ID = '1Vkt4G3p7en2A3g4D4L56qe8JY3lKF-6C';

// สำหรับเปิดหน้าเว็บแสดงผล
function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('น้องด็องทวงตังค์')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// สร้างเมนูติดตั้งใน Google Sheets เมื่อเปิดชีตขึ้นมา
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ ระบบน้องด็องทวงตังค์')
    .addItem('🚀 กดติดตั้งระบบ (สร้างตารางเริ่มต้น)', 'setupDatabase')
    .addToUi();
}

/**
 * 1. ฟังก์ชันติดตั้งระบบ (สร้างชีตเริ่มต้น)
 * จะลบชีตเดิมหากชื่อซ้ำ (เฉพาะชีตระบบเริ่มต้น) และสร้างโครงสร้างตารางข้อมูลที่สมบูรณ์ขึ้นมาใหม่
 */
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = {
    'Members': [['id', 'firstName', 'lastName', 'nickname', 'number', 'username', 'password', 'role']],
    'Folders': [['id', 'name', 'color', 'isDeleted', 'createdAt', 'deletedAt']],
    'MoneyBags': [['id', 'folderId', 'name', 'targetAmount', 'color', 'qrCode', 'isDeleted', 'createdAt', 'deletedAt']],
    'Payments': [['id', 'moneyBagId', 'memberId', 'amount', 'method', 'status', 'evidenceDriveUrl', 'slipDate']],
    'Expenses': [['id', 'name', 'category', 'amount', 'date', 'evidenceDriveUrl']],
    'Income': [['id', 'folderId', 'name', 'category', 'amount', 'date', 'evidenceDriveUrl']],
    'Announcements': [['id', 'title', 'message', 'fileId', 'createdAt']]
  };
  
  for (var name in sheets) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    } else {
      sheet.clear();
    }
    sheet.getRange(1, 1, 1, sheets[name][0].length).setValues(sheets[name])
         .setFontWeight('bold')
         .setBackground('#FF6B57')
         .setFontColor('#FFFFFF');
  }
  
  // เพิ่มข้อมูลเริ่มต้นสำหรับสมาชิกและแอดมิน เพื่อความสะดวกในการเข้าระบบครั้งแรก
  var memberSheet = ss.getSheetByName('Members');
  
  // สร้างสมาชิกตัวอย่างเลขที่ 01-37
  var initialMembers = [];
  for (var i = 1; i <= 37; i++) {
    var n = String(i).padStart(2, '0');
    initialMembers.push(['m' + n, 'สมาชิก', 'คนที่ ' + i, 'nick' + i, n, 'user' + n, n, 'user']);
  }
  // แอดมินหลักสำหรับเริ่มระบบ
  initialMembers.push(['admin', 'แอดมิน', 'ห้อง', 'แอด', '00', 'admin', 'admin01', 'admin']);
  
  memberSheet.getRange(2, 1, initialMembers.length, initialMembers[0].length).setValues(initialMembers);

  // สร้างโฟลเดอร์ "รายรับ(ส่วนห้อง)" อัตโนมัติ (ถ้ายังไม่มี)
  var folderSheet = ss.getSheetByName('Folders');
  var folderValues = folderSheet.getDataRange().getValues();
  var hasIncomeFolder = false;
  for (var fi = 1; fi < folderValues.length; fi++) {
    if (String(folderValues[fi][0]) === 'income-room') {
      hasIncomeFolder = true;
      break;
    }
  }
  if (!hasIncomeFolder) {
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    folderSheet.appendRow(['income-room', 'รายรับ(ส่วนห้อง)', '#4CAF50', false, today, '']);
  }

  SpreadsheetApp.getUi().alert('🎉 ติดตั้งระบบน้องด็องทวงตังค์ และสร้างตารางข้อมูลตัวอย่างเรียบร้อยแล้ว!');
}

/**
 * Helper: แปลงชีตเป็น JSON Array ของอ็อบเจกต์ตามหัวตาราง (Row 1 เป็นคีย์)
 */
function getSheetData(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  
  var headers = values[0];
  var list = [];
  for (var r = 1; r < values.length; r++) {
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      var val = values[r][c];
      if (val instanceof Date) {
        obj[headers[c]] = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        obj[headers[c]] = val;
      }
    }
    list.push(obj);
  }
  return list;
}

/**
 * 2. API: ดึงข้อมูลตั้งต้นทั้งหมดสำหรับ Client Side
 */
function getAppData() {
  try {
    return {
      members: getSheetData('Members'),
      folders: getSheetData('Folders'),
      bags: getSheetData('MoneyBags'),
      payments: getSheetData('Payments'),
      expenses: getSheetData('Expenses'),
      income: getSheetData('Income'),
      announcements: getSheetData('Announcements')
    };
  } catch (err) {
    throw new Error('ไม่สามารถดึงข้อมูลแอปพลิเคชันได้: ' + err.toString());
  }
}

/**
 * 3. ฟังก์ชันเกี่ยวกับการเขียนข้อมูลลง Google Sheets
 */
function appendRowToSheet(sheetName, headers, rowData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var values = [];
  for (var i = 0; i < headers.length; i++) {
    values.push(rowData[headers[i]] !== undefined ? rowData[headers[i]] : "");
  }
  sheet.appendRow(values);
}

function updateRowInSheet(sheetName, id, updateObj) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var idColIdx = headers.indexOf('id');
  if (idColIdx === -1) return;
  
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idColIdx]) === String(id)) {
      for (var key in updateObj) {
        var cIdx = headers.indexOf(key);
        if (cIdx !== -1) {
          sheet.getRange(r + 1, cIdx + 1).setValue(updateObj[key]);
        }
      }
      break;
    }
  }
}

// ============================================================
// 📁 ระบบการจัดการโฟลเดอร์เก็บไฟล์อัตโนมัติบน Google Drive
// ============================================================

// ค้นหาหรือสร้างโฟลเดอร์ย่อยภายใต้โฟลเดอร์แม่
function getOrCreateSubFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(folderName);
}

// ค้นหาหรือสร้างโฟลเดอร์หลัก "รายรับ"
function getIncomeFolder() {
  var root = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  return getOrCreateSubFolder(root, 'รายรับ');
}

// ค้นหาหรือสร้างโฟลเดอร์เฉพาะสำหรับ "ถุงเงิน" แต่ละรายการ
function getOrCreateBagFolder(bagId, bagName) {
  var incomeFolder = getIncomeFolder();
  var targetFolderName = 'ถุงเงิน_' + bagId + '_' + bagName;
  
  var folders = incomeFolder.getFolders();
  while (folders.hasNext()) {
    var f = folders.next();
    // ค้นหาโดยเช็คว่ามี ID ของถุงเงินอยู่ในชื่อโฟลเดอร์หรือไม่ ป้องกันปัญหาเปลี่ยนชื่อภายหลัง
    if (f.getName().indexOf(bagId) !== -1) {
      // อัปเดตชื่อโฟลเดอร์ให้เป็นปัจจุบันเสมอหากมีการแก้ไขชื่อถุงเงิน
      if (f.getName() !== targetFolderName) {
        f.setName(targetFolderName);
      }
      return f;
    }
  }
  return incomeFolder.createFolder(targetFolderName);
}

// ค้นหาหรือสร้างโฟลเดอร์ย่อยสำหรับเก็บ "สลิปชำระเงิน" ภายในโฟลเดอร์ถุงเงินนั้น ๆ
function getOrCreateSlipFolder(bagFolder) {
  return getOrCreateSubFolder(bagFolder, 'สลิปชำระเงิน');
}

// ============================================================
// ✅ Helper: อัปโหลดรูปภาพ base64 → Google Drive (กำหนดปลายทางได้)
// คืนค่าเป็น URL สำหรับดูภาพโดยตรงผ่านเว็บ
// ============================================================
function uploadBase64ToDrive(base64DataUrl, fileName, targetFolder) {
  try {
    var matches = base64DataUrl.match(/^data:([^;]+);base64,(.+)$/);
    if (!matches) return '';

    var mimeType = matches[1];           // เช่น "image/jpeg"
    var base64Data = matches[2];         // ข้อมูล base64 ล้วนๆ

    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, mimeType, fileName);

    // หากระบุโฟลเดอร์ปลายทางให้บันทึกที่นั่น ถ้าไม่มีให้บันทึกที่โฟลเดอร์รูท
    var folder = targetFolder || DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var file = folder.createFile(blob);

    // ตั้งสิทธิ์ให้ทุกคนที่มีลิงก์ดูได้ (ไม่ต้อง login)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // คืน URL แบบดูตรงๆ (ไม่ใช่ลิงก์ download)
    return 'https://lh3.googleusercontent.com/d/' + file.getId();

  } catch (err) {
    Logger.log('uploadBase64ToDrive error: ' + err.toString());
    return '';
  }
}

/**
 * API: เพิ่มหรือแก้ไขโฟลเดอร์
 */
function apiSaveFolder(folderData) {
  var sheetName = 'Folders';
  var headers = ['id', 'name', 'color', 'isDeleted', 'createdAt', 'deletedAt'];
  
  if (folderData.isNew) {
    folderData.id = 'f' + new Date().getTime();
    folderData.isDeleted = false;
    folderData.createdAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    appendRowToSheet(sheetName, headers, folderData);
  } else {
    updateRowInSheet(sheetName, folderData.id, { name: folderData.name, color: folderData.color });
  }
  return getAppData();
}

/**
 * API: ลบแบบ Soft-delete หรือกู้คืนโฟลเดอร์/ถุงเงิน
 */
function apiSetDeleteState(type, id, isDeleted) {
  var sheetName = type === 'folder' ? 'Folders' : 'MoneyBags';
  var dateStr = isDeleted ? Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
  updateRowInSheet(sheetName, id, { isDeleted: isDeleted, deletedAt: dateStr });
  return getAppData();
}

/**
 * API: ลบถาวรโฟลเดอร์/ถุงเงินจากระบบชีต
 */
function apiPermanentDelete(type, id) {
  var sheetName = type === 'folder' ? 'Folders' : 'MoneyBags';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var values = sheet.getDataRange().getValues();
  var idColIdx = values[0].indexOf('id');
  
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idColIdx]) === String(id)) {
      sheet.deleteRow(r + 1);
      break;
    }
  }
  return getAppData();
}

/**
 * API: เพิ่มหรือแก้ไขถุงเงิน
 * ✅ [อัปเดตระบบจัดระเบียบ] สร้างโฟลเดอร์ถุงเงินเฉพาะตัวใน "รายรับ" และเซฟรูป QR Code ลงไป
 */
function apiSaveMoneyBag(bagData) {
  var bagSheetName = 'MoneyBags';
  var bagHeaders = ['id', 'folderId', 'name', 'targetAmount', 'color', 'qrCode', 'isDeleted', 'createdAt', 'deletedAt'];

  var bagId = bagData.isNew ? 'b' + new Date().getTime() : bagData.id;
  // 📁 ค้นหาหรือสร้างโฟลเดอร์เฉพาะสำหรับเก็บถุงเงินใบนี้ไว้ในโฟลเดอร์ "รายรับ"
  var bagFolder = getOrCreateBagFolder(bagId, bagData.name);

  // แปลง QR Code base64 → อัปโหลดเข้าไปยังโฟลเดอร์ของถุงเงินใบนี้โดยเฉพาะ
  if (bagData.qrCode && bagData.qrCode.startsWith('data:')) {
    bagData.qrCode = uploadBase64ToDrive(
      bagData.qrCode,
      'qr_' + bagId + '.png',
      bagFolder
    );
  }
  
  if (bagData.isNew) {
    bagData.id = bagId;
    bagData.isDeleted = false;
    bagData.createdAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    appendRowToSheet(bagSheetName, bagHeaders, bagData);
    
    // ดึงรายชื่อสมาชิกเพื่อนำมาผูกสิทธิ์ชำระเงินค้างจ่าย (Payments) ทันที
    var memberSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members');
    var memberValues = memberSheet.getDataRange().getValues();
    var mHeaders = memberValues[0];
    var mIdIdx = mHeaders.indexOf('id');
    var mRoleIdx = mHeaders.indexOf('role');
    
    var payHeaders = ['id', 'moneyBagId', 'memberId', 'amount', 'method', 'status', 'evidenceDriveUrl', 'slipDate'];
    
    for (var r = 1; r < memberValues.length; r++) {
      if (memberValues[r][mRoleIdx] === 'user') {
        var mId = memberValues[r][mIdIdx];
        var payRow = {
          id: 'p' + new Date().getTime() + mId,
          moneyBagId: bagData.id,
          memberId: mId,
          amount: bagData.targetAmount,
          method: '',
          status: 'unpaid',
          evidenceDriveUrl: '',
          slipDate: ''
        };
        appendRowToSheet('Payments', payHeaders, payRow);
      }
    }
  } else {
    updateRowInSheet(bagSheetName, bagData.id, {
      name: bagData.name,
      targetAmount: bagData.targetAmount,
      color: bagData.color
    });
    if (bagData.qrCode) {
      updateRowInSheet(bagSheetName, bagData.id, { qrCode: bagData.qrCode });
    }
  }
  return getAppData();
}

/**
 * API: สำหรับบันทึกหรืออัปเดตสเตตัสการชำระเงินของสมาชิก (Payments)
 * ✅ [อัปเดตระบบจัดระเบียบ] ดึงโฟลเดอร์ของถุงเงินนั้น และเซฟรูปภาพสลิปลงในโฟลเดอร์ "สลิปชำระเงิน"
 */
function apiSavePayment(paymentData) {
  var sheetName = 'Payments';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var mbIdx = headers.indexOf('moneyBagId');
  var mIdIdx = headers.indexOf('memberId');
  
  // แปลงสลิป base64 → เก็บไว้ในโฟลเดอร์ย่อย "สลิปชำระเงิน" ที่อยู่ภายในโฟลเดอร์ถุงเงินนั้น ๆ
  if (paymentData.evidenceDriveUrl && paymentData.evidenceDriveUrl.startsWith('data:')) {
    // 📁 ดึงชื่อถุงเงินปัจจุบัน เพื่อเอามาใช้ประกอบการค้นหาโฟลเดอร์ถุงเงิน
    var bagName = "ไม่ระบุชื่อ";
    var bagSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MoneyBags');
    var bagValues = bagSheet.getDataRange().getValues();
    for (var i = 1; i < bagValues.length; i++) {
      if (String(bagValues[i][0]) === String(paymentData.moneyBagId)) {
        bagName = bagValues[i][2];
        break;
      }
    }
    
    // ค้นหาโฟลเดอร์ถุงเงิน -> เข้าไปยังโฟลเดอร์ย่อย "สลิปชำระเงิน" -> เซฟไฟล์สลิปลงไป
    var bagFolder = getOrCreateBagFolder(paymentData.moneyBagId, bagName);
    var slipFolder = getOrCreateSlipFolder(bagFolder);
    
    paymentData.evidenceDriveUrl = uploadBase64ToDrive(
      paymentData.evidenceDriveUrl,
      'slip_' + paymentData.memberId + '_' + new Date().getTime() + '.jpg',
      slipFolder
    );
  }
  
  var found = false;
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][mbIdx]) === String(paymentData.moneyBagId) && String(values[r][mIdIdx]) === String(paymentData.memberId)) {
      found = true;
      var updateObj = {
        method: paymentData.method,
        status: paymentData.status,
        evidenceDriveUrl: paymentData.evidenceDriveUrl || "",
        slipDate: paymentData.slipDate || ""
      };
      for (var key in updateObj) {
        var cIdx = headers.indexOf(key);
        if (cIdx !== -1) {
          sheet.getRange(r + 1, cIdx + 1).setValue(updateObj[key]);
        }
      }
      break;
    }
  }
  
  // กรณีสมาชิกไม่มีรายการ payment ตั้งต้นในถุงเงินนั้น ให้ทำการ append ใหม่ทันที ป้องกัน Error
  if (!found) {
    paymentData.id = 'p' + new Date().getTime();
    var payHeaders = ['id', 'moneyBagId', 'memberId', 'amount', 'method', 'status', 'evidenceDriveUrl', 'slipDate'];
    appendRowToSheet(sheetName, payHeaders, paymentData);
  }
  
  return getAppData();
}

/**
 * API: เพิ่มรายจ่าย (Expenses)
 */
function apiSubmitExpense(expenseData) {
  if (expenseData.evidenceDriveUrl && expenseData.evidenceDriveUrl.startsWith('data:')) {
    expenseData.evidenceDriveUrl = uploadBase64ToDrive(
      expenseData.evidenceDriveUrl,
      'expense_' + new Date().getTime() + '.jpg'
    );
  }

  expenseData.id = 'e' + new Date().getTime();
  var headers = ['id', 'name', 'category', 'amount', 'date', 'evidenceDriveUrl'];
  appendRowToSheet('Expenses', headers, expenseData);
  return getAppData();
}

/**
 * API: ลบรายจ่าย (Expenses)
 */
function apiDeleteExpense(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Expenses');
  var values = sheet.getDataRange().getValues();
  var idColIdx = values[0].indexOf('id');
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idColIdx]) === String(id)) {
      sheet.deleteRow(r + 1);
      break;
    }
  }
  return getAppData();
}

/**
 * API: เพิ่มรายรับส่วนอื่นๆ ในห้อง (Income)
 */
function apiSubmitIncome(incomeData) {
  if (incomeData.evidenceDriveUrl && incomeData.evidenceDriveUrl.startsWith('data:')) {
    incomeData.evidenceDriveUrl = uploadBase64ToDrive(
      incomeData.evidenceDriveUrl,
      'income_' + new Date().getTime() + '.jpg'
    );
  }

  incomeData.folderId = 'income-room';
  incomeData.id = 'i' + new Date().getTime();
  var headers = ['id', 'folderId', 'name', 'category', 'amount', 'date', 'evidenceDriveUrl'];
  appendRowToSheet('Income', headers, incomeData);
  return getAppData();
}

/**
 * API: ลบรายรับ(ส่วนห้อง)
 */
function apiDeleteIncome(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Income');
  var values = sheet.getDataRange().getValues();
  var idColIdx = values[0].indexOf('id');
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idColIdx]) === String(id)) {
      sheet.deleteRow(r + 1);
      break;
    }
  }
  return getAppData();
}

/**
 * API: สร้างประกาศข่าวสารใหม่ (Announcements)
 */
function apiSubmitAnnouncement(announceData) {
  announceData.id = 'a' + new Date().getTime();
  announceData.createdAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var headers = ['id', 'title', 'message', 'fileId', 'createdAt'];
  appendRowToSheet('Announcements', headers, announceData);
  return getAppData();
}

/**
 * API: บันทึกข้อมูลสมาชิกใหม่โดยแอดมิน (Members)
 */
function apiSaveMember(memberData) {
  memberData.id = 'm' + new Date().getTime();
  memberData.role = 'user';
  var headers = ['id', 'firstName', 'lastName', 'nickname', 'number', 'username', 'password', 'role'];
  appendRowToSheet('Members', headers, memberData);
  return getAppData();
}

/**
 * API: แก้ไขชื่อเล่นโปรไฟล์ส่วนตัว (Profile)
 */
function apiSaveProfile(userId, nickname) {
  updateRowInSheet('Members', userId, { nickname: nickname });
  return getAppData();
}

/**
 * API: แก้ไขรหัสผ่านการเข้าสู่ระบบส่วนตัว (Password)
 */
function apiSavePassword(userId, newPassword) {
  updateRowInSheet('Members', userId, { password: newPassword });
  return getAppData();
}
