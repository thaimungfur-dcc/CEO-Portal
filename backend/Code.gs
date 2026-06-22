/**
 * ==============================================================================
 * โค้ด Google Apps Script (Backend) สำหรับโปรเจกต์ CEO PORTAL
 * สถาปัตยกรรม: React (Frontend) -> API (GAS) -> Google Sheets
 * 
 * 📌 ขั้นตอนการใช้งานออนไลน์ (Deployment):
 * 1. เปิด Google Sheets ที่ต้องการใช้เป็นฐานข้อมูลหลัก
 * 2. ไปที่เมนู "ส่วนขยาย" (Extensions) -> "Apps Script"
 * 3. นำโค้ดทั้งหมดนี้ไปวางทับในไฟล์ `Code.gs`
 * 4. ไปที่เมนู "เรียกใช้" (Run) -> เลือกฟังก์ชัน `setupDatabase` (ทำครั้งแรกสุด เพื่อสร้างตารางคอลัมน์เปล่า)
 * 5. กดปุ่ม "การทำให้ใช้งานได้" (Deploy) ที่มุมขวาบน -> "การทำให้ใช้งานได้รายการใหม่" (New deployment)
 * 6. เลือกประเภทเป็น "เว็บแอป" (Web App) โดยกำหนด:
 *    - เรียกใช้ในฐานะ: "ฉัน" (Me)
 *    - ผู้ที่มีสิทธิ์เข้าถึง: "ทุกคน" (Anyone) -> สำคัญที่สุดเพื่อรองรับการยิง API
 * ==============================================================================
 */

const GLOBAL_SHEETS_CONFIG = {
    // ผู้ใช้งานระบบ (ADMINISTRATION)
    'Users': ['id', 'employeeId', 'name', 'role', 'permissions', 'email', 'avatar', 'isDev', 'status', 'createdAt', 'updatedAt'],
    
    // CEO PORTAL REVENUE & OPERATIONS
    'SaleRevenue': ['id', 'mm/dd/yyyy', 'ประเภท', 'ชื่อสินค้า', 'ราคาทุน', 'ราคาขาย', 'ยอดขาย (ชิ้น)', 'มูลค่าขาย(บาท)', 'createdAt', 'updatedAt'],
    'SaleComparison': ['id', 'category', 'name', 'cost', 'price', 'months', 'createdAt', 'updatedAt'],
    'CostExpense': ['id', 'mm/dd/yyyy', 'จำนวนพนักงาน (คน)', 'ค่าจ้างแรงงานรวม', 'ค่าน้ำประปา', 'ค่าไฟฟ้า', 'ค่าแก๊ส/น้ำมัน', 'ต้นทุนและค่าใช้จ่ายรวม', 'createdAt', 'updatedAt'],
    
    // CALENDAR & COMPLIANCE
    'Calendar': ['id', 'date', 'title', 'time', 'type', 'priority', 'status', 'color', 'createdAt', 'updatedAt'],
    'Task_ActionPlans': ['id', 'taskName', 'description', 'assignedTo', 'priority', 'status', 'dueDate', 'relatedLaw', 'progress', 'createdAt', 'updatedAt'],
    
    // AUDITING & SECURITY LOGS
    'AccessLogs': ['id', 'userId', 'action', 'details', 'ipAddress', 'createdAt'],
    'SystemLogs': ['id', 'userId', 'action', 'details', 'ipAddress', 'createdAt'],
    'SystemConfig': ['id', 'category', 'key', 'value', 'description', 'updatedAt']
};

/**
 * ฟังก์ชันสำหรับตั้งค่า Sheet และหัวตารางเบื้องต้น (เรียกใช้หน้า Editor แค่ครั้งเดียว)
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  for (let name in GLOBAL_SHEETS_CONFIG) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    const headers = GLOBAL_SHEETS_CONFIG[name];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold")
      .setBackground("#1d273f")
      .setFontColor("white");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }

  // สร้าง User ตัวอย่างสำหรับ Admin (ถ้ายังไม่มี)
  const userSheet = ss.getSheetByName("Users");
  if (userSheet && userSheet.getLastRow() === 1) {
    userSheet.appendRow([
      '1', 'ADMIN001', 'CEO PORTAL Admin', 'CEO', 
      JSON.stringify({canCreate:true, canEdit:true, canApprove:true, canVerify:true}), 
      'ceo.portal@company.com', '', 'true', 'Active', new Date().toISOString(), new Date().toISOString()
    ]);
  }
  Logger.log("Database Setup Complete! ฐานข้อมูล CEO PORTAL พร้อมใช้งานแล้ว");
}

/**
 * Handle OPTIONS request สำหรับ CORS Preflight
 */
function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Endpoint หลักสำหรับรับ API Request (POST) จาก React Frontend ยิงเข้ามา
 */
function doPost(e) {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  };

  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action; // read, write, update, delete, login
    const sheetName = params.sheet;
    const data = params.data; 

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return createResponse("error", "ไม่สามารถเชื่อมต่อฐานข้อมูล Spreadsheet ได้", null, headers);
    }
    
    let sheet = findSheetCaseInsensitive(ss, sheetName);
    if (!sheet && sheetName) {
      sheet = ss.insertSheet(sheetName);
      let columns = GLOBAL_SHEETS_CONFIG[sheetName] || (data && Array.isArray(data) ? Object.keys(data[0] || {}) : []);
      if (columns.length > 0) {
        sheet.getRange(1, 1, 1, columns.length).setValues([columns])
             .setFontWeight("bold").setBackground("#e8ecef");
      }
    }
    
    if (!sheet && action !== 'login') {
      return createResponse("error", "ไม่พบชีตฐานข้อมูล: " + sheetName, null, headers);
    }

    let result;
    switch (action) {
      case 'read':
        result = readData(sheet, params, headers);
        break;
      case 'lookup':
        result = lookupData(sheet, params, headers); 
        break;
      case 'login':
        result = handleLogin(ss, data, headers);
        break;
      case 'write':
      case 'update':
      case 'delete':
        // Concurrency Guard
        const lock = LockService.getScriptLock();
        if (!lock.tryLock(30000)) {
          return createResponse("error", "เซิร์ฟเวอร์ไม่ว่าง (Lock timeout) กรุณาลองใหม่อีกครั้ง", null, headers);
        }
        try {
          if (action === 'write') result = writeData(sheet, data, headers);
          else if (action === 'update') result = updateData(sheet, data, headers);
          else if (action === 'delete') result = deleteData(sheet, data, headers);
        } finally {
          lock.releaseLock();
        }
        break;
      default:
        result = createResponse("error", "ไม่รู้จักคำสั่ง Action: " + action, null, headers);
    }
    return result;
  } catch (err) {
    return createResponse("error", err.toString(), null, headers);
  }
}

function doGet(e) {
  const headers = { 'Access-Control-Allow-Origin': '*' };
  return createResponse("success", "SMART-HR API is Active. กรุณาใช้ POST สำหรับวิเคราะห์ข้อมูล.", null, headers);
}

// --------------- ACTION HANDLERS ---------------

function readData(sheet, params, headersObj) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return createResponse("success", "ดึงข้อมูลสำเร็จ", { items: [], totalCount: 0 }, headersObj);
  }

  const sheetHeaders = values[0];
  let data = values.slice(1).map(row => {
    const obj = {};
    sheetHeaders.forEach((header, i) => {
      let val = row[i];
      if (typeof val === 'string' && (val.startsWith('{') || val.startsWith('['))) {
        try { val = JSON.parse(val); } catch(e) {}
      }
      obj[header] = val;
    });
    return obj;
  });

  const limit = params.limit || null;
  const offset = params.offset || 0;
  const totalCount = data.length;
  if (limit !== null) data = data.slice(offset, offset + limit);

  return createResponse("success", "ดึงข้อมูลสำเร็จ", { items: data, totalCount, limit, offset }, headersObj);
}

// 📦 Batch Write + Auto cleanup เพื่อความรวดเร็วและสะอาดในการใช้งานย้อนหลัง
function writeData(sheet, data, headersObj) {
  if (!Array.isArray(data)) data = [data];
  if (data.length === 0) return createResponse("success", "ไม่มีข้อมูลให้เขียน", null, headersObj);

  let sheetHeaders = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
  const sheetName = sheet.getName();

  if (sheet.getLastColumn() === 0 || (sheetHeaders.length === 1 && sheetHeaders[0] === "")) {
    let columns = GLOBAL_SHEETS_CONFIG[sheetName] || Object.keys(data[0] || {});
    if (columns.length > 0) {
      sheet.getRange(1, 1, 1, columns.length).setValues([columns])
           .setFontWeight("bold").setBackground("#eaeaec");
      sheetHeaders = columns;
    }
  }
  
  const rows = data.map(item => {
    return sheetHeaders.map(h => {
      var val = item[h];
      if (Array.isArray(val) || (typeof val === 'object' && val !== null)) val = JSON.stringify(val);
      
      if (val != null && val !== "") {
        const hLower = h.toLowerCase();
        if (hLower === 'mm/dd/yyyy' || hLower === 'date' || hLower === 'duedate') {
          const strVal = String(val).trim();
          // Prepend a single quote if the pattern resembles a slash/dash separated date string or month string
          if (strVal.match(/^\d+[\/\-]\d+[\/\-]\d+/) || strVal.match(/^[A-Za-z]{3,}\s+\d+\s+\d+/)) {
            return "'" + strVal;
          }
        }
      }
      return val != null ? val : "";
    });
  });
  
  // เขียนแบบแถวเรียงกลุ่ม (Fast batch write)
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, sheetHeaders.length).setValues(rows);
  
  // Auto-Cleanup: เพิ่มประสิทธิภาพการเข้าค้นด้วยการเคลียร์แถวว่างทิ้ง
  const newLastRow = sheet.getLastRow();
  const maxRows = sheet.getMaxRows();
  if (maxRows > newLastRow + 1) {
    sheet.deleteRows(newLastRow + 1, maxRows - newLastRow - 1);
  }
  
  return createResponse("success", `บันทึกข้อมูลเรียบร้อยจำนวน ${rows.length} รายการ`, null, headersObj);
}

function updateData(sheet, data, headersObj) {
  if (!Array.isArray(data)) data = [data];
  
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  const idIndex = headers.indexOf('id');
  if (idIndex === -1) return createResponse("error", "ไม่พบคอลัมน์ 'id' ในชีต", null, headersObj);

  let updatedCount = 0;
  const updatesMap = {};
  data.forEach(item => { if (item.id) updatesMap[String(item.id)] = item; });

  for (let i = 1; i < values.length; i++) {
    const rowId = String(values[i][idIndex]);
    if (updatesMap[rowId]) {
      const updateItem = updatesMap[rowId];
      headers.forEach((header, colIdx) => {
        if (updateItem.hasOwnProperty(header)) {
          let val = updateItem[header];
          if (Array.isArray(val) || (typeof val === 'object' && val !== null)) val = JSON.stringify(val);
          
          if (val != null && val !== "") {
            const hLower = header.toLowerCase();
            if (hLower === 'mm/dd/yyyy' || hLower === 'date' || hLower === 'duedate') {
              const strVal = String(val).trim();
              if (strVal.match(/^\d+[\/\-]\d+[\/\-]\d+/) || strVal.match(/^[A-Za-z]{3,}\s+\d+\s+\d+/)) {
                val = "'" + strVal;
              }
            }
          }
          values[i][colIdx] = val != null ? val : "";
        }
      });
      updatedCount++;
    }
  }

  if (updatedCount > 0) range.setValues(values);
  return createResponse("success", `อัปเดตข้อมูลจำนวน ${updatedCount} รายการ`, null, headersObj);
}

function deleteData(sheet, data, headersObj) {
  if (!Array.isArray(data)) data = [data];
  const idColIndex = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf('id');
  
  const idsToDelete = new Set(data.map(item => String(item.id)));
  const range = sheet.getDataRange();
  const values = range.getValues();
  const newValues = [values[0]]; // หัวตาราง
  let deletedCount = 0;

  for (let i = 1; i < values.length; i++) {
    if (!idsToDelete.has(String(values[i][idColIndex]))) {
      newValues.push(values[i]);
    } else {
      deletedCount++;
    }
  }

  if (deletedCount > 0) {
    sheet.clearContents();
    sheet.getRange(1, 1, newValues.length, newValues[0].length).setValues(newValues);
  }

  return createResponse("success", `ลบข้อมูลจำนวน ${deletedCount} รายการ`, null, headersObj);
}

function handleLogin(ss, credentials, headersObj) {
  const userSheet = ss.getSheetByName("Users");
  if (!userSheet) {
    return createResponse("error", "ไม่พบชีต 'Users' ใน Google Sheets กรุณารันฟังก์ชัน setupDatabase ใน Apps Script ก่อน", null, headersObj);
  }
  const values = userSheet.getDataRange().getValues();
  const cols = values[0];
  
  const user = values.slice(1).map(r => {
    let u = {}; cols.forEach((c, i) => u[c] = r[i]); return u;
  }).find(u => String(u.employeeId) === String(credentials.employeeId));

  if (user) {
    let perms = {};
    try { perms = typeof user.permissions === 'string' ? JSON.parse(user.permissions) : user.permissions; } catch(e) {}
    
    return createResponse("success", "เข้าสู่ระบบสำเร็จ", {
      id: user.id || user.employeeId,
      employeeId: user.employeeId,
      name: user.name,
      role: user.role,
      permissions: perms,
      isDev: user.isDev === 'true' || user.isDev === true
    }, headersObj);
  } else {
    return createResponse("error", "รหัสพนักงานไม่ถูกต้อง", null, headersObj);
  }
}

function createResponse(status, message, data, headersObj) {
  return ContentService.createTextOutput(JSON.stringify({ status, message, data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function lookupData(sheet, params, headers) { 
  return readData(sheet, params, headers);
}

function findSheetCaseInsensitive(ss, name) {
  const sheets = ss.getSheets();
  const lower = name.toLowerCase();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().toLowerCase() === lower) return sheets[i];
  }
  return null;
}
