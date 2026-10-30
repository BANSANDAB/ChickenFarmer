/**
 * ============================================================================
 * ระบบบริหารจัดการฟาร์มไก่ไข่ (Egg Farm Management System)
 * Backend Google Apps Script (ไฟล์รวมสมบูรณ์ - รวมการทำงานทุก Service)
 * ============================================================================
 */

// ============================================================================
// 1. CONFIGURATION (การตั้งค่าระบบ)
// ============================================================================

const CONFIG = {
  // Google IDs ตามที่กำหนด
  SPREADSHEET_ID: '1AmRxzTRyd4-EerC4V9ISIuM1wSe2qNU_u22KkbTrLoQ',
  DRIVE_FOLDER_ID: '13faPNclR-u_n4fd7wfF5fCqPHUhwczwn',

  // ข้อมูลชีตฐานข้อมูล
  SHEET_NAMES: {
    TRANSACTIONS: 'Transactions',
    SETTINGS: 'Settings'
  },

  // หัวตารางคอลัมน์ของชีต Transactions (12 คอลัมน์ A ถึง L)
  TRANSACTION_COLUMNS: [
    'record_id',             // A: รหัสรายการแบบ UUID
    'transaction_date',      // B: วันที่ทำรายการ รูปแบบ YYYY-MM-DD (ค.ศ. สากล)
    'transaction_type',      // C: ประเภทรายการ
    'transaction_direction', // D: INCOME หรือ EXPENSE
    'detail',                // E: รายละเอียดเพิ่มเติม (จำเป็นเมื่อเป็น ค่าอื่น ๆ)
    'egg_trays',             // F: จำนวนแผงไข่ (จำนวนเต็ม >= 1 เมื่อเป็น รับค่าไข่ไก่)
    'amount',                // G: จำนวนเงิน (บาท) ทศนิยมไม่เกิน 2 ตำแหน่ง
    'note',                  // H: หมายเหตุ
    'attachment_file_id',    // I: File ID ใน Google Drive
    'created_at',            // J: เวลาที่สร้างรายการ (ISO string)
    'updated_at',            // K: เวลาที่แก้ไขรายการล่าสุด (ISO string)
    'status'                 // L: ACTIVE หรือ DELETED (สำหรับ Soft Delete)
  ],

  // หัวตารางคอลัมน์ของชีต Settings
  SETTINGS_COLUMNS: [
    'setting_key',
    'setting_value',
    'description',
    'updated_at'
  ],

  // ประเภทรายการและทิศทางการเงิน
  TRANSACTION_TYPES: {
    EGG_INCOME: 'รับค่าไข่ไก่',
    FEED_EXPENSE: 'ค่าอาหารไก่ไข่',
    STRAW_EXPENSE: 'ค่าฟางก้อน',
    OTHER_EXPENSE: 'ค่าอื่น ๆ'
  },

  DIRECTIONS: {
    INCOME: 'INCOME',
    EXPENSE: 'EXPENSE'
  },

  STATUS: {
    ACTIVE: 'ACTIVE',
    DELETED: 'DELETED'
  },

  // กำหนดค่าสำหรับไฟล์แนบ
  ATTACHMENTS: {
    MAX_SIZE_BYTES: 5 * 1024 * 1024, // 5 MB
    ALLOWED_MIME_TYPES: [
      'image/jpeg',
      'image/jpg',
      'image/png',
      'image/webp'
    ],
    FILE_NAME_PREFIX: 'EXP_'
  },

  // ค่าทั่วไปของระบบ
  SYSTEM: {
    TIMEZONE: 'Asia/Bangkok',
    LOCK_TIMEOUT_MS: 10000,
    DEFAULT_PAGE_SIZE: 10
  }
};

function getTransactionDirection(typeName) {
  if (typeName === CONFIG.TRANSACTION_TYPES.EGG_INCOME) {
    return CONFIG.DIRECTIONS.INCOME;
  }
  return CONFIG.DIRECTIONS.EXPENSE;
}

function isValidTransactionType(typeName) {
  const allowed = [
    CONFIG.TRANSACTION_TYPES.EGG_INCOME,
    CONFIG.TRANSACTION_TYPES.FEED_EXPENSE,
    CONFIG.TRANSACTION_TYPES.STRAW_EXPENSE,
    CONFIG.TRANSACTION_TYPES.OTHER_EXPENSE
  ];
  return allowed.includes(typeName);
}

function sanitizeString(input) {
  if (input === null || input === undefined) return '';
  return String(input).replace(/[<>]/g, '').trim();
}

function formatDateIso(dateVal) {
  if (!dateVal) return '';
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, CONFIG.SYSTEM.TIMEZONE, 'yyyy-MM-dd');
  }
  const str = String(dateVal).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    return str.substring(0, 10);
  }
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, CONFIG.SYSTEM.TIMEZONE, 'yyyy-MM-dd');
  }
  return str;
}


// ============================================================================
// 2. DATABASE MANAGEMENT (การเชื่อมต่อและจัดการ Google Sheets & Drive)
// ============================================================================

function getSpreadsheet() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    if (!ss) throw new Error('ไม่พบ Google Sheets ตาม ID ที่ระบุ');
    return ss;
  } catch (error) {
    throw new Error('ไม่สามารถเข้าถึง Google Sheets ได้: ' + (error.message || 'กรุณาตรวจสอบสิทธิ์'));
  }
}

function getDriveFolder() {
  try {
    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    if (!folder) throw new Error('ไม่พบโฟลเดอร์ Google Drive ตาม ID ที่ระบุ');
    return folder;
  } catch (error) {
    throw new Error('ไม่สามารถเข้าถึง Google Drive Folder ได้: ' + (error.message || 'กรุณาตรวจสอบสิทธิ์'));
  }
}

function getSheet(sheetName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function setupDatabase() {
  try {
    const folder = getDriveFolder();
    const folderName = folder.getName();

    const ss = getSpreadsheet();
    const ssName = ss.getName();

    // 1. ตรวจสอบหรือสร้างชีต Transactions
    let transSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TRANSACTIONS);
    let transCreated = false;
    if (!transSheet) {
      transSheet = ss.insertSheet(CONFIG.SHEET_NAMES.TRANSACTIONS);
      transSheet.appendRow(CONFIG.TRANSACTION_COLUMNS);
      const headerRange = transSheet.getRange(1, 1, 1, CONFIG.TRANSACTION_COLUMNS.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#16845B');
      headerRange.setFontColor('#FFFFFF');
      transSheet.setFrozenRows(1);
      transCreated = true;
    } else if (transSheet.getLastRow() === 0) {
      transSheet.appendRow(CONFIG.TRANSACTION_COLUMNS);
      const headerRange = transSheet.getRange(1, 1, 1, CONFIG.TRANSACTION_COLUMNS.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#16845B');
      headerRange.setFontColor('#FFFFFF');
      transSheet.setFrozenRows(1);
    }

    // 2. ตรวจสอบหรือสร้างชีต Settings
    let settingsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
    let settingsCreated = false;
    if (!settingsSheet) {
      settingsSheet = ss.insertSheet(CONFIG.SHEET_NAMES.SETTINGS);
      settingsSheet.appendRow(CONFIG.SETTINGS_COLUMNS);
      const headerRange = settingsSheet.getRange(1, 1, 1, CONFIG.SETTINGS_COLUMNS.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#25352D');
      headerRange.setFontColor('#FFFFFF');
      settingsSheet.setFrozenRows(1);

      const now = new Date().toISOString();
      const defaultSettings = [
        ['farm_name', 'ฟาร์มไก่ไข่สุขใจ', 'ชื่อฟาร์มไก่ไข่', now],
        ['camera_name', 'Xiaomi Smart Camera C300', 'ชื่อรุ่นกล้องวงจรปิด', now],
        ['camera_connection_mode', 'APP', 'โหมดการเชื่อมต่อกล้อง (APP หรือ GATEWAY)', now],
        ['camera_gateway_url', '', 'HTTPS HLS Stream URL จาก Video Gateway (ถ้ามี)', now],
        ['currency', 'THB', 'สกุลเงินหลัก', now]
      ];
      defaultSettings.forEach(row => settingsSheet.appendRow(row));
      settingsCreated = true;
    }

    return {
      success: true,
      message: 'เชื่อมต่อและตรวจสอบฐานข้อมูลสำเร็จ',
      data: {
        spreadsheetName: ssName,
        folderName: folderName,
        transactionsSheetReady: true,
        settingsSheetReady: true,
        transCreated: transCreated,
        settingsCreated: settingsCreated
      }
    };
  } catch (error) {
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการตรวจสอบฐานข้อมูล: ' + error.message,
      error: error.toString()
    };
  }
}

function getSettings() {
  try {
    const sheet = getSheet(CONFIG.SHEET_NAMES.SETTINGS);
    const data = sheet.getDataRange().getValues();
    const settings = {};

    if (data.length > 1) {
      for (let i = 1; i < data.length; i++) {
        const key = String(data[i][0]).trim();
        const value = data[i][1];
        if (key) {
          settings[key] = value;
        }
      }
    }
    return { success: true, settings: settings };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function saveSettings(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.SYSTEM.LOCK_TIMEOUT_MS);
    const sheet = getSheet(CONFIG.SHEET_NAMES.SETTINGS);
    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();

    const existingKeys = {};
    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0]).trim();
      existingKeys[key] = i + 1;
    }

    for (const [key, val] of Object.entries(payload)) {
      if (existingKeys[key]) {
        const rowIndex = existingKeys[key];
        sheet.getRange(rowIndex, 2).setValue(val);
        sheet.getRange(rowIndex, 4).setValue(now);
      } else {
        sheet.appendRow([key, val, '', now]);
      }
    }

    return { success: true, message: 'บันทึกการตั้งค่าสำเร็จ' };
  } catch (error) {
    return { success: false, message: 'บันทึกการตั้งค่าล้มเหลว: ' + error.message };
  } finally {
    lock.releaseLock();
  }
}

function mapRowToTransaction(row) {
  return {
    record_id: String(row[0] || ''),
    transaction_date: row[1] ? formatDateIso(row[1]) : '',
    transaction_type: String(row[2] || ''),
    transaction_direction: String(row[3] || ''),
    detail: String(row[4] || ''),
    egg_trays: row[5] ? Number(row[5]) : 0,
    amount: row[6] ? Number(row[6]) : 0,
    note: String(row[7] || ''),
    attachment_file_id: String(row[8] || ''),
    created_at: String(row[9] || ''),
    updated_at: String(row[10] || ''),
    status: String(row[11] || CONFIG.STATUS.ACTIVE)
  };
}


// ============================================================================
// 3. FILE SERVICE (อัปโหลดและดึงรูปภาพแนบจาก Google Drive)
// ============================================================================

function uploadAttachment(fileData) {
  try {
    if (!fileData || !fileData.base64) {
      return { success: false, message: 'ไม่พบข้อมูลไฟล์ที่ต้องการอัปโหลด' };
    }

    const mimeType = (fileData.mimeType || 'image/jpeg').toLowerCase();
    const isAllowedMime = CONFIG.ATTACHMENTS.ALLOWED_MIME_TYPES.some(allowed => 
      mimeType.includes(allowed) || allowed.includes(mimeType)
    );
    if (!isAllowedMime) {
      return {
        success: false,
        message: 'รองรับเฉพาะไฟล์รูปภาพประเภท JPG, JPEG, PNG และ WEBP เท่านั้น'
      };
    }

    const cleanBase64 = fileData.base64.replace(/^data:image\/[a-z]+;base64,/, '');
    const decodedBytes = Utilities.base64Decode(cleanBase64);

    if (decodedBytes.length > CONFIG.ATTACHMENTS.MAX_SIZE_BYTES) {
      return { success: false, message: 'ขนาดไฟล์เกินกำหนด (สูงสุด 5 MB)' };
    }

    let extension = 'jpg';
    if (mimeType.includes('png')) extension = 'png';
    else if (mimeType.includes('webp')) extension = 'webp';
    else if (mimeType.includes('jpeg') || mimeType.includes('jpg')) extension = 'jpg';

    const dateStr = Utilities.formatDate(new Date(), CONFIG.SYSTEM.TIMEZONE, 'yyyyMMdd');
    const uniqueId = Utilities.getUuid().substring(0, 8);
    const newFileName = `${CONFIG.ATTACHMENTS.FILE_NAME_PREFIX}${dateStr}_${uniqueId}.${extension}`;

    const folder = getDriveFolder();
    const blob = Utilities.newBlob(decodedBytes, mimeType, newFileName);
    const driveFile = folder.createFile(blob);

    return {
      success: true,
      fileId: driveFile.getId(),
      fileName: driveFile.getName(),
      message: 'อัปโหลดไฟล์สำเร็จ'
    };
  } catch (error) {
    return {
      success: false,
      message: 'การอัปโหลดไฟล์ล้มเหลว: ' + error.message
    };
  }
}

function getAttachment(fileId) {
  try {
    if (!fileId) return { success: false, message: 'ไม่พบรหัสไฟล์' };

    const file = DriveApp.getFileById(fileId);
    if (!file) return { success: false, message: 'ไม่พบไฟล์ในระบบ' };

    const blob = file.getBlob();
    const mimeType = blob.getContentType();
    const base64 = Utilities.base64Encode(blob.getBytes());
    const dataUrl = `data:${mimeType};base64,${base64}`;

    return {
      success: true,
      fileId: file.getId(),
      fileName: file.getName(),
      mimeType: mimeType,
      dataUrl: dataUrl
    };
  } catch (error) {
    return {
      success: false,
      message: 'ไม่สามารถเปิดไฟล์แนบได้: ' + error.message
    };
  }
}


// ============================================================================
// 4. TRANSACTION SERVICE (CRUD จัดการรายรับ–รายจ่าย)
// ============================================================================

function saveTransaction(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.SYSTEM.LOCK_TIMEOUT_MS);

    if (!payload) return { success: false, message: 'ไม่พบข้อมูลสำหรับบันทึก' };

    const dateStr = formatDateIso(payload.transaction_date);
    if (!dateStr || !/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
      return { success: false, message: 'กรุณาระบุวันที่ให้ถูกต้องในรูปแบบมาตรฐาน (YYYY-MM-DD)' };
    }

    const type = sanitizeString(payload.transaction_type);
    if (!isValidTransactionType(type)) {
      return { success: false, message: 'ประเภทรายการไม่ถูกต้อง' };
    }

    const direction = getTransactionDirection(type);

    let eggTrays = 0;
    if (type === CONFIG.TRANSACTION_TYPES.EGG_INCOME) {
      eggTrays = parseInt(payload.egg_trays, 10);
      if (isNaN(eggTrays) || eggTrays < 1) {
        return { success: false, message: 'จำนวนแผงไข่ต้องเป็นตัวเลขจำนวนเต็มตั้งแต่ 1 แผงขึ้นไป' };
      }
    }

    let detail = '';
    if (type === CONFIG.TRANSACTION_TYPES.OTHER_EXPENSE) {
      detail = sanitizeString(payload.detail);
      if (!detail) {
        return { success: false, message: 'กรุณากรอกรายละเอียดเพิ่มเติมสำหรับรายการค่าอื่น ๆ' };
      }
    } else {
      detail = sanitizeString(payload.detail || '');
    }

    const amount = parseFloat(payload.amount);
    if (isNaN(amount) || amount <= 0) {
      return { success: false, message: 'จำนวนเงินต้องเป็นตัวเลขที่มากกว่า 0 บาท' };
    }
    const cleanAmount = Math.round(amount * 100) / 100;

    const note = sanitizeString(payload.note || '');
    const attachmentFileId = sanitizeString(payload.attachment_file_id || '');

    const recordId = Utilities.getUuid();
    const nowIso = new Date().toISOString();
    const status = CONFIG.STATUS.ACTIVE;

    const newRow = [
      recordId,
      dateStr,
      type,
      direction,
      detail,
      eggTrays,
      cleanAmount,
      note,
      attachmentFileId,
      nowIso,
      nowIso,
      status
    ];

    const sheet = getSheet(CONFIG.SHEET_NAMES.TRANSACTIONS);
    sheet.appendRow(newRow);

    return {
      success: true,
      recordId: recordId,
      message: 'บันทึกรายการสำเร็จเรียบร้อย'
    };
  } catch (error) {
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการบันทึกข้อมูล: ' + error.message
    };
  } finally {
    lock.releaseLock();
  }
}

function updateTransaction(recordId, payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.SYSTEM.LOCK_TIMEOUT_MS);

    if (!recordId) return { success: false, message: 'ไม่พบรหัสรายการที่ต้องการแก้ไข' };

    const sheet = getSheet(CONFIG.SHEET_NAMES.TRANSACTIONS);
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) return { success: false, message: 'ไม่พบข้อมูลในระบบ' };

    let targetRowIndex = -1;
    let existingRow = null;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(recordId).trim()) {
        targetRowIndex = i + 1;
        existingRow = data[i];
        break;
      }
    }

    if (targetRowIndex === -1 || !existingRow) {
      return { success: false, message: 'ไม่พบรายการที่ต้องการแก้ไข หรือรายการอาจถูกลบไปแล้ว' };
    }

    if (String(existingRow[11]).trim() === CONFIG.STATUS.DELETED) {
      return { success: false, message: 'ไม่สามารถแก้ไขรายการที่ถูกลบไปแล้วได้' };
    }

    const dateStr = formatDateIso(payload.transaction_date || existingRow[1]);
    const type = sanitizeString(payload.transaction_type || existingRow[2]);
    if (!isValidTransactionType(type)) {
      return { success: false, message: 'ประเภทรายการไม่ถูกต้อง' };
    }

    const direction = getTransactionDirection(type);

    let eggTrays = 0;
    if (type === CONFIG.TRANSACTION_TYPES.EGG_INCOME) {
      eggTrays = parseInt(payload.egg_trays, 10);
      if (isNaN(eggTrays) || eggTrays < 1) {
        return { success: false, message: 'จำนวนแผงไข่ต้องเป็นตัวเลขจำนวนเต็มตั้งแต่ 1 แผงขึ้นไป' };
      }
    }

    let detail = '';
    if (type === CONFIG.TRANSACTION_TYPES.OTHER_EXPENSE) {
      detail = sanitizeString(payload.detail);
      if (!detail) {
        return { success: false, message: 'กรุณากรอกรายละเอียดเพิ่มเติมสำหรับรายการค่าอื่น ๆ' };
      }
    } else {
      detail = sanitizeString(payload.detail || '');
    }

    const amount = parseFloat(payload.amount);
    if (isNaN(amount) || amount <= 0) {
      return { success: false, message: 'จำนวนเงินต้องมากกว่า 0 บาท' };
    }
    const cleanAmount = Math.round(amount * 100) / 100;

    const note = sanitizeString(payload.note !== undefined ? payload.note : existingRow[7]);

    let attachmentFileId = existingRow[8];
    if (payload.attachment_file_id !== undefined && payload.attachment_file_id !== '') {
      attachmentFileId = sanitizeString(payload.attachment_file_id);
    }

    const nowIso = new Date().toISOString();

    const updatedValues = [
      recordId,
      dateStr,
      type,
      direction,
      detail,
      eggTrays,
      cleanAmount,
      note,
      attachmentFileId,
      existingRow[9],
      nowIso,
      CONFIG.STATUS.ACTIVE
    ];

    sheet.getRange(targetRowIndex, 1, 1, updatedValues.length).setValues([updatedValues]);

    return { success: true, message: 'อัปเดตข้อมูลรายการเรียบร้อยแล้ว' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาดในการอัปเดตรายการ: ' + error.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteTransaction(recordId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.SYSTEM.LOCK_TIMEOUT_MS);

    if (!recordId) return { success: false, message: 'ไม่พบรหัสรายการที่ต้องการลบ' };

    const sheet = getSheet(CONFIG.SHEET_NAMES.TRANSACTIONS);
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) return { success: false, message: 'ไม่พบข้อมูลในระบบ' };

    let targetRowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(recordId).trim()) {
        targetRowIndex = i + 1;
        break;
      }
    }

    if (targetRowIndex === -1) return { success: false, message: 'ไม่พบรายการที่ต้องการลบ' };

    const nowIso = new Date().toISOString();
    sheet.getRange(targetRowIndex, 11).setValue(nowIso);
    sheet.getRange(targetRowIndex, 12).setValue(CONFIG.STATUS.DELETED);

    return { success: true, message: 'ลบรายการสำเร็จ (สถานะเปลี่ยนเป็น DELETED)' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาดในการลบรายการ: ' + error.message };
  } finally {
    lock.releaseLock();
  }
}

function getTransactions(filters) {
  try {
    filters = filters || {};
    const sheet = getSheet(CONFIG.SHEET_NAMES.TRANSACTIONS);
    const data = sheet.getDataRange().getValues();

    const items = [];
    let totalIncome = 0;
    let totalExpense = 0;
    let totalTrays = 0;

    if (data.length > 1) {
      const search = (filters.search || '').trim().toLowerCase();
      const filterType = (filters.type || '').trim();
      const filterMonthYear = (filters.monthYear || '').trim();
      const startDate = (filters.startDate || '').trim();
      const endDate = (filters.endDate || '').trim();

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const status = String(row[11] || CONFIG.STATUS.ACTIVE).trim();

        if (status === CONFIG.STATUS.DELETED) continue;

        const item = mapRowToTransaction(row);

        if (filterType && item.transaction_type !== filterType) continue;
        if (filterMonthYear && !item.transaction_date.startsWith(filterMonthYear)) continue;
        if (startDate && item.transaction_date < startDate) continue;
        if (endDate && item.transaction_date > endDate) continue;

        if (search) {
          const matchTarget = `${item.transaction_date} ${item.transaction_type} ${item.detail} ${item.note} ${item.amount}`.toLowerCase();
          if (!matchTarget.includes(search)) continue;
        }

        if (item.transaction_direction === CONFIG.DIRECTIONS.INCOME) {
          totalIncome += item.amount;
          totalTrays += item.egg_trays;
        } else {
          totalExpense += item.amount;
        }

        items.push(item);
      }
    }

    items.sort((a, b) => {
      if (b.transaction_date !== a.transaction_date) {
        return b.transaction_date.localeCompare(a.transaction_date);
      }
      return b.created_at.localeCompare(a.created_at);
    });

    const totalCount = items.length;
    const page = Math.max(1, parseInt(filters.page, 10) || 1);
    const pageSize = Math.max(1, parseInt(filters.pageSize, 10) || CONFIG.SYSTEM.DEFAULT_PAGE_SIZE);
    const totalPages = Math.ceil(totalCount / pageSize) || 1;

    const startIndex = (page - 1) * pageSize;
    const paginatedItems = items.slice(startIndex, startIndex + pageSize);

    return {
      success: true,
      items: paginatedItems,
      totalCount: totalCount,
      totalPages: totalPages,
      page: page,
      pageSize: pageSize,
      summary: {
        totalIncome: Math.round(totalIncome * 100) / 100,
        totalExpense: Math.round(totalExpense * 100) / 100,
        netBalance: Math.round((totalIncome - totalExpense) * 100) / 100,
        totalTrays: totalTrays
      }
    };
  } catch (error) {
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการโหลดข้อมูลรายการ: ' + error.message,
      items: [],
      totalCount: 0
    };
  }
}


// ============================================================================
// 5. DASHBOARD SERVICE (ประมวลผลสถิติและกราฟ)
// ============================================================================

function getDashboardData(filters) {
  try {
    filters = filters || { period: 'this_month' };
    const sheet = getSheet(CONFIG.SHEET_NAMES.TRANSACTIONS);
    const data = sheet.getDataRange().getValues();

    const dateRange = resolveDashboardDateRange(filters);

    let totalEggIncome = 0;
    let totalExpense = 0;
    let totalEggTrays = 0;
    let totalFeedExpense = 0;
    let totalStrawExpense = 0;
    let totalOtherExpense = 0;

    const monthlyMap = {};
    const activeTransactions = [];

    if (data.length > 1) {
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const status = String(row[11] || CONFIG.STATUS.ACTIVE).trim();
        if (status === CONFIG.STATUS.DELETED) continue;

        const item = mapRowToTransaction(row);
        activeTransactions.push(item);

        if (dateRange.startDate && item.transaction_date < dateRange.startDate) continue;
        if (dateRange.endDate && item.transaction_date > dateRange.endDate) continue;

        const amt = item.amount || 0;
        const type = item.transaction_type;

        if (type === CONFIG.TRANSACTION_TYPES.EGG_INCOME) {
          totalEggIncome += amt;
          totalEggTrays += item.egg_trays || 0;
        } else if (type === CONFIG.TRANSACTION_TYPES.FEED_EXPENSE) {
          totalFeedExpense += amt;
          totalExpense += amt;
        } else if (type === CONFIG.TRANSACTION_TYPES.STRAW_EXPENSE) {
          totalStrawExpense += amt;
          totalExpense += amt;
        } else if (type === CONFIG.TRANSACTION_TYPES.OTHER_EXPENSE) {
          totalOtherExpense += amt;
          totalExpense += amt;
        } else if (item.transaction_direction === CONFIG.DIRECTIONS.EXPENSE) {
          totalExpense += amt;
        }

        if (item.transaction_date && item.transaction_date.length >= 7) {
          const ym = item.transaction_date.substring(0, 7);
          if (!monthlyMap[ym]) {
            monthlyMap[ym] = { income: 0, expense: 0, trays: 0 };
          }
          if (item.transaction_direction === CONFIG.DIRECTIONS.INCOME) {
            monthlyMap[ym].income += amt;
            monthlyMap[ym].trays += (item.egg_trays || 0);
          } else {
            monthlyMap[ym].expense += amt;
          }
        }
      }
    }

    const netProfit = totalEggIncome - totalExpense;
    const sortedMonths = Object.keys(monthlyMap).sort();
    const chartLabels = [];
    const chartIncome = [];
    const chartExpense = [];
    const chartTrays = [];

    const targetMonths = sortedMonths.length > 0 ? sortedMonths : getLastNMonths(6);
    targetMonths.forEach(ym => {
      chartLabels.push(formatYearMonthToThai(ym));
      const mData = monthlyMap[ym] || { income: 0, expense: 0, trays: 0 };
      chartIncome.push(Math.round(mData.income * 100) / 100);
      chartExpense.push(Math.round(mData.expense * 100) / 100);
      chartTrays.push(mData.trays);
    });

    activeTransactions.sort((a, b) => {
      if (b.transaction_date !== a.transaction_date) {
        return b.transaction_date.localeCompare(a.transaction_date);
      }
      return b.created_at.localeCompare(a.created_at);
    });
    const recentTransactions = activeTransactions.slice(0, 10);

    return {
      success: true,
      period: filters.period,
      dateRange: dateRange,
      cards: {
        totalEggIncome: Math.round(totalEggIncome * 100) / 100,
        totalExpense: Math.round(totalExpense * 100) / 100,
        netProfit: Math.round(netProfit * 100) / 100,
        totalEggTrays: totalEggTrays,
        totalFeedExpense: Math.round(totalFeedExpense * 100) / 100,
        totalStrawExpense: Math.round(totalStrawExpense * 100) / 100,
        totalOtherExpense: Math.round(totalOtherExpense * 100) / 100
      },
      charts: {
        monthly: {
          labels: chartLabels,
          income: chartIncome,
          expense: chartExpense,
          trays: chartTrays
        },
        expenseBreakdown: {
          labels: ['ค่าอาหารไก่ไข่', 'ค่าฟางก้อน', 'ค่าอื่น ๆ'],
          data: [
            Math.round(totalFeedExpense * 100) / 100,
            Math.round(totalStrawExpense * 100) / 100,
            Math.round(totalOtherExpense * 100) / 100
          ]
        }
      },
      recentTransactions: recentTransactions
    };
  } catch (error) {
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการโหลดข้อมูล Dashboard: ' + error.message
    };
  }
}

function resolveDashboardDateRange(filters) {
  const now = new Date();
  const tz = CONFIG.SYSTEM.TIMEZONE;
  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const thisMonthStr = Utilities.formatDate(now, tz, 'yyyy-MM');
  const thisYearStr = Utilities.formatDate(now, tz, 'yyyy');

  switch (filters.period) {
    case 'today':
      return { startDate: todayStr, endDate: todayStr, label: 'วันนี้' };
    case 'this_month':
      return { startDate: `${thisMonthStr}-01`, endDate: todayStr, label: 'เดือนนี้' };
    case 'this_year':
      return { startDate: `${thisYearStr}-01-01`, endDate: todayStr, label: 'ปีนี้' };
    case 'custom':
      return { startDate: filters.startDate || '', endDate: filters.endDate || '', label: 'ช่วงวันที่กำหนด' };
    default:
      return { startDate: `${thisMonthStr}-01`, endDate: todayStr, label: 'เดือนนี้' };
  }
}

function getLastNMonths(n) {
  const list = [];
  const now = new Date();
  for (let i = n - 1; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    list.push(Utilities.formatDate(d, CONFIG.SYSTEM.TIMEZONE, 'yyyy-MM'));
  }
  return list;
}

function formatYearMonthToThai(ym) {
  if (!ym || ym.length < 7) return ym;
  const parts = ym.split('-');
  const yearCe = parseInt(parts[0], 10);
  const monthIdx = parseInt(parts[1], 10) - 1;
  const yearBe = yearCe + 543;

  const thaiMonthsShort = [
    'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',
    'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'
  ];

  const mName = thaiMonthsShort[monthIdx] || parts[1];
  return `${mName} ${yearBe}`;
}


// ============================================================================
// 6. REPORT SERVICE (จัดทำข้อมูลสำหรับพิมพ์รายงาน A4)
// ============================================================================

function getReportData(filters) {
  try {
    filters = filters || {};
    const sheet = getSheet(CONFIG.SHEET_NAMES.TRANSACTIONS);
    const data = sheet.getDataRange().getValues();

    const items = [];
    let totalIncome = 0;
    let totalExpense = 0;
    let totalEggTrays = 0;
    let countIncome = 0;
    let countExpense = 0;

    const startDate = (filters.startDate || '').trim();
    const endDate = (filters.endDate || '').trim();
    const targetType = (filters.transactionType || '').trim();

    if (data.length > 1) {
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const status = String(row[11] || CONFIG.STATUS.ACTIVE).trim();
        if (status === CONFIG.STATUS.DELETED) continue;

        const item = mapRowToTransaction(row);

        if (startDate && item.transaction_date < startDate) continue;
        if (endDate && item.transaction_date > endDate) continue;
        if (targetType && targetType !== 'ALL' && item.transaction_type !== targetType) continue;

        if (item.transaction_direction === CONFIG.DIRECTIONS.INCOME) {
          totalIncome += item.amount;
          totalEggTrays += item.egg_trays;
          countIncome++;
        } else {
          totalExpense += item.amount;
          countExpense++;
        }

        items.push(item);
      }
    }

    items.sort((a, b) => {
      if (a.transaction_date !== b.transaction_date) {
        return a.transaction_date.localeCompare(b.transaction_date);
      }
      return a.created_at.localeCompare(b.created_at);
    });

    let farmName = 'ฟาร์มไก่ไข่';
    const settingsRes = getSettings();
    if (settingsRes.success && settingsRes.settings && settingsRes.settings.farm_name) {
      farmName = settingsRes.settings.farm_name;
    }

    const now = new Date();
    const generatedDateThai = formatFullDateTimeThai(now);

    return {
      success: true,
      farmName: farmName,
      title: 'รายงานรายรับ–รายจ่ายฟาร์มไก่ไข่',
      generatedDateThai: generatedDateThai,
      filterConditions: {
        startDate: startDate,
        endDate: endDate,
        startDateThai: startDate ? formatDateToThaiBE(startDate) : 'ตั้งแต่เริ่มต้น',
        endDateThai: endDate ? formatDateToThaiBE(endDate) : 'ถึงปัจจุบัน',
        transactionType: targetType || 'ทั้งหมด'
      },
      summary: {
        totalIncome: Math.round(totalIncome * 100) / 100,
        totalExpense: Math.round(totalExpense * 100) / 100,
        netBalance: Math.round((totalIncome - totalExpense) * 100) / 100,
        totalEggTrays: totalEggTrays,
        totalRecords: items.length,
        countIncome: countIncome,
        countExpense: countExpense
      },
      items: items
    };
  } catch (error) {
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการสร้างรายงาน: ' + error.message,
      items: []
    };
  }
}

function formatDateToThaiBE(dateStr) {
  if (!dateStr || dateStr.length < 10) return dateStr;
  const parts = dateStr.substring(0, 10).split('-');
  const day = parseInt(parts[2], 10);
  const monthIdx = parseInt(parts[1], 10) - 1;
  const yearBe = parseInt(parts[0], 10) + 543;

  const thaiMonths = [
    'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
  ];

  return `${day} ${thaiMonths[monthIdx]} ${yearBe}`;
}

function formatFullDateTimeThai(date) {
  const day = date.getDate();
  const monthIdx = date.getMonth();
  const yearBe = date.getFullYear() + 543;
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');

  const thaiMonths = [
    'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
  ];

  return `${day} ${thaiMonths[monthIdx]} ${yearBe} เวลา ${hours}:${minutes} น.`;
}


// ============================================================================
// 7. CAMERA SERVICE (ระบบกล้องวงจรปิด XIAOMI SMART CAMERA C300)
// ============================================================================

function getCameraConfig() {
  try {
    const settingsRes = getSettings();
    const settings = settingsRes.success ? settingsRes.settings : {};

    return {
      success: true,
      camera: {
        model: settings.camera_name || 'Xiaomi Smart Camera C300',
        connectionMode: settings.camera_connection_mode || 'APP',
        gatewayUrl: settings.camera_gateway_url || '',
        supportsDirectStream: false,
        appDeepLink: 'mihome://',
        appStoreUrlIos: 'https://apps.apple.com/app/mi-home-xiaomi-smarthome/id957323480',
        playStoreUrlAndroid: 'https://play.google.com/store/apps/details?id=com.xiaomi.smarthome'
      }
    };
  } catch (error) {
    return {
      success: false,
      message: 'ไม่สามารถดึงข้อมูลการตั้งค่ากล้องได้: ' + error.message
    };
  }
}

function saveCameraConfig(config) {
  try {
    if (!config) return { success: false, message: 'ไม่พบข้อมูลการตั้งค่ากล้อง' };

    const mode = (config.connectionMode || 'APP').toUpperCase();
    if (mode !== 'APP' && mode !== 'GATEWAY') {
      return { success: false, message: 'โหมดการเชื่อมต่อต้องเป็น APP หรือ GATEWAY เท่านั้น' };
    }

    let cleanGatewayUrl = (config.gatewayUrl || '').trim();
    if (mode === 'GATEWAY' && cleanGatewayUrl) {
      if (!cleanGatewayUrl.startsWith('https://')) {
        return { success: false, message: 'Video Gateway URL ต้องขึ้นต้นด้วย https://' };
      }
      if (cleanGatewayUrl.includes('@')) {
        return { success: false, message: 'ไม่อนุญาตให้ระบุ Username หรือ Password ลงใน URL' };
      }
    }

    const payload = {
      camera_connection_mode: mode,
      camera_gateway_url: cleanGatewayUrl
    };

    return saveSettings(payload);
  } catch (error) {
    return { success: false, message: 'บันทึกการตั้งค่ากล้องล้มเหลว: ' + error.message };
  }
}

function checkStreamHealth(url) {
  try {
    if (!url || !url.startsWith('https://')) {
      return { online: false, message: 'กรุณาระบุ HTTPS URL ของ Video Gateway ที่ถูกต้อง' };
    }

    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      headers: {
        'Accept': 'application/x-mpegURL, application/vnd.apple.mpegurl, */*'
      }
    });

    const statusCode = response.getResponseCode();
    if (statusCode >= 200 && statusCode < 400) {
      return {
        online: true,
        statusCode: statusCode,
        message: 'เชื่อมต่อ Video Gateway สำเร็จ (Status: ' + statusCode + ')'
      };
    } else {
      return {
        online: false,
        statusCode: statusCode,
        message: 'Video Gateway ตอบกลับด้วยสถานะ HTTP ' + statusCode
      };
    }
  } catch (error) {
    return { online: false, message: 'ไม่สามารถติดต่อ Video Gateway ได้: ' + error.message };
  }
}


// ============================================================================
// 8. MAIN WEB APP ROUTER & API ENDPOINTS
// ============================================================================

function doGet(e) {
  try {
    const template = HtmlService.createTemplateFromFile('Index');
    template.appTitle = 'ระบบบริหารจัดการฟาร์มไก่ไข่';
    template.spreadsheetId = CONFIG.SPREADSHEET_ID;

    const htmlOutput = template.evaluate();
    htmlOutput
      .setTitle('ระบบบริหารจัดการฟาร์มไก่ไข่ | Egg Farm Management System')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    return htmlOutput;
  } catch (error) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family: sans-serif; padding: 20px; color: #721c24; background: #f8d7da; border-radius: 8px;">' +
      '<h3>เกิดข้อผิดพลาดในการโหลดระบบ</h3>' +
      '<p>' + error.message + '</p>' +
      '</div>'
    );
  }
}

// API Endpoints ที่เรียกจาก google.script.run
function apiSetupDatabase() { return setupDatabase(); }
function apiGetDashboardData(filters) { return getDashboardData(filters); }
function apiGetTransactions(filters) { return getTransactions(filters); }
function apiSaveTransaction(payload) { return saveTransaction(payload); }
function apiUpdateTransaction(recordId, payload) { return updateTransaction(recordId, payload); }
function apiDeleteTransaction(recordId) { return deleteTransaction(recordId); }
function apiUploadAttachment(fileData) { return uploadAttachment(fileData); }
function apiGetAttachment(fileId) { return getAttachment(fileId); }
function apiGetReportData(filters) { return getReportData(filters); }
function apiGetSettings() { return getSettings(); }
function apiSaveSettings(payload) { return saveSettings(payload); }
function apiGetCameraConfig() { return getCameraConfig(); }
function apiSaveCameraConfig(config) { return saveCameraConfig(config); }
function apiCheckStreamHealth(url) { return checkStreamHealth(url); }
