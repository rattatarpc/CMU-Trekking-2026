/**
 * Google Apps Script Backend for:
 * CMU Trekking First Aid & Command Center
 * ระบบปฐมพยาบาลและ Command Center งานรับน้องขึ้นดอย มหาวิทยาลัยเชียงใหม่
 */

const SPREADSHEET_ID = '1lveyDmShPR7At59VQUob8iFxrKnB4hfOGalTEkVeims';
const SHEET_NAME = 'Cases';
const STATIONS = [
  { id: 'st1', code: 'Station 1 หน้า มช.', name: 'Station 1 หน้า มช.', km: 0.0, leader: '-', phone: '-' },
  { id: 'st2', code: 'Station 2 วังบัวบาน', name: 'Station 2 วังบัวบาน', km: 1.8, leader: '-', phone: '-' },
  { id: 'st3', code: 'Station 3 วัดผาลาด', name: 'Station 3 วัดผาลาด', km: 4.2, leader: '-', phone: '-' },
  { id: 'st4', code: 'Station 4 สถานีไฟป่า', name: 'Station 4 สถานีไฟป่า', km: 5.5, leader: '-', phone: '-' },
  { id: 'st5', code: 'Station 5 ศาลาชมวิว', name: 'Station 5 ศาลาชมวิว', km: 6.8, leader: '-', phone: '-' },
  { id: 'st6', code: 'Station 6 หอดูดาว', name: 'Station 6 หอดูดาว', km: 8.0, leader: '-', phone: '-' },
  { id: 'st7', code: 'Station 7 ลานชมวิว', name: 'Station 7 ลานชมวิว', km: 9.1, leader: '-', phone: '-' },
  { id: 'st8', code: 'Station 8 โค้งขุนกัณฑ์', name: 'Station 8 โค้งขุนกัณฑ์', km: 10.2, leader: '-', phone: '-' },
  { id: 'st9', code: 'Station 9 เชิงพระธาตุ', name: 'Station 9 เชิงพระธาตุ', km: 11.2, leader: '-', phone: '-' },
  { id: 'st10', code: 'Station 10 บนวัดพระธาตุ', name: 'Station 10 บนวัดพระธาตุ', km: 11.8, leader: '-', phone: '-' },
  { id: 'st11', code: 'Station 11 ลานจอดรถ 2', name: 'Station 11 ลานจอดรถ 2', km: 12.0, leader: '-', phone: '-' },
  { id: 'st12', code: 'Station 12 หน้า มช. ขาลง', name: 'Station 12 หน้า มช. ขาลง', km: 12.0, leader: '-', phone: '-' }
];

const HEADERS = [
  'Case_ID',          // A
  'Timestamp',        // B
  'Last_Updated',     // C
  'Staff_Name',       // D
  'Station',          // E
  'Patient_Type',     // F
  'Student_ID',       // G
  'Patient_Name',     // H
  'Faculty',          // I
  'Year_Level',       // J
  'Age',              // K
  'Gender',           // L
  'Underlying_Disease',// M
  'Emergency_Contact',// N
  'Triage_Level',     // O
  'Present_Illness',  // P
  'AVPU',             // Q
  'GCS',              // R
  'BP',               // S
  'PR',               // T
  'RR',               // U
  'SpO2',             // V
  'BT',               // W
  'DTX',              // X
  'PE_Note',          // Y
  'Diagnosis',        // Z
  'Management',       // AA
  'Disposition',      // AB
  'Transfer_Hospital',// AC
  'Transfer_Method',  // AD
  'Notes',            // AE
  'Status',           // AF
  'Case_Type',        // AG
  'Photo_URL',        // AH
  'Minor_Count',      // AI
  'Progression_Note', // AJ
  'Allergy',          // AK
  'High_Risk',        // AL
  'High_Risk_Reason'  // AM
];

/**
 * Web App entry point with iOS Safari optimization
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'getInitialData') {
    const res = getInitialData();
    return ContentService.createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const htmlOutput = HtmlService.createTemplateFromFile('index').evaluate();
  return htmlOutput
    .setTitle('CMU Trekking 2026')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * REST API Endpoint for GitHub Pages / External Frontend
 */
function doPost(e) {
  try {
    const contents = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const body = JSON.parse(contents);
    const action = body.action;

    if (action === 'getInitialData') {
      const res = getInitialData();
      return ContentService.createTextOutput(JSON.stringify(res))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'saveOrUpdateCase') {
      const res = saveOrUpdateCase(body.payload || body.caseData || body);
      return ContentService.createTextOutput(JSON.stringify(res))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'saveAmbulanceChecklist') {
      const res = saveAmbulanceChecklist(body.payload || body.data || body);
      return ContentService.createTextOutput(JSON.stringify(res))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: 'Unknown action: ' + action
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      status: 'error',
      error: error.message,
      message: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Initialize or get the Cases sheet
 */
function getOrCreateSheet() {
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    ss = null;
  }
  if (!ss) {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    const sheets = ss.getSheets();
    if (sheets.length === 1 && sheets[0].getLastRow() <= 1) {
      sheet = sheets[0];
      sheet.setName(SHEET_NAME);
    } else {
      sheet = ss.insertSheet(SHEET_NAME);
    }
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
      headerRange.setBackground('#5b4a8e');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  // Auto-verify and fix missing/blank headers in Row 1 (e.g. AG: Case_Type, AH: Photo_URL, AI: Minor_Count, AL: High_Risk, AM: High_Risk_Reason)
  if (sheet.getLastRow() > 0) {
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    const currentHeaderValues = headerRange.getValues()[0];
    let needsUpdate = false;
    for (let i = 0; i < HEADERS.length; i++) {
      if (!currentHeaderValues[i] || String(currentHeaderValues[i]).trim() !== HEADERS[i]) {
        currentHeaderValues[i] = HEADERS[i];
        needsUpdate = true;
      }
    }
    if (needsUpdate) {
      headerRange.setValues([currentHeaderValues]);
      headerRange.setBackground('#5b4a8e');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

/**
 * Helper to safely format cell value to string
 */
function formatTimeOnly(cellVal) {
  if (!cellVal) return '-';
  if (cellVal instanceof Date) {
    return Utilities.formatDate(cellVal, 'Asia/Bangkok', 'HH:mm') + ' น.';
  }
  const s = String(cellVal).trim();
  const timeMatch = s.match(/(\d{1,2}):(\d{2})/);
  if (timeMatch) {
    return timeMatch[1].padStart(2, '0') + ':' + timeMatch[2] + ' น.';
  }
  return s;
}

function formatCell(val) {
  if (val === null || val === undefined) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
  }
  return String(val).trim();
}

/**
 * Fetch initial bootstrap data (stations + cases + server time)
 */
function getInitialData() {
  try {
    const sheet = getOrCreateSheet();
    const lastRow = sheet.getLastRow();
    const cases = [];

    if (lastRow > 1) {
      const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const caseObj = {};
        for (let j = 0; j < HEADERS.length; j++) {
          caseObj[HEADERS[j]] = formatCell(row[j]);
        }
        if (caseObj.Case_ID) {
          cases.push(caseObj);
        }
      }
    }

    // Load Ambulance Checklists from Sub-sheet 'Ambulance_Checklist'
    const ambulanceChecklists = {};
    try {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const ambSheet = ss.getSheetByName(AMBULANCE_SHEET_NAME);
      if (ambSheet && ambSheet.getLastRow() > 1) {
        const ambData = ambSheet.getRange(2, 1, ambSheet.getLastRow() - 1, AMBULANCE_HEADERS.length).getValues();
        for (let i = 0; i < ambData.length; i++) {
          const r = ambData[i];
          const stRaw = formatCell(r[2]);
          let matchedStId = '';
          for (let s = 0; s < STATIONS.length; s++) {
            if (STATIONS[s].name === stRaw || STATIONS[s].id === stRaw || stRaw.includes(STATIONS[s].name)) {
              matchedStId = STATIONS[s].id;
              break;
            }
          }
          if (!matchedStId) matchedStId = stRaw;

          const carNoStr = formatCell(r[3]); // e.g. "คันที่ 1" or "คันที่ 2"
          let carIdx = 1;
          const carMatch = carNoStr.match(/\d+/);
          if (carMatch) carIdx = parseInt(carMatch[0], 10);

          let itemsObj = {};
          try {
            itemsObj = JSON.parse(r[13]);
          } catch(e) {}

          const timeFormatted = formatTimeOnly(r[1]);
          const defectsStr = formatCell(r[11]);
          const issuesList = defectsStr ? defectsStr.split('; ').filter(Boolean) : [];

          const chkObj = {
            checklistId: formatCell(r[0]),
            timestamp: formatCell(r[1]),
            evalTime: timeFormatted,
            completedAt: timeFormatted,
            stationId: matchedStId,
            stationName: stRaw,
            carIndex: carIdx,
            vehicleNo: carNoStr || ('คันที่ ' + carIdx),
            vehicleDesc: formatCell(r[4]),
            inspectorName: formatCell(r[5]),
            overallStatus: formatCell(r[6]),
            totalItems: Number(r[7]) || 71,
            readyCount: Number(r[8]) || 0,
            notReadyCount: Number(r[9]) || 0,
            noEquipCount: Number(r[10]) || 0,
            defectsSummary: defectsStr,
            issues: issuesList,
            unreadyNotes: formatCell(r[12]),
            items: itemsObj
          };

          ambulanceChecklists[matchedStId + '_car' + carIdx] = chkObj;
          if (carIdx === 1 || !ambulanceChecklists[matchedStId]) {
            ambulanceChecklists[matchedStId] = chkObj;
          }
        }
      }
    } catch(ambErr) {
      console.warn('Ambulance sheet read warning:', ambErr);
    }

    return {
      success: true,
      stations: STATIONS,
      cases: cases,
      ambulanceChecklists: ambulanceChecklists,
      serverTime: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss')
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Save a new case or update an existing case (Lock-protected against concurrent edits)
 */

/**
 * Save Base64 Photo to Google Drive and return file view URL
 */
function savePhotoToDrive(caseId, photoInput, index) {
  if (!photoInput || typeof photoInput !== 'string') return '';
  // If already a Google Drive link, return it as-is
  if (photoInput.startsWith('http')) return photoInput;
  if (!photoInput.startsWith('data:image')) return '';

  try {
    const folderName = 'CMU_Trekking_2026_Photos';
    let folder = null;

    try {
      const folders = DriveApp.getFoldersByName(folderName);
      if (folders.hasNext()) {
        folder = folders.next();
      } else {
        folder = DriveApp.createFolder(folderName);
      }
    } catch (fErr) {
      Logger.log('Drive folder access warning: ' + fErr.message);
      folder = DriveApp.getRootFolder();
    }

    // Try setting public share if domain policy allows (ignore error if domain blocks it)
    try {
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      Logger.log('Folder public share skipped: ' + shareErr.message);
    }

    const commaIdx = photoInput.indexOf(',');
    if (commaIdx === -1) return '';
    const metaPart = photoInput.slice(0, commaIdx);
    const base64Part = photoInput.slice(commaIdx + 1);

    const matchType = metaPart.match(/:(.*?);/);
    const contentType = matchType ? matchType[1] : 'image/jpeg';
    const bytes = Utilities.base64Decode(base64Part);
    const timeStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss');
    const idxStr = (index !== undefined && index !== null) ? ('_p' + (index + 1)) : '';
    const fileName = (caseId || 'case') + '_' + timeStr + idxStr + '.jpg';
    
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    const file = folder.createFile(blob);
    
    // Try setting public share on file
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (fileShareErr) {
      Logger.log('File public share skipped: ' + fileShareErr.message);
    }
    
    const fileId = file.getId();
    return 'https://drive.google.com/file/d/' + fileId + '/view?usp=drivesdk';
  } catch (e) {
    Logger.log('savePhotoToDrive error: ' + e.message);
    throw new Error('ไม่สามารถบันทึกรูปภาพลง Google Drive ได้: ' + e.message);
  }
}

/**
 * ฟังก์ชันสำหรับกดเรียกใช้ (Run) ใน Apps Script เพื่อขออนุญาตสิทธิ์ Google Drive (ทั้งอ่านและเขียนสร้างไฟล์)
 */
function authorizePermissions() {
  const folderName = 'CMU_Trekking_2026_Photos';
  let folder;
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }
  // ทดสอบสร้างและลบไฟล์ชั่วคราว เพื่อบังคับให้ Google ขอสิทธิ์เขียนไฟล์ (https://www.googleapis.com/auth/drive)
  const tempFile = folder.createFile('auth_test.txt', 'Test permission');
  tempFile.setTrashed(true);
  SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Drive (Full Write/Create) and Spreadsheet authorized successfully!');
}

function saveOrUpdateCase(caseData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const sheet = getOrCreateSheet();
    const nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
    let caseId = caseData.Case_ID ? String(caseData.Case_ID).trim() : '';
    let targetRow = -1;

    const lastRow = sheet.getLastRow();

    // Check if updating existing case
    if (caseId && lastRow > 1) {
      const idColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < idColumn.length; i++) {
        if (String(idColumn[i][0]).trim() === caseId) {
          targetRow = i + 2;
          break;
        }
      }
    }

    const isNew = (targetRow === -1);
    if (isNew) {
      targetRow = lastRow + 1;
      const sequence = Math.max(1, lastRow);
      const paddedSeq = ('0000' + sequence).slice(-4);
      caseId = 'CMU-2026-' + paddedSeq;
      caseData.Timestamp = nowStr;
    }

    caseData.Case_ID = caseId;
    caseData.Last_Updated = nowStr;

    // Handle multiple photos (array, JSON or comma-separated)
    let photoList = [];
    if (Array.isArray(caseData.Photo_List)) {
      photoList = caseData.Photo_List;
    } else if (caseData.Photo_Data && String(caseData.Photo_Data).startsWith('[')) {
      try {
        photoList = JSON.parse(caseData.Photo_Data);
      } catch (e) {
        photoList = [caseData.Photo_Data];
      }
    } else if (caseData.Photo_Data) {
      photoList = [caseData.Photo_Data];
    } else if (caseData.Photo_URL && String(caseData.Photo_URL).startsWith('[')) {
      try {
        photoList = JSON.parse(caseData.Photo_URL);
      } catch (e) {
        photoList = [caseData.Photo_URL];
      }
    } else if (caseData.Photo_URL) {
      if (String(caseData.Photo_URL).includes('http')) {
        photoList = caseData.Photo_URL.split(/[,\n]+/).map(s => s.trim()).filter(s => s.startsWith('http'));
      } else {
        photoList = [caseData.Photo_URL];
      }
    }

    if (photoList.length > 0) {
      const savedUrls = [];
      for (let p = 0; p < photoList.length; p++) {
        const pItem = photoList[p];
        if (pItem) {
          const url = savePhotoToDrive(caseId, pItem, p);
          if (url) savedUrls.push(url);
        }
      }
      caseData.Photo_URL = savedUrls.join(', ');
      caseData.Photo_Data = '';
      caseData.Photo_List = '';
    }

    if (caseData.Case_Type === 'Minor') {
      if (!caseData.Disposition) caseData.Disposition = 'Return to Activity';
      caseData.Status = 'Completed';
    } else {
      if (!caseData.Disposition) {
        caseData.Status = 'Active';
      } else if (caseData.Disposition === 'Transfer to hospital') {
        caseData.Status = 'Transferred';
      } else {
        caseData.Status = 'Completed';
      }
    }

    const rowValues = [];
    for (let j = 0; j < HEADERS.length; j++) {
      const key = HEADERS[j];
      let val = caseData[key];
      if (val === undefined || val === null) val = '';
      rowValues.push(String(val).trim());
    }

    sheet.getRange(targetRow, 1, 1, HEADERS.length).setValues([rowValues]);

    return {
      success: true,
      isNew: isNew,
      caseId: caseId,
      caseData: caseData,
      serverTime: nowStr
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  } finally {
    lock.releaseLock();
  }
}


// ==============================================================================
// AMBULANCE CHECKLIST SUB-SHEET BACKEND
// ==============================================================================
const AMBULANCE_SHEET_NAME = 'Ambulance_Checklist';
const AMBULANCE_HEADERS = [
  'Checklist_ID',
  'Timestamp',
  'Station',
  'Vehicle_No',
  'Vehicle',
  'Inspector',
  'Overall_Status',
  'Total_Items',
  'Ready_Count',
  'Not_Ready_Count',
  'No_Equipment_Count',
  'Defects_Summary',
  'Unready_Notes',
  'Items_Checklist_Data'
];

/**
 * ฟังก์ชันสำหรับสร้างหรือตั้งค่า Sub-sheet 'Ambulance_Checklist' เริ่มต้น
 * สามารถกดรันครั้งเดียวใน Apps Script Console ได้
 */
function setupAmbulanceChecklistSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(AMBULANCE_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(AMBULANCE_SHEET_NAME);
  }

  // เซ็ตหัวคอลัมน์สีเขียวมรกต ตัวอักษรขาว หนา
  sheet.getRange(1, 1, 1, AMBULANCE_HEADERS.length).setValues([AMBULANCE_HEADERS]);
  sheet.getRange(1, 1, 1, AMBULANCE_HEADERS.length)
    .setBackground('#059669')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 35);

  // ปรับความกว้างคอลัมน์ให้อ่านง่าย
  const colWidths = [180, 160, 180, 240, 180, 140, 100, 110, 110, 110, 320, 250];
  colWidths.forEach((w, idx) => {
    sheet.setColumnWidth(idx + 1, w);
  });

  return { success: true, message: "สร้างและจัดรูปแบบชีต 'Ambulance_Checklist' เรียบร้อยแล้ว" };
}

/**
 * บันทึกผลการตรวจเช็กอุปกรณ์รถพยาบาลลงใน Sub-sheet 'Ambulance_Checklist'
 * ใน Spreadsheet เดียวกันกับข้อมูลเคสผู้ป่วย
 */
function saveAmbulanceChecklist(data) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(AMBULANCE_SHEET_NAME);

    // ถ้ายังไม่มี Sub-sheet ให้สร้างใหม่และใส่ Header สีเขียวมรกต
    if (!sheet) {
      sheet = ss.insertSheet(AMBULANCE_SHEET_NAME);
      sheet.getRange(1, 1, 1, AMBULANCE_HEADERS.length).setValues([AMBULANCE_HEADERS]);
      sheet.getRange(1, 1, 1, AMBULANCE_HEADERS.length)
        .setBackground('#059669')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const now = new Date();
    const timestampStr = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    const id = data.id || ('AMB-CHK-' + Utilities.formatDate(now, 'Asia/Bangkok', 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 900 + 100));

    // สร้างชีตหากยังไม่มี
    if (!sheet) {
      setupAmbulanceChecklistSheet();
      sheet = ss.getSheetByName(AMBULANCE_SHEET_NAME);
    }

    const carIdxVal = data.carIndex || data.vehicleNo || 1;
    const carNoStr = 'คันที่ ' + carIdxVal;
    const row = [
      id,
      timestampStr,
      data.stationName || data.stationId || '',
      carNoStr,
      data.vehicleDesc || '',
      data.inspectorName || data.inspector || '',
      data.overallStatus || 'พร้อมใช้งาน',
      data.summaryTotal || data.totalItems || 0,
      data.summaryReady || data.readyCount || 0,
      data.summaryNotReady || data.notReadyCount || 0,
      data.summaryNoEquipment || data.noEquipCount || 0,
      data.defectsSummary || data.issues || '-',
      data.unreadyNotes || '-',
      (typeof data.itemsJson === 'string' ? data.itemsJson : JSON.stringify(data.itemsData || data.items || {}))
    ];

    sheet.appendRow(row);

    return {
      success: true,
      checklistId: id,
      timestamp: timestampStr,
      message: 'บันทึกผลตรวจเช็กรถพยาบาลลงชีต Ambulance_Checklist เรียบร้อยแล้ว'
    };
  } catch (err) {
    return {
      success: false,
      error: err.message
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ฟังก์ชันสำหรับกดรันใน Apps Script Editor เพื่อซ่อมแซมและเติมชื่อคอลัมน์ Row 1 ทั้งหมดให้อัตโนมัติทันที
 */
function fixSheetHeaders() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log('Sheet not found: ' + SHEET_NAME);
    return;
  }
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setValues([HEADERS]);
  headerRange.setBackground('#5b4a8e');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  sheet.setFrozenRows(1);
  Logger.log('Headers successfully repaired and synchronized: ' + HEADERS.join(', '));
}
