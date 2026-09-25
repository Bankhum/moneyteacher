const SHEET_ID = '1Jng_Lmo_-fnY7wnunFmV44pSBzPtAMhQw84WCaPeGCM'; 
const SHEET_NAME = 'Records';
const SETTINGS_SHEET = 'Settings';
const ADMIN_PASS = 'Bankhum36112361'; 
const DRIVE_FOLDER_ID = '1BzUqviInILFLDgqQnKvxj27mKtOMpU14'; 

function doGet(e) {
  checkAndInitSheet();
  const html = HtmlService.createTemplateFromFile('Index').evaluate();
  html.setTitle('ระบบแจ้งผลเลื่อนเงินเดือน โรงเรียนบ้านคุ้ม(ประสารราษฎร์วิทยา)');
  html.setFaviconUrl('https://img2.pic.in.th/unnamed-4051717258f61fe927.png') 
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return html;
}

function checkAndInitSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  // เพิ่ม 'fiscal_year' เข้าไปที่ส่วนท้ายสุดของ Headers เพื่อความเข้ากันได้กับโครงสร้างเดิม
  const headers = [
    'employee_id', 'full_name', 'school_name', 'position', 'academic_rank', 
    'rank', 'position_number', 'round_name', 'previous_salary', 'calculation_base', 
    'promotion_percentage', 'promotion_amount', 'special_remuneration', 'new_salary', 
    'status', 'reason', 'director_name', 'document_date', 'ack_status', 'ack_time', 'satisfaction', 'fiscal_year'
  ];
  
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
         .setBackground('#1a237e').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  let setSheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!setSheet) {
    setSheet = ss.insertSheet(SETTINGS_SHEET);
    setSheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']])
            .setBackground('#4285f4').setFontColor('#ffffff').setFontWeight('bold');
  }
}

function verifyAdminLogin(username, password) {
  if (password === ADMIN_PASS) return true;
  throw new Error('รหัสผ่านไม่ถูกต้อง');
}

function saveFileToDrive(base64Data, filePrefix) {
  if (!base64Data || !base64Data.startsWith('data:image')) return null;
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const [mimeInfo, base64Str] = base64Data.split(';base64,');
  const ext = mimeInfo.split('/')[1];
  const fileName = filePrefix + '_' + new Date().getTime() + '.' + ext;
  
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Str), mimeInfo.split(':')[1], fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getId();
}

function getAdminSettings() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SETTINGS_SHEET);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  let settings = {};
  data.forEach(row => {
    if (row[0]) settings[row[0]] = row[1];
  });
  return {
    directorName: settings['DIRECTOR_NAME'] || '',
    logoId: settings['LOGO_ID'] || '',
    signatureId: settings['SIGNATURE_ID'] || '',
    stampId: settings['STAMP_ID'] || ''
  };
}

function getBase64FromDriveId(fileId) {
  if (!fileId) return '';
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch(e) {
    return '';
  }
}

function getAdminSettingsPreview() {
  const settings = getAdminSettings();
  return {
    directorName: settings.directorName,
    logoData: getBase64FromDriveId(settings.logoId),
    signatureData: getBase64FromDriveId(settings.signatureId),
    stampData: getBase64FromDriveId(settings.stampId)
  };
}

function saveAdminSettings(settingsData) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SETTINGS_SHEET);
  const data = sheet.getDataRange().getValues();

  const updateOrInsert = (key, val) => {
    if (!val) return;
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(val);
        found = true;
        break;
      }
    }
    if (!found) sheet.appendRow([key, val]);
  };

  if (settingsData.directorName !== undefined) updateOrInsert('DIRECTOR_NAME', settingsData.directorName);
  
  if (settingsData.logo && settingsData.logo.startsWith('data:image')) {
    updateOrInsert('LOGO_ID', saveFileToDrive(settingsData.logo, 'LOGO'));
  }
  if (settingsData.signature && settingsData.signature.startsWith('data:image')) {
    updateOrInsert('SIGNATURE_ID', saveFileToDrive(settingsData.signature, 'SIG'));
  }
  if (settingsData.stamp && settingsData.stamp.startsWith('data:image')) {
    updateOrInsert('STAMP_ID', saveFileToDrive(settingsData.stamp, 'STAMP'));
  }

  return "บันทึกการตั้งค่าลงระบบและ Google Drive เรียบร้อยแล้ว";
}

function searchEmployee(employeeId) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`ไม่พบชีต '${SHEET_NAME}'`);

    const data = sheet.getDataRange().getDisplayValues();
    const headers = data.shift(); 
    
    // ค้นหาทุกรายการที่ตรงกับรหัสพนักงาน (รองรับหลายปีงบประมาณและรอบ)
    const rows = data.filter(r => r[0] == employeeId);
    if (rows.length === 0) throw new Error('ไม่พบข้อมูลสำหรับรหัสประจำตัวนี้');

    const records = rows.map(row => {
      const result = {};
      headers.forEach((header, index) => {
        result[header] = row[index];
      });
      return result;
    });

    const settings = getAdminSettingsPreview();
    
    return { records: records, settings: settings };
  } catch (error) {
    throw new Error(error.message);
  }
}

function getAdminDashboardData() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const rawData = sheet.getDataRange().getValues();
  const headers = rawData.shift();
  
  const list = [];
  
  rawData.forEach(row => {
    const pSal = parseFloat(row[8]) || 0;
    const nSal = parseFloat(row[13]) || 0;
    
    let timeStr = '-';
    if(row[19]) {
      try {
        timeStr = Utilities.formatDate(new Date(row[19]), "Asia/Bangkok", "dd/MM/yyyy HH:mm");
      } catch(e) { timeStr = row[19].toString(); }
    }
    
    list.push({
      id: row[0], name: row[1], prev_salary: pSal, percentage: parseFloat(row[10]) || 0,
      new_salary: nSal, status: row[18] || 'ยังไม่รับทราบ', time: timeStr, satisfaction: row[20] || '-',
      round: row[7] || 'ไม่ระบุ', fiscal_year: row[21] || 'ไม่ระบุ',
      academic_rank: row[4] || 'ไม่มีวิทยฐานะ'
    });
  });

  return { list: list };
}

function saveEmployeeData(formData) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const settings = getAdminSettings();
  
  const newRow = new Array(headers.length).fill('');
  headers.forEach((header, index) => {
    if (formData[header] !== undefined) newRow[index] = formData[header];
  });

  newRow[16] = settings.directorName || ''; 

  let foundIndex = -1;
  // ค้นหาโดยใช้ รหัสพนักงาน + รอบ + ปีงบประมาณ เพื่อแยกแยะรายการที่ซ้ำซ้อนในต่างวาระ
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == formData.employee_id && data[i][7] == formData.round_name && data[i][21] == formData.fiscal_year) {
      foundIndex = i + 1;
      newRow[18] = data[i][18]; // ack_status
      newRow[19] = data[i][19]; // ack_time
      newRow[20] = data[i][20]; // satisfaction
      break;
    }
  }

  if (foundIndex > -1) {
    sheet.getRange(foundIndex, 1, 1, newRow.length).setValues([newRow]);
    return 'อัพเดทข้อมูลเรียบร้อยแล้ว';
  } else {
    sheet.appendRow(newRow);
    return 'เพิ่มข้อมูลใหม่เรียบร้อยแล้ว';
  }
}

function acknowledgeResult(employeeId, roundName, fiscalYear, rating) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == employeeId && data[i][7] == roundName && data[i][21] == fiscalYear) {
      sheet.getRange(i + 1, 19).setValue('รับทราบแล้ว');
      sheet.getRange(i + 1, 20).setValue(new Date());
      sheet.getRange(i + 1, 21).setValue(rating);
      return true;
    }
  }
  return false;
}
