const SHEET_ID = '1Jng_Lmo_-fnY7wnunFmV44pSBzPtAMhQw84WCaPeGCM'; 
const SHEET_NAME = 'Records';
const SETTINGS_SHEET = 'Settings';
const ADMIN_PASS = 'Bankhum36112361'; 
const DRIVE_FOLDER_ID = '1BzUqviInILFLDgqQnKvxj27mKtOMpU14'; 

// ปรับแก้ doGet ใหม่ให้รองรับทั้งการเปิดเว็บผ่าน GAS และการดึงข้อมูลจาก GitHub
function doGet(e) {
  checkAndInitSheet();

  // 1. ถ้ามีการร้องขอข้อมูลจาก GitHub (ทำงานแบบ API)
  if (e && e.parameter && e.parameter.action) {
    try {
      let action = e.parameter.action;
      let result = {};

      if (action === 'searchEmployee') {
        result = searchEmployee(e.parameter.employeeId);
      } else if (action === 'getAdminSettings') {
        result = getAdminSettingsPreview();
      }
      // หากมีฟังก์ชันอื่นที่ต้องการดึงข้อมูล ให้เพิ่มเงื่อนไข else if ตรงนี้

      // สำคัญมาก: ต้องคืนค่าเป็น JSON เท่านั้น GitHub ถึงจะอ่านได้
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: result }))
        .setMimeType(ContentService.MimeType.JSON);

    } catch (error) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
function checkAndInitSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  // เพิ่ม 'fiscal_year' เข้าไปที่ส่วนท้ายสุดของ Headers
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
  // 2. ถ้าเปิดใช้งานผ่านลิงก์ GAS ปกติ (คืนค่าเป็นหน้าเว็บ HTML ตามโค้ดเดิมของคุณ)
  const html = HtmlService.createTemplateFromFile('Index').evaluate();
  html.setTitle('ระบบแจ้งผลเลื่อนเงินเดือน โรงเรียนบ้านคุ้ม(ประสารราษฎร์วิทยา)');
  html.setFaviconUrl('https://img2.pic.in.th/unnamed-4051717258f61fe927.png') 
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return html;
}

// ... โค้ดฟังก์ชันอื่นๆ (checkAndInitSheet, searchEmployee, ฯลฯ) คงไว้ตามเดิม ...
