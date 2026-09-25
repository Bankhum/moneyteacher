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

  // 2. ถ้าเปิดใช้งานผ่านลิงก์ GAS ปกติ (คืนค่าเป็นหน้าเว็บ HTML ตามโค้ดเดิมของคุณ)
  const html = HtmlService.createTemplateFromFile('Index').evaluate();
  html.setTitle('ระบบแจ้งผลเลื่อนเงินเดือน โรงเรียนบ้านคุ้ม(ประสารราษฎร์วิทยา)');
  html.setFaviconUrl('https://img2.pic.in.th/unnamed-4051717258f61fe927.png') 
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return html;
}

// ... โค้ดฟังก์ชันอื่นๆ (checkAndInitSheet, searchEmployee, ฯลฯ) คงไว้ตามเดิม ...
