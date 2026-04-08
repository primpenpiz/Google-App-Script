function updateAllShopeeWithTolerance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("Master");
  
  // ระบุชื่อชีตที่ต้องการจัดการ (ใส่กี่ชื่อก็ได้ใน array นี้)
  const shopeeSheetNames = ["Shopee(1)", "Shopee(2)"];

  if (!masterSheet) {
    SpreadsheetApp.getUi().alert("❌ ไม่พบชีต Master");
    return;
  }

  // --- 1. เตรียมฐานข้อมูล RSP/LTD จาก Master ---
  const masterData = masterSheet.getDataRange().getDisplayValues();
  let rspToLtdMap = {};
  
  masterData.forEach((row) => {
    let rspValue = row[1] ? row[1].toString().replace(/[, ]/g, "") : ""; 
    let ltdValue = row[3] ? row[3].toString().replace(/[, ]/g, "") : "";
    
    if (rspValue && !isNaN(rspValue)) {
      let rspKey = Math.round(Number(rspValue));
      rspToLtdMap[rspKey] = Number(ltdValue);
    }
  });

  let totalChangeCount = 0;
  const START_ROW = 7;      
  const PRICE_COL_INDEX = 6; // คอลัมน์ G (Index 6)

  // --- 2. วนลูปจัดการแต่ละชีต Shopee ---
  shopeeSheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log("⚠️ ไม่พบชีต: " + sheetName);
      return; // ข้ามไปชีตถัดไปถ้าหาไม่เจอ
    }

    Logger.log("--- เริ่มตรวจสอบ " + sheetName + " ---");
    let shopeeData = sheet.getDataRange().getDisplayValues();
    let updates = [];
    let sheetChangeCount = 0;

    for (let i = START_ROW - 1; i < shopeeData.length; i++) {
      let rawPrice = shopeeData[i][PRICE_COL_INDEX];
      let finalPrice = rawPrice; 

      if (rawPrice && rawPrice !== "") {
        let currentPriceNum = Math.round(Number(rawPrice.toString().replace(/[, ]/g, "")));

        // วนเช็ค Logic +/- 1 บาท
        for (let rspKey in rspToLtdMap) {
          let diff = Math.abs(currentPriceNum - Number(rspKey));
          if (diff <= 1) {
            finalPrice = rspToLtdMap[rspKey];
            sheetChangeCount++;
            break; 
          }
        }
      }
      updates.push([finalPrice]);
    }

    // เขียนข้อมูลกลับในแต่ละชีต
    if (updates.length > 0) {
      sheet.getRange(START_ROW, PRICE_COL_INDEX + 1, updates.length, 1).setValues(updates);
      Logger.log("✅ " + sheetName + " อัปเดตไป: " + sheetChangeCount + " รายการ");
      totalChangeCount += sheetChangeCount;
    }
  });

  // --- 3. สรุปผล ---
  SpreadsheetApp.getUi().alert(
    "✅ อัปเดตเสร็จสิ้น!\n" +
    "รวมทั้งหมด: " + totalChangeCount + " รายการ\n" +
    "(ตรวจสอบรายละเอียดได้ใน Log)"
  );
}

/**
 * ฟังก์ชันสร้างเมนูบน Google Sheets
 * จะทำงานอัตโนมัติเมื่อเปิดไฟล์ หรือกด Refresh หน้าเว็บ
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // สร้างเมนูใหม่ชื่อ "🚀 ระบบอัปเดตราคา"
  ui.createMenu('🚀 ระบบอัปเดตราคา')
    // .addItem('อัปเดตทุก Platform (Shopee/Lazada/<TikTok_ยังไม่พร้อมใช้งาน>)', 'updateAllPlatformPrices')
    .addSeparator() // เส้นคั่น
    .addItem('อัปเดตเฉพาะ Shopee', 'updateShopeePriceWithTolerance')
    .addItem('อัปเดตเฉพาะ Lazada', 'updateLazadaPriceOnly')
    .addToUi();
}
