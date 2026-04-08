function updateLazadaPriceExact() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("Master");
  const lazadaSheet = ss.getSheetByName("Lazada");

  if (!masterSheet || !lazadaSheet) {
    SpreadsheetApp.getUi().alert("❌ ไม่พบชีต Master หรือ Lazada");
    return;
  }

  // --- 1. ดึงราคากลางจาก Master (จัดการให้เป็นตัวเลขเป๊ะๆ) ---
  const masterData = masterSheet.getDataRange().getDisplayValues();
  let rspToLtdMap = {};
  
  masterData.forEach((row) => {
    let campaignValue = row[5] ? row[5].toString().replace(/[, ]/g, "") : ""; // ลบคอมม่าและเว้นวรรค
    let ltdValue = row[3] ? row[3].toString().replace(/[, ]/g, "") : "";
    
    if (campaignValue && !isNaN(campaignValue)) {
      // ใช้ Math.floor หรือ Math.round เพื่อให้เป็นจำนวนเต็มเป๊ะๆ
      let rspKey = Math.round(Number(campaignValue)).toString();
      rspToLtdMap[rspKey] = Number(ltdValue);
    }
  });

  // --- 2. ตั้งค่าคอลัมน์และแถว ---
  const START_ROW = 5;      
  const PRICE_COL_INDEX = 8; // คอลัมน์ I
  
  // ใช้ getDisplayValues เพื่อให้ได้ค่า "ตามที่ตาเห็น" ใน Lazada
  let lazadaData = lazadaSheet.getDataRange().getDisplayValues();
  let updates = [];
  let changeCount = 0;

  Logger.log("--- เริ่มตรวจสอบ Lazada (Exact Match) ---");

  // --- 3. วนลูปเช็คราคา ---
  for (let i = START_ROW - 1; i < lazadaData.length; i++) {
    let rawPrice = lazadaData[i][PRICE_COL_INDEX];
    let finalPrice = rawPrice; // ค่าเริ่มต้นคือค่าเดิม (String)

    if (rawPrice) {
      // ทำความสะอาดค่าจาก Lazada ให้เหลือแต่ตัวเลขจำนวนเต็ม
      let cleanLazadaPrice = Math.round(Number(rawPrice.toString().replace(/[, ]/g, ""))).toString();

      // เทียบแบบ Exact Match ใน Map
      if (rspToLtdMap[cleanLazadaPrice]) {
        finalPrice = rspToLtdMap[cleanLazadaPrice];
        changeCount++;
        Logger.log("✅ แถว " + (i + 1) + ": เจอคู่! " + cleanLazadaPrice + " -> เปลี่ยนเป็น " + finalPrice);
      }
    }
    updates.push([finalPrice]);
  }

  // --- 4. เขียนข้อมูลกลับ ---
  if (updates.length > 0) {
    // กำหนด Range เฉพาะส่วนที่ประมวลผล (ตั้งแต่แถว START_ROW ลงไป)
    lazadaSheet.getRange(START_ROW, PRICE_COL_INDEX + 1, updates.length, 1).setValues(updates);
    
    Logger.log("--- จบการทำงาน ---");
    Logger.log("เปลี่ยนไปทั้งหมด: " + changeCount + " รายการ");
    SpreadsheetApp.getUi().alert("สำเร็จ! Lazada เปลี่ยนไป " + changeCount + " รายการ");
  }
}
