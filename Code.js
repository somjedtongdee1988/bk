/**
 * ระบบยืม-คืนพัสดุและโสตทัศนูปกรณ์อัจฉริยะ CPE มรพส. (เวอร์ชันตะกร้า E-Commerce - Ultimate Production 2026)
 * พัฒนาโดย: AI พัฒนาระบบด้วย GAS (Senior Apps Script Specialist)
 */

const PROFILE_FOLDER_ID = "1PbgnS8eZXOdKXLPFM-XSeBwEQCnKQYh1";
const ITEM_FOLDER_ID = "1bFnS6npqJXKnpzuL8YjFBo7K3GqQ_yW5";
const SPREADSHEET_ID = "1rK7WMeaicIUnvMVVQHdn5gxaIXwlv-AzFSVGe6CwiX8"; 
const ADMIN_EMAIL_DEFAULT = "somjedtongdee@psru.ac.th";
const SYSTEM_EMAIL = "somjedtongdee@psru.ac.th";

function uploadFileToDrive(username, base64Data, folderId) {
  if (!base64Data || base64Data === "" || !base64Data.includes("base64,")) return "";
  try {
    const folder = DriveApp.getFolderById(folderId);
    const splitData = base64Data.split("base64,");
    const contentType = splitData[0].split(":")[1].split(";")[0];
    const decodedBytes = Utilities.base64Decode(splitData[1]);
    const fileName = "img_" + username.trim() + "_" + new Date().getTime() + ".jpg";
    const blob = Utilities.newBlob(decodedBytes, contentType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return "https://drive.google.com/thumbnail?sz=w200&id=" + file.getId();
  } catch (e) {
    Logger.log("Drive Upload Error: " + e.toString());
    return "";
  }
}

function initDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName("Items")) {
    const itemSheet = ss.insertSheet("Items");
    itemSheet.appendRow(["itemID", "itemName", "itemType", "quantity", "location", "status", "itemPic", "qrCode"]);
    itemSheet.appendRow(["CPE-001", "Projector Sony 4K", "ครุภัณฑ์", 5, "ห้องแล็บ 404", "พร้อมใช้งาน", "", ""]);
    itemSheet.appendRow(["CPE-002", "Visualizer", "ครุภัณฑ์", 2, "ห้องพักอาจารย์", "พร้อมใช้งาน", "", ""]);
  }
  if (!ss.getSheetByName("Transactions")) {
    const transSheet = ss.insertSheet("Transactions");
    // เพิ่มคอลัมน์ K (col 11) สำหรับเก็บ Email ผู้ยืม (ใช้ส่งเมลจากคอลัมน์ K)
    transSheet.appendRow(["transID","itemID","borrowerName","borrowerEmail","borrowDate","dueDate","returnDate","status","purpose","borrowQty","borrowerEmail_K"]);
  }
  if (!ss.getSheetByName("Users")) {
    const userSheet = ss.insertSheet("Users");
    userSheet.appendRow(["username", "role", "password", "fullName", "email", "phone", "position", "profilePic"]);
    userSheet.appendRow(["admin@psru.ac.th", "ADMIN", "123456", "ผู้ดูแลระบบ คลังพัสดุ", "admin@psru.ac.th", "055-111111", "อาจารย์ประจำสาขา", ""]);
    userSheet.appendRow(["boss@psru.ac.th", "SUPER_ADMIN", "654321", "หัวหน้าสาขาวิศวกรรม", "boss@psru.ac.th", "055-222222", "ผู้บริหาร", ""]);
    userSheet.appendRow(["user@psru.ac.th", "USER", "111111", "", "", "", "", ""]);
  }
}

function doGet() {
  initDatabase();
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('ระบบยืม-คืนพัสดุ CPE มรพส.')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// 2. EMAIL UTILITY FUNCTIONS (ฟังก์ชันระบบอีเมล)
// ==========================================

/**
 * ค้นหาอีเมลของผู้ใช้งานจาก Username โดยดึงข้อมูลจากชีต 'Users'
 * @param {string} username - ชื่อผู้ใช้ที่ต้องการค้นหา
 * @return {string|null} อีเมลของผู้ใช้ หรือ null ถ้าไม่พบ
 */
function getUserEmail(identifier) {
  try {
    if (!identifier) return null;
    const candidate = String(identifier).trim();
    if (candidate.indexOf('@') !== -1) return candidate; // ถ้าเป็นอีเมลแล้ว ให้คืนเลย

    const res = _readSheetAll('Users');
    const data = res.values || [];
    const target = candidate.toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || !row[0]) continue;
      const username = String(row[0] || "").trim().toLowerCase();
      const fullName = String(row[3] || "").trim().toLowerCase();
      const email = String(row[4] || "").trim();
      if (username === target || fullName === target) {
        return email || null;
      }
      // หาก identifier ตรงกับ email ในชีต ให้คืนค่าเลย
      if (email && email.toLowerCase() === target) return email;
    }
  } catch (error) {
    Logger.log("Error ในการค้นหาอีเมล: " + error.toString());
  }
  return null;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function uploadProfileImageToDrive(username, base64Data) {
  if (!base64Data || base64Data === "" || !base64Data.includes("base64,")) {
    return "";
  }
  try {
    return uploadFileToDrive(username, base64Data, PROFILE_FOLDER_ID);
  } catch (e) {
    Logger.log("Drive Upload Error: " + e.toString());
    return "";
  }
}
/**
 * [TARGET FIX] ฟังก์ชันรวมยอดพัสดุทั้งหมดในตะกร้าคำขอเดียวกัน 
 * เพื่อส่งอีเมลตอบกลับผลลัพธ์หาผู้ยืม 1 ฉบับเท่านั้น เมื่อเจ้าหน้าที่พิจารณาครบถ้วนทุกชิ้นแล้ว
 */
function checkAndSendSummaryEmailToUser(transId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const transSheet = ss.getSheetByName("Transactions");
    const itemSheet = ss.getSheetByName("Items");
    const transData = transSheet.getDataRange().getValues();
    const itemData = itemSheet.getDataRange().getValues();

    const itemMap = {};
    for (let i = 1; i < itemData.length; i++) {
      if (itemData[i][0]) itemMap[itemData[i][0].toString().trim()] = itemData[i][1];
    }

    let basketItems = [];
    let hasPending = false;
    let borrowerName = "";
    let borrowerEmail = "";

    for (let j = 1; j < transData.length; j++) {
      if (String(transData[j][0]).trim() === String(transId).trim()) {
        borrowerName = transData[j][2] || borrowerName;
        // index 10 = col K (0-based)
        borrowerEmail = String(transData[j][10] || borrowerEmail || "").trim();
        const status = transData[j][7] || "";
        if (status === "รออนุมัติ") { hasPending = true; }
        basketItems.push({
          itemId: transData[j][1],
          itemName: itemMap[String(transData[j][1]).trim()] || "",
          qty: transData[j][9] || 1,
          status: status
        });
      }
    }

    // หากยังมีรายการรออนุมัติ ให้ยังไม่ส่งสรุป (เพียงอัปเดตสถานะใน sheet เท่านั้น)
    if (hasPending) {
      Logger.log("checkAndSendSummaryEmailToUser: trans %s still has pending items; skip sending summary", transId);
      return;
    }

    // fallback: หา email จาก Users ถ้ายังว่าง
    if ((!borrowerEmail || borrowerEmail === "") && borrowerName) {
      const uEmail = getUserEmail(borrowerName);
      if (uEmail) borrowerEmail = uEmail;
    }

    if (!borrowerEmail || basketItems.length === 0) {
      Logger.log("checkAndSendSummaryEmailToUser: no email or no items for trans " + transId);
      return;
    }

    // สร้าง HTML รายการสรุป (แยกแสดง Approved / Rejected)
    let rowsHtml = "";
    basketItems.forEach(it => {
      const badge = (it.status === "กำลังยืม") 
        ? `<span style="color:#059669;font-weight:bold;">✅ อนุมัติ</span>` 
        : `<span style="color:#e11d48;font-weight:bold;">❌ ไม่อนุมัติ</span>`;
      rowsHtml += `<tr>
        <td style="padding:8px;border:1px solid #e2e8f0;font-family:monospace;">${it.itemId}</td>
        <td style="padding:8px;border:1px solid #e2e8f0;">${it.itemName}</td>
        <td style="padding:8px;border:1px solid #e2e8f0;text-align:center;">${it.qty}</td>
        <td style="padding:8px;border:1px solid #e2e8f0;text-align:center;">${badge}</td>
      </tr>`;
    });

    const subject = `✅ [สรุปผลการพิจารณา] แจ้งสถานะคำขอยืมพัสดุ เลขธุรกรรม: ${transId}`;
    const htmlBody = `
      <div style="font-family:Arial,Helvetica,sans-serif;color:#111;">
        <h3>สรุปผลการพิจารณาคำขอยืมพัสดุ (หมายเลข: ${transId})</h3>
        <p>เรียนคุณ <strong>${borrowerName || '-'}</strong></p>
        <p>สถานะรวมของรายการที่ร้องขอมีดังนี้</p>
        <table style="width:100%;border-collapse:collapse;margin-top:10px;">
          <thead>
            <tr>
              <th style="padding:8px;border:1px solid #e2e8f0;background:#f8fafc;text-align:left;">รหัสพัสดุ</th>
              <th style="padding:8px;border:1px solid #e2e8f0;background:#f8fafc;text-align:left;">ชื่อพัสดุ</th>
              <th style="padding:8px;border:1px solid #e2e8f0;background:#f8fafc;text-align:center;">จำนวน</th>
              <th style="padding:8px;border:1px solid #e2e8f0;background:#f8fafc;text-align:center;">ผลการพิจารณา</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
        </table>
        <p style="margin-top:12px;">หากได้รับอนุมัติ กรุณาติดต่อเจ้าหน้าที่เพื่อนัดรับพัสดุ (นายสมเจตน์ ทองดี) โทร.081-673-8958</p>
        <p style="color:#6b7280;font-size:0.9em;">ระบบ CPE Smart Asset Management</p>
      </div>
    `;

    try {
      GmailApp.sendEmail(borrowerEmail, subject, "", { htmlBody: htmlBody, replyTo: SYSTEM_EMAIL, name: "CPE Smart Asset Management" });
      Logger.log("checkAndSendSummaryEmailToUser: summary sent to %s for trans %s", borrowerEmail, transId);
    } catch (e) {
      Logger.log("checkAndSendSummaryEmailToUser: failed to send summary email: " + e.toString());
      try { GmailApp.sendEmail(ADMIN_EMAIL_DEFAULT, `⚠️ Email send failed for ${transId}`, `Error: ${e.toString()}\nBorrowerEmail:${borrowerEmail}`); } catch(e2){Logger.log(e2.toString());}
    }

  } catch (e) {
    Logger.log("Error ในการจัดส่งอีเมลสรุปผลรวมหาผู้ยืม: " + e.toString());
  }
}

function _findTransactionById(transId) {
  try {
    const ss = _openSS();
    const sh = ss.getSheetByName('Transactions');
    if (!sh) return null;
    const lastRow = sh.getLastRow();
    const lastCol = Math.max(1, sh.getLastColumn());
    if (lastRow < 2) return null;
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const vals = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0]) === String(transId)) {
        const obj = {};
        for (let c = 0; c < headers.length; c++) obj[headers[c]] = vals[i][c];
        return { sheet: sh, row: i + 2, data: obj, headers: headers };
      }
    }
  } catch (e) {
    Logger.log("_findTransactionById error: " + e.toString());
  }
  return null;
}

/**
 * หา username จากชีต Users โดยรับ identifier ที่อาจเป็น username / fullName / email
 */
function findUsernameByIdentifier(identifier) {
  try {
    if (!identifier) return "";
    const id = String(identifier).trim().toLowerCase();
    const users = (_readSheetAll('Users').values) || [];
    for (let r = 1; r < users.length; r++) {
      const row = users[r] || [];
      const username = String(row[0] || "").trim();
      const fullName = String(row[3] || "").trim().toLowerCase();
      const email = String(row[4] || "").trim().toLowerCase();
      if (!username) continue;
      if (username.toLowerCase() === id) return username;
      if (fullName && fullName === id) return username;
      if (email && email === id) return username;
    }
  } catch (e) {
    Logger.log("findUsernameByIdentifier error: " + e.toString());
  }
  return "";
}

/**
 * ส่งอีเมลแจ้งผู้ยืมเมื่ออนุมัติ
 */
function sendApprovalEmail(transId) {
  try {
    const tx = _findTransactionById(transId);
    if (!tx) { Logger.log("sendApprovalEmail: transaction not found: " + transId); return false; }
    const d = tx.data;
    const borrowerEmail = _resolveBorrowerEmail(d);
    if (!borrowerEmail) { Logger.log("sendApprovalEmail: no borrower email for " + transId); return false; }

    const itemId = d['itemID'] || d['itemId'] || d['รหัสพัสดุ'] || "";
    const itemName = d['itemName'] || d['ชื่อพัสดุ'] || "";
    const borrowDate = d['borrowDate'] || d['วันที่ยืม'] || "";
    const dueDate = d['dueDate'] || d['กำหนดคืน'] || "";

    const subject = `แจ้งผลการอนุมัติคำขอยืมพัสดุ ${itemId}`;
    const htmlBody = `<p>เรียนคุณ ${d['borrowerName'] || ''}</p>
      <p>คำขอการยืมพัสดุ <strong>${itemId} ${itemName}</strong> ได้รับการอนุมัติแล้ว</p>
      <ul>
        <li>วันที่ยืม: ${borrowDate}</li>
        <li>กำหนดคืน: ${dueDate}</li>
      </ul>
      <p>ขอบคุณครับ/ค่ะ</p>`;

    GmailApp.sendEmail(borrowerEmail, subject, "", { htmlBody: htmlBody, replyTo: SYSTEM_EMAIL, name: "CPE Smart Asset Management" });
    return true;
  } catch (e) {
    Logger.log('sendApprovalEmail error: ' + e.toString());
    return false;
  }
}

/**
 * ส่งอีเมลแจ้งผู้ยืมเมื่อปฏิเสธ พร้อมเหตุผล
 */
function sendRejectionEmail(transId, reason) {
  try {
    const tx = _findTransactionById(transId);
    if (!tx) { Logger.log("sendRejectionEmail: transaction not found: " + transId); return false; }
    const d = tx.data;
    const borrowerEmail = _resolveBorrowerEmail(d);
    if (!borrowerEmail) { Logger.log("sendRejectionEmail: no borrower email for " + transId); return false; }

    const itemId = d['itemID'] || d['itemId'] || d['รหัสพัสดุ'] || "";
    const itemName = d['itemName'] || d['ชื่อพัสดุ'] || "";

    const subject = `แจ้งผลการขอยืมพัสดุ ${itemId} - ไม่อนุมัติ`;
    const htmlBody = `<p>เรียนคุณ ${d['borrowerName'] || ''}</p>
      <p>คำขอการยืมพัสดุ <strong>${itemId} ${itemName}</strong> ถูกปฏิเสธ</p>
      <p><strong>เหตุผล:</strong> ${reason || 'ไม่ระบุ'}</p>
      <p>หากต้องการข้อมูลเพิ่มเติม กรุณาติดต่อเจ้าหน้าที่</p>`;

    GmailApp.sendEmail(borrowerEmail, subject, "", { htmlBody: htmlBody, replyTo: SYSTEM_EMAIL, name: "CPE Smart Asset Management" });
    return true;
  } catch (e) {
    Logger.log('sendRejectionEmail error: ' + e.toString());
    return false;
  }
}

/**
 * ส่งอีเมลแจ้งเตือนผู้ใช้งานเมื่อคำขอยืมได้รับการ อนุมัติ หรือ ปฏิเสธ
 * @param {string} username - ชื่อผู้ขอยืม
 * @param {string} itemCode - รหัสพัสดุ (เช่น CPE-002)
 * @param {string} itemName - ชื่อพัสดุ
 * @param {string} status - สถานะ ('approved' หรือ 'rejected')
 */
function sendApprovalEmailToUser(username, itemCode, itemName, status) {
  try {
    const userEmail = getUserEmail(username);
    if (!userEmail) {
      Logger.log("sendApprovalEmailToUser: no email for username=" + username);
      return false;
    }

    var subject = "";
    var statusText = "";
    var statusColor = "";
    var noteText = "";

    if (status === "approved" || status === "อนุมัติ") {
      subject = "✅ [อนุมัติ] ผลการคำขอยืมพัสดุระบบ CPE มรพส.";
      statusText = "ได้รับการอนุมัติ";
      statusColor = "#28a745";
      noteText = "กรุณาติดต่อรับพัสดุตามประกาศของเจ้าหน้าที่";
    } else {
      subject = "❌ [ปฏิเสธ] ผลการคำขอยืมพัสดุระบบ CPE มรพส.";
      statusText = "ปฏิเสธการอนุมัติ";
      statusColor = "#dc3545";
      noteText = "หากมีข้อสงสัยกรุณาติดต่อเจ้าหน้าที่คลังพัสดุ";
    }

    var htmlBody = `
      <div style="font-family: Arial,Helvetica,sans-serif; padding:16px; color:#333; max-width:680px;">
        <h3 style="color:#0b5394;">แจ้งเตือนสถานะการยืมพัสดุ</h3>
        <p>สวัสดีคุณ <b>${username}</b></p>
        <table style="width:100%; border-collapse:collapse; margin:12px 0;">
          <tr><td style="padding:8px; border:1px solid #e5e7eb; font-weight:bold; width:30%;">รหัสพัสดุ</td><td style="padding:8px; border:1px solid #e5e7eb;">${itemCode}</td></tr>
          <tr><td style="padding:8px; border:1px solid #e5e7eb; font-weight:bold;">ชื่อพัสดุ</td><td style="padding:8px; border:1px solid #e5e7eb;">${itemName}</td></tr>
          <tr><td style="padding:8px; border:1px solid #e5e7eb; font-weight:bold;">ผลการพิจารณา</td><td style="padding:8px; border:1px solid #e5e7eb; color:${statusColor}; font-weight:bold;">${statusText}</td></tr>
        </table>
        <p style="background:#fff3cd;padding:10px;border-radius:4px;color:#856404;">${noteText}</p>
        <p style="font-size:0.85em;color:#6b7280;">ระบบ CPE Smart Asset Management</p>
      </div>
    `;

    GmailApp.sendEmail(userEmail, subject, "", { htmlBody: htmlBody, replyTo: SYSTEM_EMAIL, name: "CPE Smart Asset Management" });
    Logger.log("sendApprovalEmailToUser: sent to " + userEmail + " (username=" + username + ")");
    return true;
  } catch (e) {
    Logger.log("sendApprovalEmailToUser error: " + e.toString());
    return false;
  }
}

/**
 * ส่งอีเมลแจ้งเตือนผู้ดูแลระบบ (Admin) เมื่อมีผู้ใช้งานส่งคำขอยืมเข้ามาใหม่
 * @param {Object} txData - ข้อมูลธุรกรรมการยืม
 */
function sendNotificationToAdmin(txData) {
  var subject = "🚨 [คำขอใหม่] มีรายการขอยืมพัสดุรอการอนุมัติ";
  
  var htmlBody = `
    <div style="font-family: Arial, sans-serif; padding: 20px; max-width: 600px; border: 1px solid #ddd;">
      <h3 style="color: #d9534f;">มีรายการขออนุมัติยืมพัสดุใหม่เข้ามาในระบบ</h3>
      <p><b>ผู้ขอยืม:</b> ${txData.username || 'ไม่ระบุชื่อ'}</p>
      <p><b>รหัสพัสดุ:</b> ${txData.itemCode}</p>
      <p><b>ชื่อพัสดุ:</b> ${txData.itemName}</p>
      <br>
      <p>กรุณาเข้าสู่ระบบเพื่อดำเนินการตรวจสอบและกดอนุมัติหรือปฏิเสธคำขอ</p>
      <p><a href="${ScriptApp.getService().getUrl()}" style="background-color: #0275d8; color: white; padding: 10px 15px; text-decoration: none; border-radius: 4px; display: inline-block;">เปิดระบบจัดการ (Web App)</a></p>
    </div>
  `;
  
  MailApp.sendEmail({
    to: ADMIN_EMAIL_DEFAULT,
    subject: subject,
    htmlBody: htmlBody
  });
}

// ==========================================
// 3. CORE TRANSACTION LOGIC (ส่วนบันทึก/อนุมัติข้อมูล)
// ==========================================

/**
 * ฟังก์ชันสำหรับรับรายการยืมพัสดุจากหน้าบ้าน (User ส่งคำขอ)
 */
function createTransaction(username, itemCode, itemName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("Transactions");
    
    // โค้ดส่วนบันทึกข้อมูลลงชีตของคุณเดิม...
    // เช่น sheet.appendRow([new Date(), username, itemCode, itemName, "รออนุมัติ"]);
    
    // [เพิ่มระบบอีเมล] ส่งแจ้งเตือนหา Admin ทันทีที่มีคำขอใหม่
    sendNotificationToAdmin({username: username, itemCode: itemCode, itemName: itemName});
    
    return { success: true, message: "ส่งคำขอยืมสำเร็จและแจ้งเตือนผู้ดูแลระบบแล้ว" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * ฟังก์ชันอัปเดตสถานะการอนุมัติ (เมื่อ Admin กดปุ่ม "ตกลง" หรืออนุมัติพัสดุ)
 * ค้นหาข้อมูลแถวที่บันทึกพัสดุชิ้นนั้นๆ แล้วทำการอัปเดตสถานะ พร้อมส่งเมลหา User
 */
function updateApprovalStatus(txIdOrUsername, itemCode, status) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("Transactions");
    var data = sheet.getDataRange().getValues();
    
    var updated = false;
    var targetUsername = "";
    var targetItemName = "";
    
    // ค้นหาแถวธุรกรรมในชีตเพื่ออัปเดต (สมมติค้นหาจาก รหัสพัสดุ และสถานะเดิมที่เป็น 'รออนุมัติ')
    for (var i = 1; i < data.length; i++) {
      // ตัวอย่าง: คอลัมน์รหัสพัสดุ (เช่น index 2) และ สถานะปัจจุบัน (เช่น index 4)
      if (data[i][2] == itemCode && (data[i][4] == "รออนุมัติ" || data[i][4] == "Pending")) {
        
        targetUsername = data[i][1]; // ดึงชื่อผู้ใช้เก็บไว้ส่งอีเมล (คอลัมน์ B)
        targetItemName = data[i][3]; // ดึงชื่อพัสดุเก็บไว้ส่งอีเมล (คอลัมน์ D)
        
        // อัปเดตช่องสถานะในชีต (เช่น คอลัมน์ E แถวที่ i+1)
        var displayStatus = (status === "approved") ? "อนุมัติแล้ว" : "ปฏิเสธ";
        sheet.getRange(i + 1, 5).setValue(displayStatus); 
        
        updated = true;
        break;
      }
    }
    
    // กรณีที่โค้ดเดิมของคุณไม่ได้ใช้ระบบวนลูปค้นหา แต่รับค่า Username มาโดยตรง
    if (!updated) {
      targetUsername = txIdOrUsername;
      targetItemName = "พัสดุรหัส " + itemCode;
    }
    
    // [เพิ่มระบบอีเมล] เรียกฟังก์ชันส่งอีเมลแจ้งผู้ใช้โดยตรงหลังจากบันทึกชีตสำเร็จ
    sendApprovalEmailToUser(targetUsername, itemCode, targetItemName, status);
    
    return { success: true, message: "อัปเดตสถานะและส่งอีเมลแจ้งเตือนเรียบร้อยแล้ว" };
    
  } catch (error) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.toString() };
  }
}

function processFormLogin(username, password) {
  try {
    initDatabase();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const userSheet = ss.getSheetByName("Users");
    const userData = userSheet.getDataRange().getValues();
    const targetUser = username ? username.trim().toLowerCase() : "";
    const targetPass = password ? password.toString().trim() : "";

    for (let i = 1; i < userData.length; i++) {
      if (!userData[i][0] || userData[i][2] === undefined) continue;
      const sheetUser = String(userData[i][0]).trim().toLowerCase();
      const sheetRole = String(userData[i][1]).trim();
      const sheetPass = String(userData[i][2]).trim();

      if (sheetUser === targetUser && sheetPass === targetPass) {
        const profile = {
          fullName: userData[i][3] || "",
          email: userData[i][4] || "",
          phone: userData[i][5] || "",
          position: userData[i][6] || "",
          profilePic: userData[i][7] || "",
          idCode: userData[i][8] || "", 
          major: userData[i][9] || "",
          lineId: userData[i][10] || ""
        };
        const isProfileComplete = (profile.fullName !== "" && profile.email !== "" && profile.phone !== "" && profile.position !== "" && profile.profilePic !== "");
        return { success: true, username: userData[i][0], role: sheetRole, profile: profile, isProfileComplete: isProfileComplete, message: "เข้าสู่ระบบสำเร็จ" };
      }
    }
    return { success: false, message: "❌ ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง" };
  } catch (e) { return { success: false, message: "เกิดข้อผิดพลาด: " + e.toString() }; }
}

function registerNewUser(username, password, fullName, email, profilePicBase64) {
  try {
    initDatabase();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Users");
    const data = sheet.getDataRange().getValues();
    const targetUser = username.trim().toLowerCase();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toLowerCase() === targetUser) {
        return { success: false, message: "❌ Username นี้มีอยู่ในระบบแล้ว" };
      }
    }
    const imageUrl = uploadFileToDrive(targetUser, profilePicBase64, PROFILE_FOLDER_ID);
    sheet.appendRow([targetUser, "USER", password.trim(), fullName.trim(), email.trim(), "", "", imageUrl]);
    return { success: true, message: "🎉 สมัครสมาชิกสำเร็จ! กรุณาล็อกอินเพื่อตั้งค่าโปรไฟล์ครั้งแรก" };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function updateProfileData(username, fullName, email, phone, position, idCode, major, lineId, profilePicBase64) {
  try {
    initDatabase();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Users");
    const data = sheet.getDataRange().getValues();
    const targetUser = username ? username.toString().trim().toLowerCase() : "";

    if (targetUser === "") {
      return { success: false, message: "❌ ไม่สามารถอัปเดตได้เนื่องจากไม่ระบุบัญชีผู้ใช้" };
    }

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const sheetUser = data[i][0].toString().trim().toLowerCase();

      if (sheetUser === targetUser) {
        const row = i + 1;
        sheet.getRange(row, 4).setValue(fullName.trim());
        sheet.getRange(row, 5).setValue(email.trim());
        sheet.getRange(row, 6).setValue(phone.trim());
        sheet.getRange(row, 7).setValue(position.trim());
        
        sheet.getRange(row, 9).setValue(idCode ? idCode.trim() : "");
        sheet.getRange(row, 10).setValue(major ? major.trim() : "");
        sheet.getRange(row, 11).setValue(lineId ? lineId.trim() : "");

        if (profilePicBase64 && profilePicBase64 !== "") {
          const newImageUrl = uploadFileToDrive(sheetUser, profilePicBase64, PROFILE_FOLDER_ID);
          if (newImageUrl !== "") {
            sheet.getRange(row, 8).setValue(newImageUrl);
          }
        }
        return { success: true, message: "🔒 อัปเดตข้อมูลโปรไฟล์ส่วนตัวของคุณเรียบร้อยแล้ว!" };
      }
    }
    return { success: false, message: "❌ ไม่พบบัญชีผู้ใช้งานนี้ในระบบคลังพัสดุ (" + targetUser + ")" };
  } catch (e) {
    return { success: false, message: "เกิดข้อผิดพลาดบนเซิร์ฟเวอร์: " + e.toString() };
  }
}

function getUserLoanHistory(username, role, page, pageSize, full) {
  try {
    page = parseInt(page, 10) || 1;
    pageSize = parseInt(pageSize, 10) || 5; // UI ต้องการ 5 รายการเริ่มต้น
    full = !!full; // ถ้า true -> อ่านทั้งชีต
    const neededCount = page * pageSize;

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const transSheet = ss.getSheetByName("Transactions");
    const itemMap = _getItemMapCached(300);

    if (!transSheet) return { success: true, history: [], total: 0, page, pageSize, isGlobalView: false, hasMore: false };

    const lastTransRow = transSheet.getLastRow();
    const lastTransCol = Math.max(1, transSheet.getLastColumn());
    if (lastTransRow <= 1) return { success: true, history: [], total: 0, page, pageSize, isGlobalView: false, hasMore: false };

    // ถ้าไม่ขอ full ให้จำกัดจำนวนแถวที่จะอ่านจากท้ายชีต (ลด I/O)
    const SCAN_LIMIT = 1000; // ปรับได้ตามขนาดข้อมูล
    let readStartRow = 1;
    let readRowCount = lastTransRow;
    if (!full) {
      readRowCount = Math.min(lastTransRow, Math.max(neededCount, SCAN_LIMIT));
      readStartRow = Math.max(1, lastTransRow - readRowCount + 1);
    }

    const transData = transSheet.getRange(readStartRow, 1, readRowCount, lastTransCol).getValues();

    const targetUser = username ? username.trim().toLowerCase() : "";
    const userRole = role ? role.trim().toUpperCase() : "USER";
    const ssTimeZone = ss.getSpreadsheetTimeZone();

    // เก็บตำแหน่งแถว (index ใน transData) ของรายการที่ตรงเงื่อนไข (วนจากท้ายที่อ่านได้)
    const matchedIndices = [];
    for (let i = transData.length - 1; i >= 1; i--) {
      const row = transData[i];
      if (!row || !row[1]) continue;
      const transUser = row[3] ? String(row[3]).trim().toLowerCase() : "";
      if (userRole === "ADMIN" || userRole === "SUPER_ADMIN" || transUser === targetUser) {
        matchedIndices.push(i);
        // ถ้าไม่ได้อ่านทั้งหมดและเก็บครบพอสำหรับหน้า requested ก็หยุด (performance)
        if (!full && matchedIndices.length >= neededCount) break;
      }
    }

    // คำนวน totalMatches และ hasMore ให้ครอบคลุมกรณีที่อ่านทั้งชีตแล้วด้วย
    let totalMatches = matchedIndices.length;
    let hasMore = false;
    if (!full) {
      if (readStartRow > 1) {
        // อ่านเฉพาะหน้าต่างท้าย: ถ้าได้ครบ neededCount มีความเป็นไปได้ว่ามีเพิ่ม
        hasMore = matchedIndices.length >= neededCount;
      } else {
        // อ่านทั้งชีตแล้ว: ถ้าจำนวนรายการมากกว่า pageSize ให้แสดงปุ่ม "ดูเพิ่มเติมทั้งหมด"
        hasMore = matchedIndices.length > pageSize;
      }
    } else {
      // full = true -> โหลดทั้งหมดแล้ว ไม่มี more
      hasMore = false;
    }

    // สร้างรายการที่จะส่งกลับเฉพาะสำหรับหน้า (pagination) โดยแปลง matchedIndices เป็นข้อมูลจริง
    const historyList = [];
    const startIndex = (page - 1) * pageSize;
    for (let k = startIndex; k < Math.min(matchedIndices.length, startIndex + pageSize); k++) {
      const idx = matchedIndices[k];
      const row = transData[idx];

      const bRaw = row[4];
      const dRaw = row[5];
      const rRaw = row[6];

      const borrowDate = bRaw ? Utilities.formatDate(bRaw instanceof Date ? bRaw : new Date(bRaw), ssTimeZone, "dd/MM/yyyy HH:mm") : "-";
      const dueDate = dRaw ? Utilities.formatDate(dRaw instanceof Date ? dRaw : new Date(dRaw), ssTimeZone, "dd/MM/yyyy") : "-";
      const returnDate = rRaw ? Utilities.formatDate(rRaw instanceof Date ? rRaw : new Date(rRaw), ssTimeZone, "dd/MM/yyyy HH:mm") : "-";

      const borrowerEmail = row[3] ? String(row[3]).trim().toLowerCase() : "";

      historyList.push({
        transId: row[0],
        itemId: row[1],
        itemName: itemMap[String(row[1]).trim()] || "ไม่พบชื่อพัสดุในคลัง",
        borrowerName: row[2] ? String(row[2]).trim() : "ไม่ระบุชื่อ",
        borrowerEmail: borrowerEmail,
        borrowDate: borrowDate,
        dueDate: dueDate,
        returnDate: returnDate,
        status: row[7],
        purpose: row[8] || "-",
        qty: row[9] || 1
      });
    }

    const isGlobalView = (userRole === "ADMIN" || userRole === "SUPER_ADMIN");
    // ส่ง total กลับเมื่ออ่านทั้งชีตจริง ๆ หรือเมื่อ full=true
    const totalToReturn = (readStartRow === 1 || full) ? totalMatches : null;
    return { success: true, history: historyList, total: totalToReturn, page: page, pageSize: pageSize, isGlobalView: isGlobalView, hasMore: hasMore };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/* เพิ่ม helper functions เพื่อประสิทธิภาพ I/O และ cache */
function _openSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/* อ่านข้อมูลชีตโดยจำกัดแถวจริง (รวม header ถ้ามี) */
function _readSheetAll(sheetName) {
  const ss = _openSS();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { sheet: null, values: [] };
  const lastRow = Math.max(1, sheet.getLastRow());
  const lastCol = Math.max(1, sheet.getLastColumn());
  if (lastRow === 1 && lastCol === 1 && !sheet.getRange(1,1).getValue()) return { sheet, values: [] };
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  return { sheet, values };
}

/* สร้างหรือดึง itemMap จาก CacheService (key -> itemName) */
function _getItemMapCached(ttlSeconds) {
  ttlSeconds = ttlSeconds || 300;
  const cache = CacheService.getScriptCache();
  const cacheKey = 'itemMap_v2';
  let itemMap = {};
  const cached = cache.get(cacheKey);
  if (cached) {
    try { itemMap = JSON.parse(cached); } catch (e) { itemMap = {}; }
  }
  if (!cached || Object.keys(itemMap).length === 0) {
    const res = _readSheetAll('Items');
    const itemData = res.values || [];
    for (let i = 1; i < itemData.length; i++) {
      if (!itemData[i] || !itemData[i][0]) continue;
      itemMap[String(itemData[i][0]).trim()] = itemData[i][1] || "";
    }
    try { cache.put(cacheKey, JSON.stringify(itemMap), ttlSeconds); } catch (e) { /* ignore cache errors */ }
  }
  return itemMap;
}

/* ฟอร์แมตวันที่ปลอดภัย (ลดการเรียก Utilities.formatDate กระจัดกระจาย) */
function _formatDateSafe(raw, tz, fmt) {
  if (!raw) return "-";
  const d = raw instanceof Date ? raw : new Date(raw);
  try { return Utilities.formatDate(d, tz, fmt); } catch (e) { return d.toLocaleString(); }

}

/* ปรับปรุง getDashboardData ให้อ่านเฉพาะแถวจริงและลดการประมวลผลซ้ำ */
function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const itemRes = _readSheetAll("Items");
    const transRes = _readSheetAll("Transactions");
    const userRes = _readSheetAll("Users");

    const itemData = itemRes.values || [];
    const transData = transRes.values || [];
    const userData = userRes.values || [];

    const items = [];
    let totalStock = 0;
    let borrowedCount = 0;

    for (let i = 1; i < itemData.length; i++) {
      const row = itemData[i];
      if (!row || !row[0]) continue;
      const qty = parseInt(row[3]) || 0;
      const status = row[5] || "";
      if (status !== "ชำรุด") totalStock += qty;
      items.push({
        id: row[0],
        name: row[1],
        type: row[2],
        quantity: qty,
        location: row[4],
        status: status,
        itemPic: row[6] || "",
        qr: row[7] || `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${row[0]}`
      });
    }

    const users = [];
    for (let k = 1; k < userData.length; k++) {
      if (!userData[k] || !userData[k][0]) continue;
      users.push({ username: userData[k][0], role: userData[k][1], password: userData[k][2] });
    }

    let pendingReturnCount = 0;
    let pendingApprovalCount = 0;
    for (let j = 1; j < transData.length; j++) {
      const row = transData[j];
      if (!row || !row[7]) continue;
      const status = row[7];
      if (status === "กำลังยืม") {
        pendingReturnCount++;
        borrowedCount += Number(row[9] || 1);
      } else if (status === "รออนุมัติ") {
        pendingApprovalCount++;
      }
    }

    return {
      summary: {
        total: items.filter(item => Number(item.quantity) > 0).length,
        available: totalStock,
        borrowed: borrowedCount,
        pendingReturn: pendingReturnCount,
        pendingApproval: pendingApprovalCount
      },
      items: items,
      users: users
    };
  } catch (error) {
    throw new Error(error.message);
  }
}

/* ปรับปรุง getExecutiveReportData ให้ใช้ itemMap ลดการค้นหาใหม่ ๆ */
function getExecutiveReportData(filterType) {
  try {
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transRes = _readSheetAll("Transactions");
    const itemRes = _readSheetAll("Items");
    const transData = transRes.values || [];
    const itemData = itemRes.values || [];

    const now = new Date();
    let totalBorrowInRange = 0, pendingReturnInRange = 0, returnSuccessInRange = 0;
    const locationStats = {}, topItemsStats = {}, itemInfoMap = {};
    let totalItemsCount = 0, availableCount = 0, borrowedCount = 0;

    for (let k = 1; k < itemData.length; k++) {
      const r = itemData[k];
      if (!r || !r[0]) continue;
      totalItemsCount++;
      if (r[5] === "พร้อมใช้งาน") availableCount++;
      if (r[5] === "ถูกยืม") borrowedCount++;
      itemInfoMap[r[0]] = { name: r[1], location: r[4] || "ไม่ระบุตำแหน่ง" };
    }

    for (let i = 1; i < transData.length; i++) {
      const row = transData[i];
      if (!row || !row[4]) continue;
      const borrowDate = new Date(row[4]);
      let isMatch = false;
      if (filterType === "WEEK") {
        const oneWeekAgo = new Date(); oneWeekAgo.setDate(now.getDate() - 7);
        if (borrowDate >= oneWeekAgo) isMatch = true;
      } else if (filterType === "MONTH") {
        if (borrowDate.getMonth() === now.getMonth() && borrowDate.getFullYear() === now.getFullYear()) isMatch = true;
      } else if (filterType === "YEAR") {
        if (borrowDate.getFullYear() === now.getFullYear()) isMatch = true;
      } else {
        isMatch = true;
      }

      if (isMatch) {
        totalBorrowInRange++;
        if (row[7] === "กำลังยืม") pendingReturnInRange++;
        if (row[7] === "คืนแล้ว") returnSuccessInRange++;
        const itemId = row[1];
        const itemInfo = itemInfoMap[itemId] || { name: "ไม่ทราบชื่อ", location: "ไม่ระบุตำแหน่ง" };
        topItemsStats[itemInfo.name] = (topItemsStats[itemInfo.name] || 0) + 1;
        locationStats[itemInfo.location] = (locationStats[itemInfo.location] || 0) + 1;
      }
    }

    return {
      success: true,
      kpi: {
        totalItems: totalItemsCount,
        availableItems: availableCount,
        borrowedItems: borrowedCount,
        rangeBorrows: totalBorrowInRange,
        rangePending: pendingReturnInRange,
        rangeReturned: returnSuccessInRange
      },
      charts: {
        locationLabels: Object.keys(locationStats),
        locationValues: Object.values(locationStats),
        itemLabels: Object.keys(topItemsStats),
        itemValues: Object.values(topItemsStats)
      }
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function saveItemData(mode, id, name, loc, stat, itemType, quantity, itemPicBase64) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Items");
    const data = sheet.getDataRange().getValues();
    const targetId = id.trim();
    const parsedQuantity = parseInt(quantity) || 0;

    let imageUrl = "";
    if (itemPicBase64 && itemPicBase64.includes("base64,")) {
      imageUrl = uploadFileToDrive(targetId, itemPicBase64, ITEM_FOLDER_ID);
    } else {
      imageUrl = itemPicBase64;
    }

    if (mode === "ADD") {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] && data[i][0].toString().trim() === targetId) return { success: false, message: "❌ รหัสพัสดุนี้มีอยู่แล้ว" };
      }
      const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${targetId}`;
      sheet.appendRow([targetId, name.trim(), itemType.trim(), parsedQuantity, loc.trim(), stat.trim(), imageUrl, qrUrl]);
      return { success: true, message: "🎉 เพิ่มข้อมูลพัสดุสำเร็จ!" };
    } else if (mode === "EDIT") {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] && data[i][0].toString().trim() === targetId) {
          sheet.getRange(i + 1, 2).setValue(name.trim());
          sheet.getRange(i + 1, 3).setValue(itemType.trim());
          sheet.getRange(i + 1, 4).setValue(parsedQuantity);
          sheet.getRange(i + 1, 5).setValue(loc.trim());
          sheet.getRange(i + 1, 6).setValue(stat.trim());
          if (imageUrl !== "") sheet.getRange(i + 1, 7).setValue(imageUrl);
          return { success: true, message: "🔒 อัปเดตข้อมูลพัสดุสำเร็จ!" };
        }
      }
    }
    return { success: false, message: "ไม่พบรหัสอุปกรณ์ที่ต้องการแก้ไข" };
  } catch (e) { return { success: false, message: "เกิดข้อผิดพลาด: " + e.toString() }; }
}

function deleteItemData(id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Items");
    const data = sheet.getDataRange().getValues();
    const targetId = id.trim();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === targetId) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "🗑️ ลบพัสดุครุภัณฑ์ออกจากคลังเรียบร้อยแล้ว" };
      }
    }
    return { success: false, message: "ไม่พบข้อมูลพัสดุที่ต้องการลบ" };
  } catch (e) { return { success: false, message: "เกิดข้อผิดพลาด: " + e.toString() }; }
}

function saveUserData(mode, user, role, pass) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Users");
    const data = sheet.getDataRange().getValues();
    const targetUser = user.trim().toLowerCase();
    if (mode === "ADD") {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] && data[i][0].toString().trim().toLowerCase() === targetUser) {
          return { success: false, message: "❌ Username นี้มีในระบบแล้ว" };
        }
      }
      sheet.appendRow([targetUser, role, pass.trim(), "", "", "", "", ""]);
      return { success: true, message: "🎉 เพิ่มบัญชีผู้ใช้งานใหม่สำเร็จ" };
    } else if (mode === "EDIT") {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] && data[i][0].toString().trim().toLowerCase() === targetUser) {
          sheet.getRange(i + 1, 2).setValue(role);
          if (pass && pass.trim() !== "") {
            sheet.getRange(i + 1, 3).setValue(pass.trim());
          }
          return { success: true, message: "🔒 ปรับปรุงระดับสิทธิ์ผู้ใช้เรียบร้อยแล้ว" };
        }
      }
    }
    return { success: false, message: "ไม่พบบัญชีผู้ใช้ที่ต้องการแก้ไข" };
  } catch (e) { return { success: false, message: "เกิดข้อผิดพลาด: " + e.toString() }; }
}

function deleteUserData(username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Users");
    const data = sheet.getDataRange().getValues();
    const targetUser = username.trim().toLowerCase();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toLowerCase() === targetUser) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "🗑️ ลบบัญชีผู้ใช้งานออกจากระบบเรียบร้อยแล้ว" };
      }
    }
    return { success: false, message: "ไม่พบบัญชีผู้ใช้ที่ต้องการลบ" };
  } catch (e) { return { success: false, message: "เกิดข้อผิดพลาด: " + e.toString() }; }
}

/* ปรับ getAdminEmails ให้กรองและอ่านเฉพาะแถวจริง */
function getAdminEmails() {
  try {
    const res = _readSheetAll('Users');
    const data = res.values || [];
    const adminEmails = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || !row[0]) continue;
      const role = String(row[1] || "").trim().toUpperCase();
      const email = row[4] || "";
      if (role === "ADMIN" && email) adminEmails.push(String(email).trim());
    }
    if (adminEmails.length === 0) return [ADMIN_EMAIL_DEFAULT];
    return adminEmails;
  } catch (err) {
    return [ADMIN_EMAIL_DEFAULT];
  }
}

/* ปรับปรุง borrowCartItems ให้อ่านแค่แถวจริงและตรวจสอบแบบ batch */
function borrowCartItems(cartItems, borrowerName, borrowerEmail, borrowDateStr, dueDateStr, purpose) {
  try {
    const ss = _openSS();
    const itemSheet = ss.getSheetByName("Items");
    const transSheet = ss.getSheetByName("Transactions");

    const lastItemRow = Math.max(1, itemSheet.getLastRow());
    const lastItemCol = Math.max(1, itemSheet.getLastColumn());
    const itemValues = lastItemRow >= 1 ? itemSheet.getRange(1, 1, lastItemRow, lastItemCol).getValues() : [];

    const nowTime = new Date();
    const bDate = new Date(borrowDateStr);
    bDate.setHours(nowTime.getHours(), nowTime.getMinutes(), nowTime.getSeconds());
    const dDate = new Date(dueDateStr);
    const transId = "TX-" + nowTime.getTime();

    if (Math.ceil(Math.abs(dDate - bDate) / (1000 * 60 * 60 * 24)) > 30) return { success: false, message: "❌ ไม่อนุญาตให้ยืมเกิน 30 วัน" };

    const itemMap = {};
    for (let i = 1; i < itemValues.length; i++) {
      if (!itemValues[i] || !itemValues[i][0]) continue;
      itemMap[String(itemValues[i][0]).trim()] = i;
    }

    for (let item of cartItems) {
      const idx = itemMap[item.id.trim()];
      if (idx === undefined || parseInt(itemValues[idx][3]) < item.qty) return { success: false, message: `❌ พัสดุ ${item.id} ไม่พอให้ยืม` };
    }

    // Resolve email และ username ให้ถูกต้อง
    let resolvedEmail = "";
    let resolvedUsername = "";

    // ถ้า borrowerEmail param เป็น email ให้ใช้เลย
    if (borrowerEmail && String(borrowerEmail).trim().indexOf('@') !== -1) {
      resolvedEmail = String(borrowerEmail).trim();
      // หา username จาก email ถาม Users
      resolvedUsername = findUsernameByIdentifier(resolvedEmail) || findUsernameByIdentifier(borrowerName) || "";
    } else {
      // borrowerEmail param อาจเป็น username หรือว่าง
      if (borrowerEmail && String(borrowerEmail).trim() !== "") {
        resolvedUsername = String(borrowerEmail).trim();
        resolvedEmail = findUserEmailByIdentifier(resolvedUsername) || getUserEmail(resolvedUsername) || "";
      }
      // ถ้ายังไม่มี email ให้ลองหาโดยใช้ borrowerName
      if (!resolvedEmail && borrowerName) {
        resolvedEmail = findUserEmailByIdentifier(borrowerName) || getUserEmail(borrowerName) || "";
        if (!resolvedUsername) resolvedUsername = findUsernameByIdentifier(borrowerName);
      }
    }

    // ถ้ายังไม่มี username ให้ใส่ borrowerName เป็นค่า fallback
    if (!resolvedUsername) resolvedUsername = borrowerName || "";

    const newTrans = [];
    let itemDetailsHtml = "";

    for (let item of cartItems) {
      itemDetailsHtml += `<li>รหัสพัสดุ: ${item.id} | ชื่อพัสดุ: ${item.name} | จำนวน: ${item.qty} ชิ้น</li>`;
      // คอลัมน์: A transId, B itemID, C borrowerName, D borrowerUsername (หรือชื่อที่แสดง), E borrowDate, F dueDate, G returnDate, H status, I purpose, J borrowQty, K borrowerEmail (จริง)
      newTrans.push([transId, item.id, borrowerName || "", resolvedUsername || "", bDate, dDate, "", "รออนุมัติ", purpose || "", item.qty, resolvedEmail || ""]);
    }

    if (newTrans.length > 0) {
      transSheet.getRange(transSheet.getLastRow() + 1, 1, newTrans.length, newTrans[0].length).setValues(newTrans);
    }

    const adminList = getAdminEmails();
    if (adminList.length > 0) {
      const emailSubject = `🔊 คำขอยืมพัสดุครุภัณฑ์ใหม่รอการพิจารณาอนุมัติ [ธุรกรรม: ${transId}]`;
      const emailBody = `<h3>ระบบยืม-คืนพัสดุอัจฉริยะ CPE มรพส.</h3>
        <p><b>ผู้ขอส่งคำยืม:</b> ${borrowerName} (${resolvedUsername}) ${resolvedEmail ? "(" + resolvedEmail + ")" : ""}</p>
        <p><b>วัตถุประสงค์:</b> ${purpose}</p>
        <p><b>รายการพัสดุที่ขอยืม:</b></p>
        <ul>${itemDetailsHtml}</ul>
        <p>โปรดตรวจสอบและพิจารณาคำขอผ่านระบบ ระบบยืม-คืนพัสดุอัจฉริยะ CPE มรพส.</p>`;

      for (let email of adminList) {
        try { GmailApp.sendEmail(email, emailSubject, "", { htmlBody: emailBody }); }
        catch (mailErr) { Logger.log("Admin Mail Send Warning: " + mailErr.toString()); }
      }
    }

    return { success: true, message: "🎉 ส่งคำขอยืมสำเร็จ! อยู่ระหว่างรอเจ้าหน้าที่ผู้ดูแลระบบพิจารณาอนุมัติ" };
  } catch (e) { return { success: false, message: "Error: " + e.toString() }; }
}

/**
 * helper: หาอีเมลจาก Users sheet โดยรับ identifier ที่อาจเป็น username หรือ fullName
 */
function findUserEmailByIdentifier(identifier) {
  try {
    if (!identifier) return null;
    const idRaw = String(identifier).trim();
    Logger.log("findUserEmailByIdentifier: resolving -> " + idRaw);
    const usersRes = _readSheetAll('Users');
    const users = usersRes.values || [];

    const norm = s => (String(s||'').trim().toLowerCase().replace(/\s+/g,' ')).replace(/[^\u0E00-\u0E7F0-9a-z\s]/g, '');
    const idNorm = norm(idRaw);
    const idNoSpace = idNorm.replace(/\s+/g,'');

    for (let r = 1; r < users.length; r++) {
      const row = users[r] || [];
      const username = String(row[0] || '').trim();
      const fullName = String(row[3] || '').trim();
      const email = String(row[4] || '').trim();
      if (!email) continue;
      const uNorm = norm(username);
      const fNorm = norm(fullName);

      if (uNorm === idNorm || fNorm === idNorm) {
        Logger.log("findUserEmailByIdentifier: exact match -> " + email + " (row " + (r+1) + ")");
        return email;
      }
      if (uNorm && uNorm.indexOf(idNorm) !== -1) { Logger.log("findUserEmailByIdentifier: username includes -> " + email); return email; }
      if (fNorm && fNorm.indexOf(idNorm) !== -1) { Logger.log("findUserEmailByIdentifier: fullname includes -> " + email); return email; }
      if (uNorm.replace(/\s+/g,'') === idNoSpace) { Logger.log("findUserEmailByIdentifier: username no-space match -> " + email); return email; }
      if (fNorm.replace(/\s+/g,'') === idNoSpace) { Logger.log("findUserEmailByIdentifier: fullname no-space match -> " + email); return email; }
      if (email.toLowerCase() === idNorm) { Logger.log("findUserEmailByIdentifier: identifier is email -> " + email); return email; }
    }

    Logger.log("findUserEmailByIdentifier: no match for -> " + idRaw);
  } catch (e) {
    Logger.log("findUserEmailByIdentifier error: " + e.toString());
  }
  return null;
}


/* แทนที่ส่วนหา/ส่งอีเมลใน approveSingleBorrowRequest ด้วยโค้ดนี้ */
function approveSingleBorrowRequest(transId, itemId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const transSheet = ss.getSheetByName("Transactions");
    const itemSheet = ss.getSheetByName("Items");
    const transData = transSheet.getDataRange().getValues();
    const itemData = itemSheet.getDataRange().getValues();
    
    let foundTransRow = -1;
    let qtyToBorrow = 1;
    let borrowerIdentifier = "";

    for (let j = 1; j < transData.length; j++) {
      if (String(transData[j][0]).trim() === String(transId).trim()
          && String(transData[j][1]).trim() === String(itemId).trim()
          && String(transData[j][7]).trim() === "รออนุมัติ") {
        foundTransRow = j + 1;
        qtyToBorrow = parseInt(transData[j][9]) || 1;
        borrowerIdentifier = String(transData[j][2] || "").trim();
        break;
      }
    }

    if (foundTransRow === -1) return { success: false, message: "❌ ไม่พบรายการคำขอยืมรออนุมัติที่ตรงกัน" };

    // ลดสต็อก
    let foundItemRow = -1;
    let currentStock = 0;
    for (let i = 1; i < itemData.length; i++) {
      if (String(itemData[i][0]).trim() === String(itemId).trim()) {
        foundItemRow = i + 1;
        currentStock = parseInt(itemData[i][3]) || 0;
        break;
      }
    }
    if (foundItemRow !== -1) {
      if (currentStock < qtyToBorrow) return { success: false, message: `❌ พัสดุในคลังเหลือ ${currentStock} ชิ้น ไม่พอต่อการขอยืม ${qtyToBorrow}` };
      const nextStock = currentStock - qtyToBorrow;
      itemSheet.getRange(foundItemRow, 4).setValue(nextStock);
      itemSheet.getRange(foundItemRow, 6).setValue(nextStock > 0 ? "พร้อมใช้งาน" : "ถูกยืม");
    }

    // อัปเดต transaction (วันที่ยืม, สถานะ)
    transSheet.getRange(foundTransRow, 5).setValue(new Date()); 
    transSheet.getRange(foundTransRow, 8).setValue("กำลังยืม");

    Logger.log("approveSingleBorrowRequest updated sheet -> transId:%s itemId:%s row:%s", transId, itemId, foundTransRow);

    // ไม่ส่งเมลแยกแต่ละชิ้น ตัดการส่งออกไป — เรียกสรุปผลรวมแทน (สรุปจะถูกส่งเมื่อไม่มีรายการรออนุมัติแล้ว)
    try { checkAndSendSummaryEmailToUser(transId); } catch(e){ Logger.log("checkAndSendSummaryEmailToUser err: " + e); }

    return { success: true, message: "✅ อนุมัติการยืมพัสดุชิ้นนี้เรียบร้อยแล้ว" };
  } catch (e) { return { success: false, message: e.toString() }; }
}


/* แทนที่ส่วนหา/ส่งอีเมลใน rejectSingleBorrowRequest ด้วยโค้ดนี้ */
function rejectSingleBorrowRequest(transId, itemId, reason) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const transSheet = ss.getSheetByName("Transactions");
    const transData = transSheet.getDataRange().getValues();

    let foundTransRow = -1;
    let borrowerIdentifier = "";

    for (let j = 1; j < transData.length; j++) {
      if (String(transData[j][0]).trim() === String(transId).trim()
          && String(transData[j][1]).trim() === String(itemId).trim()
          && String(transData[j][7]).trim() === "รออนุมัติ") {
        foundTransRow = j + 1;
        borrowerIdentifier = String(transData[j][2] || "").trim();
        break;
      }
    }

    if (foundTransRow === -1) return { success: false, message: "❌ ไม่พบรายการคำขอยืมรออนุมัติในระบบชีต" };

    transSheet.getRange(foundTransRow, 8).setValue("ไม่อนุมัติ");
    if (reason && reason.trim() !== "") {
      let currentPurpose = transSheet.getRange(foundTransRow, 9).getValue();
      transSheet.getRange(foundTransRow, 9).setValue(currentPurpose + " [เหตุผลปฏิเสธ: " + reason.trim() + "]");
    }

    Logger.log("rejectSingleBorrowRequest updated sheet -> transId:%s itemId:%s row:%s reason:%s", transId, itemId, foundTransRow, reason);

    // ไม่ส่งเมลแยกแต่ละชิ้น — เรียกสรุปผลรวมแทน
    try { checkAndSendSummaryEmailToUser(transId); } catch(e){ Logger.log("checkAndSendSummaryEmailToUser err: " + e); }

    return { success: true, message: "❌ ปฏิเสธคำขอยืมพัสดุชิ้นนี้เรียบร้อยแล้ว" };
  } catch (e) { return { success: false, message: e.toString() }; }
}


function returnSingleItem(transId, itemId, qty) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const itemSheet = ss.getSheetByName("Items");
    const transSheet = ss.getSheetByName("Transactions");
    const itemData = itemSheet.getDataRange().getValues();

    for (let i = 1; i < itemData.length; i++) {
      if (itemData[i][0].toString().trim() === itemId.trim()) {
        itemData[i][3] = parseInt(itemData[i][3]) + parseInt(qty);
        itemData[i][5] = "พร้อมใช้งาน";
        itemSheet.getRange(i + 1, 1, 1, itemData[0].length).setValues([itemData[i]]);
        break;
      }
    }

    const transData = transSheet.getDataRange().getValues();
    for (let j = 1; j < transData.length; j++) {
      if (transData[j][0].toString().trim() === transId.trim() && transData[j][1].toString().trim() === itemId.trim() && transData[j][7] === "กำลังยืม") {
        transSheet.getRange(j + 1, 7).setValue(new Date()); 
        transSheet.getRange(j + 1, 8).setValue("คืนแล้ว");
        return { success: true, message: "📥 ส่งคืนพัสดุ " + itemId + " เข้าคลังเรียบร้อยแล้ว" };
      }
    }
    return { success: false, message: "❌ ไม่พบรายการธุรกรรมที่ตรงกับพัสดุชิ้นนี้" };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function _resolveBorrowerEmail(txDataRow) {
  try {
    if (!txDataRow) return null;

    // ตรวจทุก key ใน obj เพื่อหา string ที่มี '@'
    for (let k in txDataRow) {
      try {
        const v = txDataRow[k];
        if (v && String(v).trim().indexOf('@') !== -1) {
          return String(v).trim();
        }
      } catch (e) { /* ignore */ }
    }

    // ถ้ายังไม่มีลอง keys แบบเดิมและ lookup จาก Users
    const rawEmails = [
      txDataRow['borrowerEmail'],
      txDataRow['email'],
      txDataRow['ผู้ยืมอีเมล'],
      txDataRow['borrower']
    ];
    for (let v of rawEmails) {
      if (v && String(v).trim() !== "") {
        const s = String(v).trim();
        if (s.indexOf('@') !== -1) return s;
        const resolved = getUserEmail(s);
        if (resolved) return resolved;
      }
    }

    const borrowerName = txDataRow['borrowerName'] || txDataRow['ผู้ยืมชื่อ'] || "";
    if (borrowerName && String(borrowerName).trim() !== "") {
      const resolved = getUserEmail(String(borrowerName).trim());
      if (resolved) return resolved;
    }

  } catch (e) {
    Logger.log("resolveBorrowerEmail error: " + e.toString());
  }
  return null;
}

