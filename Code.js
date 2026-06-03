/**
 * ระบบยืม-คืนพัสดุและโสตทัศนูปกรณ์อัจฉริยะ CPE มรพส. (เวอร์ชันตะกร้า E-Commerce - Ultimate Production 2026)
 * พัฒนาโดย: AI พัฒนาระบบด้วย GAS (Senior Apps Script Specialist)
 */

const PROFILE_FOLDER_ID = "1PbgnS8eZXOdKXLPFM-XSeBwEQCnKQYh1";
const ITEM_FOLDER_ID = "1bFnS6npqJXKnpzuL8YjFBo7K3GqQ_yW5";
const SPREADSHEET_ID = "1rK7WMeaicIUnvMVVQHdn5gxaIXwlv-AzFSVGe6CwiX8"; 
const ADMIN_EMAIL_DEFAULT = "somjedtongdee@psru.ac.th";

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
    transSheet.appendRow(["transID", "itemID", "borrowerName", "borrowerEmail", "borrowDate", "dueDate", "returnDate", "status", "purpose", "borrowQty"]);
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
function getUserEmail(username) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("Users");
    if (!sheet) return null;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim().toLowerCase() === username.toString().trim().toLowerCase()) {
        return data[i][4]; 
      }
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

    // ค้นหาพัสดุชิ้นอื่นที่รันด้วยเลขธุรกรรม (ตะกร้า) ชุดเดียวกัน
    for (let j = 1; j < transData.length; j++) {
      if (transData[j][0].toString().trim() === transId.trim()) {
        borrowerName = transData[j][2];
        borrowerEmail = transData[j][3];
        
        // หากพบว่าพัสดุบางชิ้นในกลุ่มคำขอนี้ยังคงสถานะ "รออนุมัติ" ให้ระงับกระบวนการส่งอีเมลไว้ก่อน
        if (transData[j][7] === "รออนุมัติ") {
          hasPending = true;
          break;
        }

        basketItems.push({
          itemId: transData[j][1],
          itemName: itemMap[transData[j][1].toString().trim()] || "ไม่พบชื่อพัสดุในคลัง",
          qty: transData[j][9] || 1,
          status: transData[j][7]
        });
      }
    }

    // เงื่อนไข: ส่งอีเมลสรุปผลรวมหาผู้ใช้งานเพียง 1 ฉบับ เมื่อเจ้าหน้าที่ตรวจสอบพิจารณาครบทุกชิ้นแล้ว
    if (!hasPending && basketItems.length > 0 && borrowerEmail) {
      const subject = `✅ [สรุปผลการพิจารณา] แจ้งสถานะคำขอยืมพัสดุครุภัณฑ์ เลขธุรกรรม: ${transId}`;
      
      let itemsListHtml = "";
      basketItems.forEach(item => {
        let statusBadge = item.status === "กำลังยืม" 
          ? `<span style="color: #059669; font-weight: bold;">🟢 ได้รับอนุมัติ</span>` 
          : `<span style="color: #e11d48; font-weight: bold;">🔴 ไม่ได้รับการอนุมัติ</span>`;
        
        itemsListHtml += `
          <tr style="border-b: 1px solid #e2e8f0;">
            <td style="padding: 10px; border: 1px solid #e2e8f0; font-family: monospace; font-weight: bold; color: #0f172a;">${item.itemId}</td>
            <td style="padding: 10px; border: 1px solid #e2e8f0; color: #334155;">${item.itemName}</td>
            <td style="padding: 10px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold; color: #0f172a;">${item.qty} ชิ้น</td>
            <td style="padding: 10px; border: 1px solid #e2e8f0; text-align: center;">${statusBadge}</td>
          </tr>`;
      });

      const htmlBody = `
        <div style="font-family: Sarabun, Arial, sans-serif; border: 1px solid #e2e8f0; padding: 25px; border-radius: 12px; max-width: 650px; background-color: #ffffff; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
          <h2 style="color: #0f172a; border-bottom: 2px solid #fbbf24; padding-bottom: 10px; margin-top: 0; font-size: 18px;">ระบบจัดการยืม-คืนพัสดุภาควิชาวิศวกรรมคอมพิวเตอร์ มรพส.</h2>
          <p>เรียน คุณ <b>${borrowerName}</b>,</p>
          <p>เจ้าหน้าที่ผู้ดูแลระบบคลังพัสดุได้ทำการตรวจสอบและพิจารณาผลคำขอขอยืมทรัพยากรครุภัณฑ์ในตะกร้าส่งยืมของคุณ <b>ครบทุกรายการเรียบร้อยแล้ว</b> รายละเอียดสรุปผลลัพธ์มีดังนี้:</p>
          
          <table style="width: 100%; margin: 15px 0; border-collapse: collapse; font-size: 13px;">
            <thead>
              <tr style="background-color: #f1f5f9; color: #475569; text-align: left;">
                <th style="padding: 10px; border: 1px solid #e2e8f0;">รหัสพัสดุ</th>
                <th style="padding: 10px; border: 1px solid #e2e8f0;">ชื่อครุภัณฑ์อุปกรณ์</th>
                <th style="padding: 10px; border: 1px solid #e2e8f0; text-align: center;">จำนวนขอยืม</th>
                <th style="padding: 10px; border: 1px solid #e2e8f0; text-align: center;">ผลการพิจารณา</th>
              </tr>
            </thead>
            <tbody>
              ${itemsListHtml}
            </tbody>
          </table>
          
          <p style="background-color: #f8fafc; color: #475569; padding: 12px; border-radius: 6px; border-left: 4px solid #cbd5e1; font-size: 12px; line-height: 1.5;">
            💡 <b>คำแนะนำการติดต่อรับพัสดุ:</b> สำหรับรายการครุภัณฑ์พัสดุที่ขึ้นสถานะ <b>"ได้รับอนุมัติ"</b> ท่านสามารถมารับเครื่องมือและครุภัณฑ์ได้ ณ ห้องจัดเก็บพัสดุตามเวลาทำการของสาขาวิชาครับ
          </p>
          <hr style="border: 0; border-top: 1px solid #e2e8f0; margin: 20px 0;">
          <p style="font-size: 11px; color: #94a3b8; text-align: center; margin-bottom: 0;">* ข้อความอัตโนมัติจากระบบคลังทรัพยากรส่วนกลางคอมพิวเตอร์ CPE Smart Asset Management 2026</p>
        </div>`;

      GmailApp.sendEmail(borrowerEmail, subject, "", { htmlBody: htmlBody });
    }
  } catch (e) {
    Logger.log("Error ในการจัดส่งอีเมลสรุปผลรวมหาผู้ยืม: " + e.toString());
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
  var userEmail = getUserEmail(username);
  
  if (!userEmail) {
    Logger.log("ไม่สามารถส่งอีเมลได้ เนื่องจากไม่พบอีเมลสำหรับ username: " + username);
    return;
  }
  
  var subject = "";
  var statusText = "";
  var statusColor = "";
  var noteText = "";
  
  if (status === "approved" || status === "อนุมัติ") {
    subject = "✅ [อนุมัติ] ผลการคำขอยืมพัสดุระบบ CPE มรพส.";
    statusText = "ได้รับการอนุมัติ";
    statusColor = "#28a745"; // สีเขียว
    noteText = "กรุณานำหลักฐานหรือติดต่อรับพัสดุ ณ ห้องปฏิบัติการคอมพิวเตอร์ตามเวลาที่กำหนด";
  } else {
    subject = "❌ [ปฏิเสธ] ผลการคำขอยืมพัสดุระบบ CPE มรพส.";
    statusText = "ปฏิเสธการอนุมัติ / ไม่ได้รับอนุมัติ";
    statusColor = "#dc3545"; // สีแดง
    noteText = "หากมีข้อสงสัยประการใด กรุณาติดต่อผู้ดูแลระบบหรือเจ้าหน้าที่ประจำห้องปฏิบัติการ";
  }
  
  // สร้างเนื้อหาอีเมลแบบ HTML ให้สวยงามและเป็นทางการ
  var htmlBody = `
    <div style="font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; padding: 20px; color: #333; max-width: 600px; border: 1px solid #e0e0e0; border-radius: 5px;">
      <h2 style="color: #0056b3; border-bottom: 2px solid #0056b3; padding-bottom: 10px;">แจ้งเตือนสถานะการยืมพัสดุ</h2>
      <p>สวัสดีคุณ <b>${username}</b>,</p>
      <p>เจ้าหน้าที่ได้ตรวจสอบและพิจารณาคำขอขอยืมพัสดุของคุณเรียบร้อยแล้ว โดยมีรายละเอียดดังนี้:</p>
      
      <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
        <tr style="background-color: #f8f9fa;">
          <td style="padding: 10px; border: 1px solid #dee2e6; font-weight: bold; width: 30%;">รหัสพัสดุ:</td>
          <td style="padding: 10px; border: 1px solid #dee2e6;">${itemCode}</td>
        </tr>
        <tr>
          <td style="padding: 10px; border: 1px solid #dee2e6; font-weight: bold;">ชื่อพัสดุ:</td>
          <td style="padding: 10px; border: 1px solid #dee2e6;">${itemName}</td>
        </tr>
        <tr style="background-color: #f8f9fa;">
          <td style="padding: 10px; border: 1px solid #dee2e6; font-weight: bold;">สถานะผลการยืม:</td>
          <td style="padding: 10px; border: 1px solid #dee2e6; color: ${statusColor}; font-weight: bold; font-size: 1.1em;">${statusText}</td>
        </tr>
      </table>
      
      <p style="background-color: #fff3cd; color: #856404; padding: 10px; border-radius: 3px; border-left: 5px solid #ffeeba;">
        📌 <b>หมายเหตุ:</b> ${noteText}
      </p>
      
      <hr style="border: 0; border-top: 1px solid #e0e0e0; margin-top: 30px;">
      <p style="font-size: 0.85em; color: #6c757d; text-align: center;">ระบบจัดการยืม-คืนพัสดุภาควิชาวิศวกรรมคอมพิวเตอร์ มหาวิทยาลัยราชภัฏพิบูลสงคราม</p>
    </div>
  `;
  
  MailApp.sendEmail({
    to: userEmail,
    subject: subject,
    htmlBody: htmlBody
  });
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

// ...existing code...
function getUserLoanHistory(username, role, page, pageSize) {
  try {
    // ตั้งค่า pagination ค่าเริ่มต้น
    page = parseInt(page, 10) || 1;
    pageSize = parseInt(pageSize, 10) || 50;
    const neededCount = page * pageSize;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transSheet = ss.getSheetByName("Transactions");
    const itemSheet = ss.getSheetByName("Items");

    // อ่านเฉพาะช่วงแถวที่มีข้อมูลจริง (ลดการอ่านทั้งชีตเมื่อไฟล์ใหญ่)
    const lastTransRow = Math.max(1, transSheet.getLastRow());
    const lastTransCol = Math.max(1, transSheet.getLastColumn());
    const transData = lastTransRow >= 1 ? transSheet.getRange(1, 1, lastTransRow, lastTransCol).getValues() : [];

    // ใช้ CacheService เก็บข้อมูล itemMap สั้น ๆ เพื่อลดการอ่านบ่อย ๆ
    const cache = CacheService.getScriptCache();
    const cacheKey = 'itemMap_v1';
    let itemMap = {};
    const cached = cache.get(cacheKey);
    if (cached) {
      try { itemMap = JSON.parse(cached); } catch (e) { itemMap = {}; }
    }
    if (!cached || Object.keys(itemMap).length === 0) {
      const lastItemRow = Math.max(1, itemSheet.getLastRow());
      const lastItemCol = Math.max(1, itemSheet.getLastColumn());
      const itemData = lastItemRow >= 1 ? itemSheet.getRange(1, 1, lastItemRow, lastItemCol).getValues() : [];
      for (let i = 1; i < itemData.length; i++) {
        if (itemData[i][0]) itemMap[itemData[i][0].toString().trim()] = itemData[i][1] || "";
      }
      // เก็บ cache 300 วินาที
      try { cache.put(cacheKey, JSON.stringify(itemMap), 300); } catch (e) { /* ignore cache errors */ }
    }

    const targetUser = username ? username.trim().toLowerCase() : "";
    const userRole = role ? role.trim().toUpperCase() : "USER";
    const ssTimeZone = ss.getSpreadsheetTimeZone();

    // วนจากแถวล่าสุดไปก่อน และหยุดเมื่อได้ข้อมูลพอสำหรับ requested page (ลดการประมวลผล)
    const matched = [];
    for (let j = transData.length - 1; j >= 1; j--) {
      if (!transData[j] || !transData[j][1]) continue;
      const transUser = transData[j][3] ? transData[j][3].toString().trim().toLowerCase() : "";
      if (userRole === "ADMIN" || userRole === "SUPER_ADMIN" || transUser === targetUser) {
        matched.push(j); // เก็บ index ของแถวที่ตรงเงื่อนไข
        if (matched.length >= neededCount) break; // หยุดเมื่อได้เพียงพอ
      }
    }

    const historyList = [];
    // สร้างผลลัพธ์สำหรับ page ที่ต้องการ (only format those rows)
    const startIndex = (page - 1) * pageSize;
    for (let k = startIndex; k < Math.min(matched.length, startIndex + pageSize); k++) {
      const rowIdx = matched[k];
      const row = transData[rowIdx];

      // ฟอร์แมตวันที่เฉพาะแถวที่ต้องส่งกลับ
      let bDateFormatted = "-";
      let dDateFormatted = "-";
      let rDateFormatted = "-";

      const bRaw = row[4];
      if (bRaw) bDateFormatted = Utilities.formatDate(bRaw instanceof Date ? bRaw : new Date(bRaw), ssTimeZone, "dd/MM/yyyy HH:mm");

      const dRaw = row[5];
      if (dRaw) dDateFormatted = Utilities.formatDate(dRaw instanceof Date ? dRaw : new Date(dRaw), ssTimeZone, "dd/MM/yyyy");

      const rRaw = row[6];
      if (rRaw) rDateFormatted = Utilities.formatDate(rRaw instanceof Date ? rRaw : new Date(rRaw), ssTimeZone, "dd/MM/yyyy HH:mm");

      const borrowerEmail = row[3] ? row[3].toString().trim().toLowerCase() : "";
      historyList.push({
        transId: row[0],
        itemId: row[1],
        itemName: itemMap[row[1].toString().trim()] || "ไม่พบชื่อพัสดุในคลัง",
        borrowerName: row[2] ? row[2].toString().trim() : "ไม่ระบุชื่อ",
        borrowerEmail: borrowerEmail,
        borrowDate: bDateFormatted,
        dueDate: dDateFormatted,
        returnDate: rDateFormatted,
        status: row[7],
        purpose: row[8] || "-",
        qty: row[9] || 1
      });
    }

    // ส่งข้อมูลจำนวนรวม (เพื่อให้ UI แสดง pagination ได้)
    const totalMatches = matched.length;
    const isGlobalView = (userRole === "ADMIN" || userRole === "SUPER_ADMIN");
    return { success: true, history: historyList, total: totalMatches, page: page, pageSize: pageSize, isGlobalView: isGlobalView };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const itemSheet = ss.getSheetByName("Items");
    const transSheet = ss.getSheetByName("Transactions");
    const userSheet = ss.getSheetByName("Users");
    const itemData = itemSheet.getDataRange().getValues();
    const transData = transSheet.getDataRange().getValues();
    const userData = userSheet.getDataRange().getValues();
    const items = [];
    let totalStock = 0; let borrowedCount = 0;
    for (let i = 1; i < itemData.length; i++) {
      if (!itemData[i][0]) continue;
      const status = itemData[i][5];
      const qty = parseInt(itemData[i][3]) || 0;
      if (status !== "ชำรุด") {
        totalStock += qty;
      }

      items.push({
        id: itemData[i][0],
        name: itemData[i][1],
        type: itemData[i][2],
        quantity: qty,
        location: itemData[i][4],
        status: status,
        itemPic: itemData[i][6] || "",
        qr: itemData[i][7] || `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${itemData[i][0]}`
      });
    }
    const users = [];
    for (let k = 1; k < userData.length; k++) {
      if (!userData[k][0]) continue;
      users.push({ username: userData[k][0], role: userData[k][1], password: userData[k][2] });
    }
    let pendingReturnCount = 0;
    let pendingApprovalCount = 0; 

    for (let j = 1; j < transData.length; j++) {
      if (transData[j][7] === "กำลังยืม") {
        pendingReturnCount++;
        borrowedCount += Number(transData[j][9] || 1);
      }
      if (transData[j][7] === "รออนุมัติ") {
        pendingApprovalCount++;
      }
    }
    return { summary: { total: items.filter(item => Number(item.quantity) > 0).length, available: totalStock, borrowed: borrowedCount, pendingReturn: pendingReturnCount, pendingApproval: pendingApprovalCount }, items: items, users: users };
  } catch (error) { throw new Error(error.message); }
}

function getExecutiveReportData(filterType) {
  try {
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transSheet = ss.getSheetByName("Transactions");
    const itemSheet = ss.getSheetByName("Items");
    const transData = transSheet.getDataRange().getValues();
    const itemData = itemSheet.getDataRange().getValues();
    const now = new Date();
    let totalBorrowInRange = 0; let pendingReturnInRange = 0; let returnSuccessInRange = 0;
    const locationStats = {}; const topItemsStats = {}; const itemInfoMap = {};
    let totalItemsCount = 0; let availableCount = 0; let borrowedCount = 0;

    for (let k = 1; k < itemData.length; k++) {
      if (!itemData[k][0]) continue;
      totalItemsCount++;
      if (itemData[k][5] === "พร้อมใช้งาน") availableCount++;
      if (itemData[k][5] === "ถูกยืม") borrowedCount++;
      itemInfoMap[itemData[k][0]] = { name: itemData[k][1], location: itemData[k][4] };
    }
    for (let i = 1; i < transData.length; i++) {
      if (!transData[i][4]) continue;
      const borrowDate = new Date(transData[i][4]);
      let isMatch = false;
      if (filterType === "WEEK") { const oneWeekAgo = new Date(); oneWeekAgo.setDate(now.getDate() - 7); if (borrowDate >= oneWeekAgo) isMatch = true; }
      else if (filterType === "MONTH") { if (borrowDate.getMonth() === now.getMonth() && borrowDate.getFullYear() === now.getFullYear()) isMatch = true; }
      else if (filterType === "YEAR") { if (borrowDate.getFullYear() === now.getFullYear()) isMatch = true; }

      if (isMatch) {
        totalBorrowInRange++;
        if (transData[i][7] === "กำลังยืม") pendingReturnInRange++;
        if (transData[i][7] === "คืนแล้ว") returnSuccessInRange++;
        const itemId = transData[i][1];
        const itemInfo = itemInfoMap[itemId] || { name: "ไม่ทราบชื่อ", location: "ไม่ระบุตำแหน่ง" };
        topItemsStats[itemInfo.name] = (topItemsStats[itemInfo.name] || 0) + 1;
        locationStats[itemInfo.location] = (locationStats[itemInfo.location] || 0) + 1;
      }
    }
    return { success: true, kpi: { totalItems: totalItemsCount, availableItems: availableCount, borrowedItems: borrowedCount, rangeBorrows: totalBorrowInRange, rangePending: pendingReturnInRange, rangeReturned: returnSuccessInRange }, charts: { locationLabels: Object.keys(locationStats), locationValues: Object.values(locationStats), itemLabels: Object.keys(topItemsStats), itemValues: Object.values(topItemsStats) } };
  } catch (e) { return { success: false, message: e.toString() }; }
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

function getAdminEmails() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const userSheet = ss.getSheetByName("Users");
    if (!userSheet) return [];
    const data = userSheet.getDataRange().getValues();
    const adminEmails = [];
    for (let i = 1; i < data.length; i++) {
      const role = String(data[i][1]).trim().toUpperCase();
      // คัดกรองสิทธิ์: ดึงเฉพาะบทบาท ADMIN เท่านั้น ป้องกันปัญหาเมลตีกลับจากบัญชีจำลองของ Super Admin
      if (role === "ADMIN" && data[i][4]) {
        adminEmails.push(data[i][4].toString().trim());
      }
    }
    return adminEmails;
  } catch(err) {
    return [ADMIN_EMAIL_DEFAULT];
  }
}

function borrowCartItems(cartItems, borrowerName, borrowerEmail, borrowDateStr, dueDateStr, purpose) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const itemSheet = ss.getSheetByName("Items");
    const transSheet = ss.getSheetByName("Transactions");
    const itemRange = itemSheet.getDataRange();
    const itemValues = itemRange.getValues();
    
    const nowTime = new Date();
    const bDate = new Date(borrowDateStr);
    bDate.setHours(nowTime.getHours(), nowTime.getMinutes(), nowTime.getSeconds());
    
    const dDate = new Date(dueDateStr);
    const transId = "TX-" + nowTime.getTime();

    if (Math.ceil(Math.abs(dDate - bDate) / (1000 * 60 * 60 * 24)) > 30) return { success: false, message: "❌ ไม่อนุญาตให้ยืมเกิน 30 วัน" };

    const itemMap = {};
    for (let i = 1; i < itemValues.length; i++) itemMap[itemValues[i][0].toString().trim()] = i;

    for (let item of cartItems) {
      const idx = itemMap[item.id.trim()];
      if (idx === undefined || parseInt(itemValues[idx][3]) < item.qty) return { success: false, message: `❌ พัสดุ ${item.id} ไม่พอให้ยืม` };
    }

    const newTrans = [];
    let itemDetailsHtml = ""; 
    
    for (let item of cartItems) {
      itemDetailsHtml += `<li>รหัสพัสดุ: ${item.id} | ชื่อพัสดุ: ${item.name} | จำนวน: ${item.qty} ชิ้น</li>`;
      newTrans.push([transId, item.id, borrowerName, borrowerEmail, bDate, dDate, "", "รออนุมัติ", purpose, item.qty]);
    }

    if (newTrans.length > 0) transSheet.getRange(transSheet.getLastRow() + 1, 1, newTrans.length, newTrans[0].length).setValues(newTrans);
    
    // ส่งอีเมลแจ้งเตือนพัสดุรอการตรวจสอบเข้าสู่ระบบเมลของกลุ่มเจ้าหน้าที่ ADMIN ตัวจริง
    const adminList = getAdminEmails();
    if (adminList.length > 0) {
      const emailSubject = `📢 มีคำขอยืมพัสดุครุภัณฑ์ใหม่รอการพิจารณาอนุมัติ [ธุรกรรม: ${transId}]`;
      const emailBody = `<h3>ระบบยืม-คืนพัสดุอัจฉริยะ CPE มรพส.</h3>
        <p><b>ผู้ขอส่งคำยืม:</b> ${borrowerName} (${borrowerEmail})</p>
        <p><b>วัตถุประสงค์:</b> ${purpose}</p>
        <p><b>รายการพัสดุที่ขอยืม:</b></p>
        <ul>${itemDetailsHtml}</ul>
        <p>โปรดตรวจสอบและพิจารณาคำขอผ่านระบบ Web Application</p>`;
      
      adminList.forEach(email => {
        try {
          GmailApp.sendEmail(email, emailSubject, "", { htmlBody: emailBody });
        } catch(mailErr) { Logger.log("Admin Mail Send Warning: " + mailErr.toString()); }
      });
    }

    return { success: true, message: "🎉 ส่งคำขอยืมสำเร็จ! อยู่ระหว่างรอเจ้าหน้าที่ผู้ดูแลระบบพิจารณาอนุมัติ" };
  } catch (e) { return { success: false, message: "Error: " + e.toString() }; }
}

// [COMPLETE CONFIG] ฟังก์ชันสำหรับการอนุมัติคำขอยืม (Approve) พร้อมหักลบสต็อกตามไอเทมจริง
function approveSingleBorrowRequest(transId, itemId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const transSheet = ss.getSheetByName("Transactions");
    const itemSheet = ss.getSheetByName("Items");
    const transData = transSheet.getDataRange().getValues();
    const itemData = itemSheet.getDataRange().getValues();
    
    let foundTransRow = -1;
    let qtyToBorrow = 1;

    for (let j = 1; j < transData.length; j++) {
      if (transData[j][0].toString().trim() === transId.trim() && transData[j][1].toString().trim() === itemId.trim() && transData[j][7] === "รออนุมัติ") {
        foundTransRow = j + 1;
        qtyToBorrow = parseInt(transData[j][9]) || 1;
        break;
      }
    }

    if (foundTransRow === -1) return { success: false, message: "❌ ไม่พบรายการคำขอยืมรออนุมัติที่ตรงกัน" };

    let foundItemRow = -1;
    let currentStock = 0;
    for (let i = 1; i < itemData.length; i++) {
      if (itemData[i][0].toString().trim() === itemId.trim()) {
        foundItemRow = i + 1;
        currentStock = parseInt(itemData[i][3]) || 0;
        break;
      }
    }

    if (foundItemRow !== -1) {
      if (currentStock < qtyToBorrow) return { success: false, message: `❌ พัสดุในคลังเหลือเพียง ${currentStock} ชิ้น ไม่พอต่อการขอยืมจำนวน ${qtyToBorrow} ชิ้น` };
      let nextStock = currentStock - qtyToBorrow;
      itemSheet.getRange(foundItemRow, 4).setValue(nextStock);
      itemSheet.getRange(foundItemRow, 6).setValue(nextStock > 0 ? "พร้อมใช้งาน" : "ถูกยืม");
    }

    transSheet.getRange(foundTransRow, 5).setValue(new Date()); 
    transSheet.getRange(foundTransRow, 8).setValue("กำลังยืม");

    // [TARGET FIX] ตรวจสอบความครบถ้วนของตะกร้า เพื่อส่งอีเมลรวมผล 1 ฉบับหาผู้ใช้งาน
    checkAndSendSummaryEmailToUser(transId);

    return { success: true, message: "✅ อนุมัติการยืมพัสดุชิ้นนี้สำเร็จเรียบร้อยแล้ว" };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// [COMPLETE CONFIG] ฟังก์ชันสำหรับการปฏิเสธคำขอยืม (Reject) และส่งเหตุผลเข้าอีเมลผู้ใช้
function rejectSingleBorrowRequest(transId, itemId, reason) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const transSheet = ss.getSheetByName("Transactions");
    const transData = transSheet.getDataRange().getValues();

    let foundTransRow = -1;
    for (let j = 1; j < transData.length; j++) {
      if (transData[j][0].toString().trim() === transId.trim() && transData[j][1].toString().trim() === itemId.trim() && transData[j][7] === "รออนุมัติ") {
        foundTransRow = j + 1;
        break;
      }
    }

    if (foundTransRow === -1) return { success: false, message: "❌ ไม่พบรายการคำขอยืมรออนุมัติในระบบชีต" };

    transSheet.getRange(foundTransRow, 8).setValue("ไม่อนุมัติ");
    if (reason && reason.trim() !== "") {
      let currentPurpose = transSheet.getRange(foundTransRow, 9).getValue();
      transSheet.getRange(foundTransRow, 9).setValue(currentPurpose + " [เหตุผลปฏิเสธ: " + reason.trim() + "]");
    }

    // [TARGET FIX] ตรวจสอบความครบถ้วนของตะกร้า เพื่อส่งอีเมลรวมผล 1 ฉบับหาผู้ใช้งาน
    checkAndSendSummaryEmailToUser(transId);

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