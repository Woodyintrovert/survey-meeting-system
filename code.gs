const SPREADSHEET_ID = "1FGF30Mo8oXer12AsAVzgqM3-4wKgQCnANbkPXnN28TY";

/* ===============================
   ENTRY POINT (รองรับหลายหน้า)
================================ */
function doGet(e) {
  const page = e?.parameter?.page || "index";
  const t = HtmlService.createTemplateFromFile(page);

  // --- ปรับปรุงตรงนี้: ส่งค่า mid แยกออกมาต่างหาก ---
  t.mid = e.parameter.mid || ""; 
  t.params = e.parameter || {};
  // ---------------------------------------------

  return t.evaluate()
    .setTitle("Survey System")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ===============================
   MASTER: รายชื่อกรรมการ
================================ */
function getCommitteeList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ชีต1"); // 👈 มั่นใจว่าชื่อ "ชีต1" จริงๆ ไม่ใช่ "Sheet1"
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues().slice(1);
  return data.map(r => r[0]).filter(Boolean);
}

/* ===============================
   CREATE MEETING (หัวใจระบบ)
================================ */
function createMeeting(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName("ชีต3");
  if (!sh) throw new Error("ไม่พบชีต3");

  const meetingId = Utilities.getUuid();

  data.members.forEach(name => {
    sh.appendRow([
      meetingId,                 // A meetingId
      data.meetingName,           // B
      new Date(data.meetingDate), // C
      data.meetingTime,           // D
      data.meetingPlace,          // E
      data.surveyBy,              // F
      name,                       // G
      "ACTIVE",                    // H
      "OPEN"                      // I 👈 เพิ่มตรงนี้เพื่อให้ครั้งแรกไม่ว่าง
    ]);
  });

  return ScriptApp.getService().getUrl()
    + "?page=survey_action&mid=" + meetingId;
}

/* ===============================
   LOAD MEETING BY ID
================================ */
function getMeetingById(meetingId) {
  console.log("กำลังค้นหา ID: " + meetingId);
  if (!meetingId) return null;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. ดึงข้อมูลรายชื่อจาก "ชีต3"
  const sh3 = ss.getSheetByName("ชีต3");
  if (!sh3) throw new Error("ไม่พบชีต3");
  const data3 = sh3.getDataRange().getValues().slice(1);
  const mid = String(meetingId).trim();
  
  // กรองแถวที่มี meetingId ตรงกันจากชีต3
  const rows = data3.filter(r => String(r[0]).trim() === mid);

  if (rows.length === 0) {
    console.log("❌ ไม่พบข้อมูลในระบบ (ชีต3)");
    return null;
  }

  // 2. ดึงข้อมูลการสำรวจทั้งหมดจาก "ชีต2"
  const sh2 = ss.getSheetByName("ชีต2");
  let allSurveys = [];
  if (sh2) {
    // ดึงมาทั้งหมดเพื่อหาค่าล่าสุด
    allSurveys = sh2.getDataRange().getValues();
  }

  // 3. Mapping รายชื่อกรรมการ พร้อมดึง "สถานะล่าสุด"
  const mappedMembers = rows.map(r => {
    const memberName = String(r[6]).trim(); // ชีต3 คอลัมน์ 7 (index 6)
    if (!memberName) return null;

    // --- Logic ใหม่: ค้นหาข้อมูลล่าสุดจาก ชีต2 โดยวนจากล่างขึ้นบน ---
    let saved = null;
    for (let i = allSurveys.length - 1; i >= 0; i--) {
      const sMid = String(allSurveys[i][1]).trim(); // ชีต2 คอลัมน์ 2 (index 1)
      const sName = String(allSurveys[i][3]).trim(); // ชีต2 คอลัมน์ 4 (index 3)
      
      if (sMid === mid && sName === memberName) {
        saved = allSurveys[i];
        break; // เจออันล่างสุด (ล่าสุด) แล้วให้หยุดหาทันที
      }
    }

    if (saved) {
      // คืนค่าสถานะล่าสุด (ตรวจสอบค่า "TRUE" หรือ true จากการบันทึก)
      return {
        name: memberName,
        onsite: String(saved[4]).toUpperCase() === "TRUE", // คอลัมน์ 5 (index 4)
        online: String(saved[5]).toUpperCase() === "TRUE", // คอลัมน์ 6 (index 5)
        leave:  String(saved[6]).toUpperCase() === "TRUE", // คอลัมน์ 7 (index 6)
        food:   String(saved[7]).toUpperCase() === "TRUE", // คอลัมน์ 8 (index 7)
        note:   saved[8] || ""                             // คอลัมน์ 9 (index 8)
      };
    }
    // ถ้าไม่เคยตอบเลย ส่งแค่ชื่อเป็น String ไป (ฝั่ง HTML จะรู้เองว่าต้องแสดงเป็นค่าว่าง)
    return memberName;
  }).filter(Boolean);

  // จุดที่ปรับปรุง: ถ้าคอลัมน์ I (index 8) ว่าง ให้ส่งค่า "OPEN" ไปแทน
    const rawStatus = rows[0][8]; 
    const currentStatus = (rawStatus === "CLOSED") ? "CLOSED" : "OPEN";

  // 4. ส่งค่ากลับไปยังหน้าเว็บ
  const result = {
    meetingId: mid,
    meetingName: rows[0][1], // ชีต3 คอลัมน์ 2
    meetingDate: rows[0][2] instanceof Date ? rows[0][2].toISOString() : rows[0][2],
    meetingTime: rows[0][3],
    meetingPlace: rows[0][4],
    surveyBy: rows[0][5],
    status_system: rows[0][8] || "OPEN", // 👈 เพิ่มบรรทัดนี้เพื่อดึงค่าจากคอลัมน์ I
    status_system: currentStatus, 
    members: mappedMembers 
  };
  
  console.log("ส่งข้อมูลสำเร็จสำหรับ: " + result.meetingName);
  return result;
}

/* ===============================
   บันทึกข้อมูลลง "ชีต2" (คอลัมน์ A-I)
================================ */
function saveSurveyResult(d) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName("ชีต2");
  if (!sh) throw new Error("ไม่พบชีต2");

  sh.appendRow([
    new Date(),             // A: ประทับวันและเวลา
    d.meetingId,            // B: meetingId
    d.meetingName,          // C: ชื่อการประชุม
    d.memberName,           // D: ชื่อกรรมการ
    d.onsite ? "TRUE" : "", // E: Onsite
    d.online ? "TRUE" : "", // F: Online
    d.leave  ? "TRUE" : "", // G: ลา
    d.food   ? "TRUE" : "", // H: อาหาร
    d.note   || ""          // I: หมายเหตุ
  ]);
}

function getMeetingNameList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName("ชีต3");
  if (!sh) return [];

  // ดึงข้อมูลชื่อการประชุมจากคอลัมน์ B (index 1)
  const values = sh.getRange("B2:B" + sh.getLastRow())
    .getValues()
    .flat()         // แปลงจาก Array ซ้อน Array ให้เป็น Array ชั้นเดียว
    .filter(String); // กรองเอาเฉพาะแถวที่มีตัวอักษร (ไม่เอาแถวว่าง)

  // 🔥 ใช้ Set เพื่อกำจัดชื่อที่ซ้ำกันออกไป และส่งกลับเป็น Array
  return [...new Set(values)];
}

function getMembersByMeetingName(meetingName) {
  if (!meetingName) return [];

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName("ชีต3");
  if (!sh) return [];

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();

  return data
    .filter(r => r[1] === meetingName)
    .map(r => r[6])
    .filter((v, i, a) => v && a.indexOf(v) === i);
}

/**
 * เพิ่มรายชื่อกรรมการคนใหม่เข้าไปในชีต3 สำหรับการประชุมเดิมที่มีอยู่แล้ว
 */
function addMemberToMeeting(mid, newMemberName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName("ชีต3");
  const data = sh.getDataRange().getValues();
  
  // ค้นหาแถวแรกที่มี Meeting ID ตรงกัน เพื่อคัดลอกข้อมูลส่วนหัว (Meeting Info)
  let meetingInfo = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(mid).trim()) {
      meetingInfo = data[i]; // เก็บข้อมูลแถวนี้ไว้เป็นต้นแบบ
      break;
    }
  }

  if (meetingInfo) {
    // เพิ่มแถวใหม่: คัดลอกคอลัมน์ A-F และใส่ชื่อใหม่ในคอลัมน์ G
    sh.appendRow([
      meetingInfo[0], // A: ID
      meetingInfo[1], // B: ชื่อประชุม
      meetingInfo[2], // C: วันที่
      meetingInfo[3], // D: เวลา
      meetingInfo[4], // E: สถานที่
      meetingInfo[5], // F: ผู้สำรวจ
      newMemberName   // G: ชื่อกรรมการคนใหม่
    ]);
    return true;
  }
  throw new Error("ไม่พบข้อมูลการประชุมต้นฉบับ");
}

/**
 * ดึงรายชื่อกรรมการทั้งหมดจาก ชีต1 (คอลัมน์ A) มาทำเป็น Datalist
 */
function getMasterMemberList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName("ชีต1"); // ชื่อชีตที่มีรายชื่อกรรมการหลัก
  if (!sh) return [];
  
  const data = sh.getRange("A2:A" + sh.getLastRow()).getValues();
  return data.map(r => r[0]).filter(name => name !== ""); // กรองเฉพาะแถวที่มีชื่อ
}

/**
 * ลบรายชื่อกรรมการออกจากชีต3
 */
function removeMemberFromMeeting(mid, memberName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName("ชีต3");
  const data = sh.getDataRange().getValues();
  
  // ค้นหาจากล่างขึ้นบนเพื่อป้องกัน Error จากการขยับแถว
  for (let i = data.length - 1; i >= 1; i--) {
    // คอลัมน์ A (index 0) คือ Meeting ID และ G (index 6) คือชื่อกรรมการ
    if (String(data[i][0]).trim() === String(mid).trim() && 
        String(data[i][6]).trim() === String(memberName).trim()) {
      sh.deleteRow(i + 1);
      return true;
    }
  }
  throw new Error("ไม่พบข้อมูลที่ต้องการลบในฐานข้อมูล");
}

/**
 * 1. ฟังก์ชันส่งรหัสผ่านไปยังอีเมล (ดึงจากคอลัมน์ J)
 */
/**
 * 1. ฟังก์ชันส่งรหัสผ่านไปยังอีเมล (ดึงจากคอลัมน์ J และตรวจสอบ K)
 */
/**
 * 1. ฟังก์ชันส่งรหัสผ่านไปยังอีเมล (ดึงจากคอลัมน์ J และตรวจสอบ K)
 */
function sendAdminPasscode(mid) {
  // บังคับคำนวณชีตใหม่ทั้งหมด
  SpreadsheetApp.flush(); 

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh3 = ss.getSheetByName("ชีต3");
  const data = sh3.getDataRange().getValues();
  
  let passcode = "";
  let email = "";
  let meetingName = "";
  let foundRows = [];

  // 1. วนหาข้อมูล
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(mid).trim()) {
      foundRows.push(i + 1);
      if (!meetingName) meetingName = data[i][1];
      
      // อ่านรหัส J (Index 9)
      let p = String(data[i][9]).trim();
      if (!passcode && p !== "" && p !== "undefined") passcode = p;

      // --- จุดที่แก้: อ่าน K ตรงๆ (Index 10) ---
      let e = String(data[i][10]).trim();
      if (!email && e !== "" && e.indexOf('@') !== -1) {
        email = e;
      }
    }
  }

  // 2. สร้างรหัสใหม่ถ้าไม่มี
  if (!passcode) {
    passcode = Math.floor(1000 + Math.random() * 9000).toString();
    if (foundRows.length > 0) {
      foundRows.forEach(row => {
        sh3.getRange(row, 10).setValue(passcode);
      });
    }
  }

  // 3. ตรวจสอบอีเมลและส่ง
  if (email && email.indexOf('@') !== -1) {
    try {
      MailApp.sendEmail(email, "🔑 รหัสผ่านจัดการระบบสำรวจ: " + meetingName, "รหัสคือ: " + passcode);
      return "ส่งรหัสไปที่อีเมล " + email + " เรียบร้อยแล้ว";
    } catch (err) {
      throw new Error("ส่งเมลไม่สำเร็จ: " + err.message);
    }
  } else {
    // แก้ไขข้อความ Error ให้รู้ว่าเจาะจงที่ ID ไหน
    throw new Error("❌ ไม่พบอีเมลในคอลัมน์ K (ID: " + mid + ")");
  }
}

/**
 * 2. ฟังก์ชันตรวจสอบรหัสและสลับสถานะ OPEN <-> CLOSED
 */
function verifyAndToggleSystem(mid, enteredCode) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh3 = ss.getSheetByName("ชีต3");
  const data = sh3.getDataRange().getValues();
  
  let correctCode = "";
  let currentStatus = "";
  let foundRows = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === mid) {
      correctCode = String(data[i][9]).trim();  // คอลัมน์ J
      currentStatus = String(data[i][8]).trim(); // คอลัมน์ I
      foundRows.push(i + 1);
    }
  }

  if (enteredCode === correctCode && correctCode !== "") {
    // สลับสถานะ: ถ้าว่างหรือเป็น OPEN ให้เป็น CLOSED / ถ้าเป็น CLOSED ให้เป็น OPEN
    const newStatus = (currentStatus === "CLOSED") ? "OPEN" : "CLOSED";
    
    foundRows.forEach(rowIdx => {
      sh3.getRange(rowIdx, 9).setValue(newStatus);
    });
    return { success: true, newStatus: newStatus };
  }
  return { success: false };
}

/**
 * ฟังก์ชันบันทึกคะแนนและคอมเมนต์ลงใน ชีต4
 */
function saveFeedback(comment) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh4 = ss.getSheetByName("ชีต4");
    
    // บันทึกข้อมูล: คอลัมน์ A = 5, คอลัมน์ B = ข้อความคอมเมนต์
    sh4.appendRow([5, comment, new Date()]); 
    
    return "ขอบคุณสำหรับคะแนนความพึงพอใจครับ!";
  } catch (err) {
    throw new Error("เกิดข้อผิดพลาดในการบันทึก: " + err.message);
  }
}
