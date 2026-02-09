/**
 * ------------------------------------------------------------------
 * GOOGLE APPS SCRIPT - VERSION 5.1 (High Performance & Concurrency)
 * ------------------------------------------------------------------
 */

const SPREADSHEET_ID = ""; 
const SHEET_USERS = "Users";
const SHEET_SETTINGS = "Settings";
const SHEET_DATA_MASTER = "Data";
const SHEET_CRITERIA = "Criteria";

function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents);
    const action = request.action;

    if (action === 'login') {
      return handleLogin(request.username, request.password);
    } 
    else if (action === 'switchRoom') {
      return handleSwitchRoom(request.room, request.level);
    }
    else if (action === 'getAdminStats') {
      return getAdminOverview();
    }
    else if (action === 'saveScores') {
      return handleSaveScores(request);
    }
    
    return responseJSON({ status: 'error', message: 'Unknown action' });

  } catch (err) {
    return responseJSON({ status: 'error', message: err.toString() });
  }
}

// ------------------------------------------------------------------
// Core Logic
// ------------------------------------------------------------------

function handleLogin(username, password) {
  // Lock Critical Section: การสร้างชีตใหม่ไม่ควรทำพร้อมกัน
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // รอคิวสูงสุด 10 วินาที
  } catch (e) {
    return responseJSON({ status: 'error', message: 'ระบบกำลังทำงานหนัก กรุณาลองใหม่' });
  }

  try {
    const ss = getSS();
    const userSheet = ss.getSheetByName(SHEET_USERS);
    if (!userSheet) return responseJSON({ status: 'error', message: 'ไม่พบชีต Users' });
    
    const usersData = userSheet.getDataRange().getDisplayValues();
    const user = usersData.slice(1).find(r => r[0] == username && r[1] == password);
    
    if (!user) {
      return responseJSON({ status: 'error', message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' });
    }

    const role = user[4].toLowerCase() === 'admin' ? 'admin' : 'teacher';
    
    if (role === 'admin') {
      const allRooms = usersData.slice(1)
        .filter(r => r[3] && r[4] && r[4].toLowerCase() !== 'admin')
        .map(r => ({ teacher: r[2], room: r[3], level: r[4] }));

      let firstRoomData = {};
      if (allRooms.length > 0) {
        // Load data without lock inside (logic separated)
        firstRoomData = loadRoomDataInternal(ss, allRooms[0].room, allRooms[0].level);
      }

      return responseJSON({
        status: 'success',
        role: 'admin',
        user: { name: user[2], room: 'All', level: 'Admin' },
        allRooms: allRooms,
        ...firstRoomData
      });

    } else {
      const roomData = loadRoomDataInternal(ss, user[3], user[4]);
      return responseJSON({
        status: 'success',
        role: 'teacher',
        user: { name: user[2], room: user[3], level: user[4] },
        ...roomData
      });
    }
  } catch (err) {
    return responseJSON({ status: 'error', message: err.toString() });
  } finally {
    lock.releaseLock(); // ปลดล็อคเสมอ
  }
}

function handleSwitchRoom(room, level) {
  const ss = getSS();
  // Reuse internal logic (no need for global lock here as mostly reading)
  // But creating sheet inside loadRoomData needs care, handleLogin already covers creation mostly.
  // We add lock here just in case Admin switches to a room not yet created.
  const lock = LockService.getScriptLock();
  try { lock.waitLock(5000); } catch (e) {}

  try {
    const data = loadRoomDataInternal(ss, room, level);
    return responseJSON({ status: 'success', ...data });
  } finally {
    lock.releaseLock();
  }
}

// Logic แยกออกมาเพื่อเรียกใช้ภายใน (Internal Use)
function loadRoomDataInternal(ss, targetRoom, targetLevel) {
  const settingSheet = ss.getSheetByName(SHEET_SETTINGS);
  const allSettings = settingSheet.getDataRange().getDisplayValues();
  
  const config = allSettings.slice(1)
    .filter(r => r[0] == targetLevel)
    .map(r => ({
      header: r[1],
      max: parseInt(r[2]),
      colIndex: parseInt(r[3]),
      group: r[4],
      criteriaId: r[5] || null
    }));

  if (config.length === 0) throw new Error(`ไม่พบการตั้งค่าสำหรับระดับชั้น ${targetLevel}`);

  let roomSheet = ss.getSheetByName(targetRoom);
  
  // Create sheet if not exists
  if (!roomSheet) {
    const masterSheet = ss.getSheetByName(SHEET_DATA_MASTER);
    if (!masterSheet) throw new Error('ไม่พบชีตต้นฉบับ (Data)');

    const allData = masterSheet.getDataRange().getDisplayValues();
    const header = allData[0];
    const roomStudents = allData.slice(1).filter(r => r[4] == targetRoom);

    if (roomStudents.length > 0) {
      roomSheet = ss.insertSheet(targetRoom);
      roomSheet.appendRow(header);
      roomSheet.getRange(2, 1, roomStudents.length, roomStudents[0].length).setValues(roomStudents);
    } else {
       return { config: config, criteria: {}, students: [] };
    }
  }

  // Update headers (Batch update for headers is fast enough)
  config.forEach(c => {
    roomSheet.getRange(1, c.colIndex).setValue(c.header);
  });

  // Criteria
  const criteriaSheet = ss.getSheetByName(SHEET_CRITERIA);
  const criteriaMap = {};
  if (criteriaSheet) {
    const criteriaData = criteriaSheet.getDataRange().getDisplayValues().slice(1);
    criteriaData.forEach(row => {
      const id = row[0];
      if (!id) return;
      if (!criteriaMap[id]) criteriaMap[id] = [];
      criteriaMap[id].push({ min: parseFloat(row[1]), max: parseFloat(row[2]), label: row[3] });
    });
  }

  // Get Data
  const roomData = roomSheet.getDataRange().getDisplayValues();
  const students = roomData.slice(1).map((r, i) => {
    const s = {
      rowIndex: i + 2,
      id: r[0],
      title: r[1],
      name: r[2],
      surname: r[3],
      room: r[4],
      scores: {}
    };
    config.forEach(c => {
      s.scores[c.colIndex] = r[c.colIndex - 1] === "" ? null : r[c.colIndex - 1];
    });
    return s;
  });

  return { config, criteria: criteriaMap, students };
}

function getAdminOverview() {
  const ss = getSS();
  const userSheet = ss.getSheetByName(SHEET_USERS);
  const usersData = userSheet.getDataRange().getDisplayValues().slice(1);
  const rooms = usersData
    .filter(r => r[3] && r[4] && r[4].toLowerCase() !== 'admin')
    .map(r => ({ teacher: r[2], room: r[3], level: r[4] }));

  const stats = rooms.map(roomObj => {
    const sheet = ss.getSheetByName(roomObj.room);
    let total = 0;
    let completed = 0;

    if (sheet) {
      const data = sheet.getDataRange().getDisplayValues().slice(1);
      total = data.length;
      completed = countCompletedInSheet(ss, roomObj.level, data);
    } else {
      const master = ss.getSheetByName(SHEET_DATA_MASTER);
      const mData = master.getDataRange().getDisplayValues().slice(1);
      total = mData.filter(r => r[4] == roomObj.room).length;
      completed = 0;
    }

    return {
      room: roomObj.room,
      level: roomObj.level,
      teacher: roomObj.teacher,
      total: total,
      completed: completed,
      percent: total === 0 ? 0 : Math.round((completed / total) * 100)
    };
  });

  return responseJSON({ status: 'success', data: stats });
}

function countCompletedInSheet(ss, level, studentRows) {
  const sSheet = ss.getSheetByName(SHEET_SETTINGS);
  const settings = sSheet.getDataRange().getDisplayValues().slice(1)
    .filter(r => r[0] == level)
    .map(r => parseInt(r[3]) - 1); 

  if (settings.length === 0 || studentRows.length === 0) return 0;

  return studentRows.filter(row => {
    return settings.every(idx => row[idx] !== "" && row[idx] !== null && row[idx] !== undefined);
  }).length;
}

// *** KEY IMPROVEMENT: BATCH UPDATE & LOCKING ***
function handleSaveScores(payload) {
  const lock = LockService.getScriptLock();
  try {
    // รอคิวสูงสุด 10 วินาที หากมีคนอื่นกำลังบันทึกอยู่
    lock.waitLock(10000);
  } catch (e) {
    return responseJSON({ status: 'error', message: 'Server is busy. Please try again.' });
  }

  try {
    const ss = getSS();
    const updates = payload.data; // [{rowIndex, scores: {col:val}}, ...]
    const targetRoom = payload.room;

    if (!targetRoom) return responseJSON({ status: 'error', message: 'ไม่ระบุห้องเรียน' });

    const ws = ss.getSheetByName(targetRoom);
    if (!ws) return responseJSON({ status: 'error', message: `ไม่พบชีตข้อมูลของห้อง ${targetRoom}` });

    // 1. หาคอลัมน์ทั้งหมดที่จะถูกอัปเดต
    let columnsToUpdate = new Set();
    updates.forEach(u => {
      Object.keys(u.scores).forEach(k => columnsToUpdate.add(parseInt(k)));
    });
    const sortedCols = Array.from(columnsToUpdate).sort((a, b) => a - b);
    
    // เตรียม Map ข้อมูลเพื่อค้นหาเร็วๆ (rowIndex -> scores)
    const updateMap = {};
    updates.forEach(u => {
      updateMap[u.rowIndex] = u.scores;
    });

    const lastRow = ws.getLastRow();
    if (lastRow < 2) return responseJSON({ status: 'success', message: 'ไม่มีข้อมูลให้บันทึก' });

    // 2. วนลูปทีละคอลัมน์ (Batch by Column) - เร็วกว่าเขียนทีละเซลล์ 100 เท่า
    sortedCols.forEach(colIndex => {
      // อ่านข้อมูลทั้งคอลัมน์ขึ้นมาในแรม
      const range = ws.getRange(2, colIndex, lastRow - 1, 1);
      const currentValues = range.getValues(); // [[val], [val], ...]
      
      let isModified = false;
      
      // อัปเดตข้อมูลในแรม
      for (let i = 0; i < currentValues.length; i++) {
        const rowIndex = i + 2; // แปลง Array Index เป็น Sheet Row Index
        if (updateMap[rowIndex] && updateMap[rowIndex][colIndex] !== undefined) {
           const newVal = updateMap[rowIndex][colIndex];
           // เช็คว่าค่าเปลี่ยนจริงไหม (ลดการเขียนซ้ำ)
           if (currentValues[i][0] != newVal) {
             currentValues[i][0] = newVal;
             isModified = true;
           }
        }
      }
      
      // เขียนกลับลงชีตทีเดียวทั้งคอลัมน์ (1 API Call per Column)
      if (isModified) {
        range.setValues(currentValues);
      }
    });

    return responseJSON({ status: 'success', message: 'บันทึกข้อมูลเรียบร้อยแล้ว' });

  } catch (err) {
    return responseJSON({ status: 'error', message: err.toString() });
  } finally {
    lock.releaseLock(); // ปลดล็อคเพื่อให้คนต่อไปทำงานต่อ
  }
}

function getSS() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
