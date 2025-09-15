// กำหนดชื่อชีท
const SHEET_RECORD = 'sheetrecord';
const SHEET_USER = 'sheetuser';  
const SHEET_ACTIVITY = 'sheetactivity';
const APP_VERSION = '2025-09-15-04'; // update when backend changes

// รับ GET request จาก web app (แก้ไข CORS)
function doGet(e) {
  try {
    const action = e.parameter.action;
    const data = e.parameter.data ? JSON.parse(e.parameter.data) : {};
    
    let result;
    if (action === 'addRecord') {
      result = addRecord(data);
    } else if (action === 'addUser') {
      result = addUser(data);
    } else if (action === 'addActivity') {
      result = addActivity(data);
    } else if (action === 'deleteActivity') {
      result = deleteActivity(data);
    } else if (action === 'updateActivity') {
      result = updateActivity(data);
    } else if (action === 'getInitialData') {
      result = getInitialData();
    } else if (action === 'getDashboardData') {
      result = getDashboardData(data);
    } else if (action === 'getUserInfo') {
      result = getUserInfo(data);
    } else if (action === 'getActivitiesByRoundAndYear') {
      result = getActivitiesByRoundAndYear(data);
    } else if (action === 'authenticateUser') {
      result = authenticateUser(data);
    } else if (action === 'getAllUsers') {
      result = getAllUsers(data);
    } else if (action === 'getActivityReport') {
      result = getActivityReport(data);
    } else if (action === 'exportDashboardToExcel') {
      result = exportDashboardToExcel(data);
    } else if (action === 'uploadImage') {
      // รองรับพารามิเตอร์หลายชื่อ (imageData | data | image)
      result = uploadImageToDrive(
        (typeof data.imageData !== 'undefined' ? data.imageData : (typeof data.data !== 'undefined' ? data.data : data.image)),
        data.filename,
        data.recordId
      );
    } else if (action === 'getVersion') {
      result = {status:'success', version: APP_VERSION, timestamp: new Date().toISOString()};
    } else {
      result = {status: 'error', message: 'Unknown action: ' + action};
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log('doGet Error: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error', 
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// รองรับ POST request ด้วย (สำหรับ fallback)
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    delete params.action;
    
    let result;
    if (action === 'addRecord') {
      result = addRecord(params);
    } else if (action === 'addUser') {
      result = addUser(params);
    } else if (action === 'addActivity') {
      result = addActivity(params);
    } else if (action === 'deleteActivity') {
      result = deleteActivity(params);
    } else if (action === 'updateActivity') {
      result = updateActivity(params);
    } else if (action === 'getInitialData') {
      result = getInitialData();
    } else if (action === 'getDashboardData') {
      result = getDashboardData(params);
    } else if (action === 'getUserInfo') {
      result = getUserInfo(params);
    } else if (action === 'getActivitiesByRoundAndYear') {
      result = getActivitiesByRoundAndYear(params);
    } else if (action === 'authenticateUser') {
      result = authenticateUser(params);
    } else if (action === 'getAllUsers') {
      result = getAllUsers(params);
    } else if (action === 'getActivityReport') {
      result = getActivityReport(params);
    } else if (action === 'exportDashboardToExcel') {
      result = exportDashboardToExcel(params);
    } else if (action === 'uploadImage') {
      result = uploadImageToDrive(params.imageData, params.filename, params.recordId);
    } else if (action === 'getVersion') {
      result = {status:'success', version: APP_VERSION, timestamp: new Date().toISOString()};
    } else {
      result = {status: 'error', message: 'Unknown action: ' + action};
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log('doPost Error: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error', 
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ส่งออก Dashboard เป็นไฟล์ Excel (.xlsx) ใน Google Drive และคืนลิงก์ดาวน์โหลด
function exportDashboardToExcel(params) {
  try {
    // ใช้ฟังก์ชันเดิมเพื่อคำนวณข้อมูล
    const dash = getDashboardData({ year: params.year || '', round: params.round || '' });
    if (dash.status !== 'success') {
      return { status: 'error', message: 'ไม่สามารถประมวลผลข้อมูล Dashboard ได้' };
    }

    const data = dash.data || { summary: {}, userSummary: {}, activityDetails: [], totalRecords: 0 };
    
    // สร้างสเปรดชีตชั่วคราว
    const timestamp = new Date();
    const fileNameBase = `Dashboard_${params.year || 'ทุกปี'}_R${params.round || 'ทุก'}_${Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}`;
    const tempSs = SpreadsheetApp.create(`${fileNameBase}_TEMP`);

    // Sheet: Summary by activity
    const summarySheet = tempSs.getActiveSheet();
    summarySheet.setName('Summary');
    const summaryRows = [['กิจกรรม', 'จำนวนผู้เข้าร่วม']];
    Object.keys(data.summary).forEach(act => {
      summaryRows.push([act, data.summary[act]]);
    });
    if (summaryRows.length === 1) {
      summaryRows.push(['-', 0]);
    }
    summarySheet.getRange(1, 1, summaryRows.length, 2).setValues(summaryRows);
    summarySheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e5e7eb');
    summarySheet.autoResizeColumns(1, 2);

    // Sheet: Users summary
    const usersSheet = tempSs.insertSheet('Users');
    const usersRows = [['USER', 'PScode', 'จำนวนครั้งทั้งหมด', 'กิจกรรมที่เข้าร่วม (ไม่ซ้ำ)']];
    Object.keys(data.userSummary || {}).forEach(userDisplay => {
      const list = data.userSummary[userDisplay] || [];
      const distinct = Array.from(new Set(list));
      const pscode = extractPscode(userDisplay);
      usersRows.push([userDisplay, pscode, list.length, distinct.join(', ')]);
    });
    if (usersRows.length === 1) {
      usersRows.push(['-', '-', 0, '-']);
    }
    usersSheet.getRange(1, 1, usersRows.length, 4).setValues(usersRows);
    usersSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e5e7eb');
    usersSheet.autoResizeColumns(1, 4);

    // Sheet: Details
    const detailsSheet = tempSs.insertSheet('Details');
    const detailsHeader = ['idrecord', 'timestamp', 'กิจกรรม', 'วันที่จัด', 'รอบ', 'ปี', 'ภาพ', 'พิกัด', 'user', 'PScode'];
    const detailsRows = [detailsHeader];
    (data.activityDetails || []).forEach(r => {
      detailsRows.push([
        r.idrecord, r.timestamp, r.activity, r.activityDate, r.round, r.year,
        r.image, r.coordinates, r.user, r.pscode
      ]);
    });
    if (detailsRows.length === 1) {
      detailsRows.push(['-', '-', '-', '-', '-', '-', '-', '-', '-', '-']);
    }
    detailsSheet.getRange(1, 1, detailsRows.length, detailsHeader.length).setValues(detailsRows);
    detailsSheet.getRange(1, 1, 1, detailsHeader.length).setFontWeight('bold').setBackground('#e5e7eb');
    detailsSheet.autoResizeColumns(1, detailsHeader.length);

    // แปลงเป็นไฟล์ Excel (.xlsx)
    const folderName = 'ActivityExports';
    let folder;
    const folders = DriveApp.getFoldersByName(folderName);
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    // ใช้ UrlFetchApp export Spreadsheet -> XLSX (หลีกเลี่ยงปัญหา getAs ไม่รองรับ)
    const exportUrl = `https://docs.google.com/spreadsheets/d/${tempSs.getId()}/export?format=xlsx`;
    const token = ScriptApp.getOAuthToken();
    const fetchResp = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: `Bearer ${token}` },
      muteHttpExceptions: true
    });
    if (fetchResp.getResponseCode() !== 200) {
      throw new Error('Export failed: HTTP ' + fetchResp.getResponseCode() + ' - ' + fetchResp.getContentText());
    }
    const excelBlob = fetchResp.getBlob().setName(`${fileNameBase}.xlsx`);
    const xlsxFile = folder.createFile(excelBlob);
    xlsxFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // ย้ายสเปรดชีตชั่วคราวไปถังขยะ เพื่อไม่ให้รก
    DriveApp.getFileById(tempSs.getId()).setTrashed(true);

    return {
      status: 'success',
      file: {
        name: xlsxFile.getName(),
        id: xlsxFile.getId(),
        url: xlsxFile.getUrl(),
        downloadUrl: `https://drive.google.com/uc?export=download&id=${xlsxFile.getId()}`
      }
    };
  } catch (error) {
    Logger.log('exportDashboardToExcel Error: ' + error.toString());
    return { status: 'error', message: error.toString() };
  }
}

// เพิ่มข้อมูลผู้ใช้
function addUser(params) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_USER);
    if (!sheet) {
      throw new Error(`ไม่พบชีท ${SHEET_USER}`);
    }
    
    // ตรวจสอบว่า PScode ซ้ำหรือไม่
    const existingData = sheet.getDataRange().getValues();
    const existingPScode = existingData.slice(1).find(row => row[0] === params.pscode);
    
    if (existingPScode) {
      return {status: 'error', message: 'PScode นี้มีอยู่ในระบบแล้ว'};
    }
    
    // ตรวจสอบว่า ID 13 หลักซ้ำหรือไม่
    if (params.citizenId) {
      const existingId = existingData.slice(1).find(row => row[1] === params.citizenId);
      if (existingId) {
        return {status: 'error', message: 'เลขบัตรประชาชนนี้มีอยู่ในระบบแล้ว'};
      }
    }
    
    sheet.appendRow([
      params.pscode || '',
      params.citizenId || '',
      params.fullName || '',
      params.group || '',
      params.level || 'pharmacist',
      params.department || '',
      params.password || '@12345',
      params.status || 'TRUE'
    ]);
    
    return {status: 'success', message: 'บันทึกข้อมูลผู้ใช้สำเร็จ'};
  } catch (error) {
    Logger.log('addUser Error: ' + error.toString());
    return {status: 'error', message: error.toString()};
  }
}

// เพิ่มข้อมูลกิจกรรมหลัก
function addActivity(params) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ACTIVITY);
    if (!sheet) {
      throw new Error(`ไม่พบชีท ${SHEET_ACTIVITY}`);
    }
    
    // สร้างรหัสกิจกรรมอัตโนมัติ
    let activityId = params.activityId;
    if (!activityId || activityId.trim() === '') {
      // สร้างรหัสจากปี พ.ศ. + รอบ + เลขลำดับ
      const buddhistYear = params.buddhistYear || (new Date().getFullYear() + 543);
      const round = params.round || 1;
      const yearSuffix = String(buddhistYear).slice(-2); // เอา 2 หลักท้าย เช่น 68 จาก 2568
      
      // หาเลขลำดับถัดไป
      const existingData = sheet.getDataRange().getValues();
      const prefix = `ACT${yearSuffix}${String(round).padStart(2, '0')}`;
      let maxNumber = 0;
      
      existingData.slice(1).forEach(row => {
        const existingId = row[0] || '';
        if (existingId.startsWith(prefix)) {
          const numberPart = existingId.substring(prefix.length);
          const num = parseInt(numberPart);
          if (!isNaN(num) && num > maxNumber) {
            maxNumber = num;
          }
        }
      });
      
      const nextNumber = String(maxNumber + 1).padStart(3, '0');
      activityId = `${prefix}${nextNumber}`;
    }
    
    // ตรวจสอบว่า ID กิจกรรมซ้ำหรือไม่
    const existingData = sheet.getDataRange().getValues();
    const existingId = existingData.slice(1).find(row => row[0] === activityId);
    
    if (existingId) {
      return {status: 'error', message: 'รหัสกิจกรรมนี้มีอยู่ในระบบแล้ว: ' + activityId};
    }
    
    // แปลงปี ค.ศ. เป็น พ.ศ. ถ้าจำเป็น
    let buddhistYear = params.buddhistYear;
    if (!buddhistYear && params.activityDate) {
      const year = new Date(params.activityDate).getFullYear();
      buddhistYear = year + 543;
    }
    
    sheet.appendRow([
      activityId,
      params.activityName || '',
      params.activityDate ? new Date(params.activityDate) : new Date(),
      params.round || 1,
      buddhistYear || (new Date().getFullYear() + 543),
      (typeof params.status !== 'undefined' ? params.status : 'TRUE')
    ]);
    
    return {
      status: 'success', 
      message: 'บันทึกข้อมูลกิจกรรมสำเร็จ',
      activityId: activityId
    };
  } catch (error) {
    Logger.log('addActivity Error: ' + error.toString());
    return {status: 'error', message: error.toString()};
  }
}

// ดึงข้อมูลผู้ใช้จาก PScode
function getUserInfo(params) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_USER);
    if (!sheet) {
      return {status: 'error', message: `ไม่พบชีท ${SHEET_USER}`};
    }
    
    const data = sheet.getDataRange().getValues();
    const userRow = data.slice(1).find(row => row[0] === params.pscode);
    
    if (!userRow) {
      return {status: 'error', message: 'ไม่พบข้อมูลผู้ใช้'};
    }
    
    return {
      status: 'success',
      user: {
        pscode: userRow[0],
        citizenId: userRow[1],
        fullName: userRow[2],
        group: userRow[3],
        level: userRow[4],
        department: userRow[5],
        password: userRow[6],
        status: userRow[7]
      }
    };
  } catch (error) {
    Logger.log('getUserInfo Error: ' + error.toString());
    return {status: 'error', message: error.toString()};
  }
}

// ฟังก์ชันตรวจสอบการเข้าสู่ระบบ
function authenticateUser(params) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_USER);
    if (!sheet) {
      return {status: 'error', message: `ไม่พบชีท ${SHEET_USER}`};
    }
    
    const data = sheet.getDataRange().getValues();
    const userRow = data.slice(1).find(row => 
      row[0] === params.pscode && row[6] === params.password && row[7] === 'TRUE'
    );
    
    if (!userRow) {
      return {status: 'error', message: 'PScode หรือรหัสผ่านไม่ถูกต้อง หรือบัญชีถูกปิดใช้งาน'};
    }
    
    return {
      status: 'success',
      message: 'เข้าสู่ระบบสำเร็จ',
      user: {
        pscode: userRow[0],
        citizenId: userRow[1],
        fullName: userRow[2],
        group: userRow[3],
        level: userRow[4],
        department: userRow[5],
        status: userRow[7]
      }
    };
  } catch (error) {
    Logger.log('authenticateUser Error: ' + error.toString());
    return {status: 'error', message: error.toString()};
  }
}

// ฟังก์ชันดึงรายชื่อผู้ใช้ทั้งหมด (สำหรับ admin)
function getAllUsers(params) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_USER);
    if (!sheet) {
      return {status: 'error', message: `ไม่พบชีท ${SHEET_USER}`};
    }
    
    const data = sheet.getDataRange().getValues();
    const users = data.slice(1).map(row => ({
      pscode: row[0],
      citizenId: row[1],
      fullName: row[2],
      group: row[3],
      level: row[4],
      department: row[5],
      status: row[7]
    }));
    
    return {status: 'success', users};
  } catch (error) {
    Logger.log('getAllUsers Error: ' + error.toString());
    return {status: 'error', message: error.toString()};
  }
}

// เพิ่มข้อมูลกิจกรรม
function addRecord(params) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_RECORD);
    if (!sheet) {
      throw new Error(`ไม่พบชีท ${SHEET_RECORD}`);
    }
    
  // ตรวจสอบว่าผู้ใช้มีอยู่ในระบบหรือไม่
  const userInfo = getUserInfo({pscode: params.pscode});
    if (userInfo.status === 'error') {
      return {status: 'error', message: 'ไม่พบข้อมูลผู้ใช้ กรุณาลงทะเบียนก่อน'};
    }
    
    // ดึงข้อมูลกิจกรรมจาก activity sheet เพื่อหาวันที่จัดและรอบ
    const activitySheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ACTIVITY);
    let activityDate = '';
    let activityRound = params.round || ''; // ใช้ round ที่ส่งมาจาก frontend ก่อน
    let activityYear = '';
    
    if (activitySheet && activitySheet.getLastRow() > 1) {
      const activityData = activitySheet.getDataRange().getValues();
      const selectedActivity = activityData.slice(1).find(row => row[1] === params.activity);
      
      if (selectedActivity) {
        activityDate = selectedActivity[2] ? formatDate(new Date(selectedActivity[2])) : '';
        // ถ้าไม่ได้ส่ง round มา ให้ใช้จาก activity sheet
        if (!params.round) {
          activityRound = selectedActivity[3] || '';
        }
        activityYear = selectedActivity[4] || '';
      }
    }
    
    // ตรวจสอบไม่ให้บันทึกกิจกรรมซ้ำ (กิจกรรม + รอบ + ปี + ผู้ใช้)
    // ใช้ PScode ในการตรวจสอบเพื่อความแม่นยำ
    const allRecords = sheet.getDataRange().getValues();
    const targetActivity = params.activity || '';
    const targetRound = String(activityRound || '');
    const targetYear = String(activityYear || '');
    const targetPscode = userInfo.user.pscode;
    const isDuplicate = allRecords.slice(1).some(row => {
      const rowActivity = row[2] || '';
      const rowRound = String(row[4] || '');
      const rowYear = String(row[5] || '');
      const rowUserDisplay = row[8] || '';
      const rowPscode = extractPscode(rowUserDisplay);
      return rowActivity === targetActivity 
        && rowRound === targetRound 
        && rowYear === targetYear 
        && rowPscode === targetPscode;
    });
    
    if (isDuplicate) {
      return { 
        status: 'error', 
        code: 'DUPLICATE', 
        message: 'ไม่สามารถบันทึกได้: ผู้ใช้นี้บันทึกกิจกรรมนี้แล้ว (กิจกรรม-รอบ-ปี ซ้ำ)'
      };
    }
    
  const idrecord = new Date().getTime();
    const timestamp = new Date();

  // เตรียมชื่อผู้ใช้ในรูปแบบ "ชื่อเต็ม(PScode)"
  const userDisplay = `${userInfo.user.fullName || params.pscode}(${userInfo.user.pscode || params.pscode})`;
    
    // จัดการภาพ: ถ้ามี base64 ส่งมา -> อัพโหลดจริงไป Drive แล้วเก็บ URL, ถ้าไม่มี -> "ไม่มีภาพ"
    let imageInfo = 'ไม่มีภาพ';
    if (params.image && (params.image.data || params.image.filename)) {
      try {
        const uploadResult = uploadImageToDrive(params.image.data || '', params.image.filename, idrecord);
        if (uploadResult.status === 'success') {
          imageInfo = uploadResult.viewUrl || uploadResult.fileUrl || uploadResult.fileName;
        } else {
          imageInfo = params.image.filename || 'อัพโหลดไม่สำเร็จ';
        }
      } catch (imgErr) {
        Logger.log('Image upload fail (continue without stopping): ' + imgErr.toString());
        imageInfo = params.image?.filename || 'อัพโหลดผิดพลาด';
      }
    }
    
    // จัดการพิกัด
    let coordinates = '';
    if (params.coordinates) {
      coordinates = params.coordinates;
    } else if (params.latitude && params.longitude) {
      coordinates = `${params.latitude}, ${params.longitude}`;
    }
    
    sheet.appendRow([
      idrecord,
      timestamp,
      params.activity || '',
      activityDate,
      activityRound,
      activityYear,
      imageInfo,
      coordinates,
      userDisplay
    ]);
    
    return {
      status: 'success', 
      message: 'บันทึกข้อมูลสำเร็จ',
      record: {
        idrecord: idrecord,
        activity: params.activity,
        activityDate: activityDate,
        round: activityRound,
        year: activityYear,
        image: imageInfo,
        coordinates: coordinates,
        user: userDisplay,
        pscode: userInfo.user.pscode
      }
    };
  } catch (error) {
    Logger.log('addRecord Error: ' + error.toString());
    return {status: 'error', message: error.toString()};
  }
}

// ดึงข้อมูลกิจกรรมและ filter
function getInitialData() {
  try {
    const activitySheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ACTIVITY);
    
    // ถ้าไม่มีชีท หรือชีทว่าง สร้างข้อมูลตัวอย่าง
    if (!activitySheet || activitySheet.getLastRow() <= 1) {
      return {
        status: 'success',
        activities: [
          {name: 'กิจกรรมตัวอย่าง 1', date: '15/01/2025', round: 1, buddhistYear: 2568},
          {name: 'กิจกรรมตัวอย่าง 2', date: '20/02/2025', round: 1, buddhistYear: 2568},
          {name: 'กิจกรรมตัวอย่าง 3', date: '10/07/2025', round: 2, buddhistYear: 2568}
        ],
        filters: {
          years: [2567, 2568, 2569],
          rounds: [1, 2],
          buddhistYears: [2567, 2568, 2569]
        }
      };
    }
    
    const data = activitySheet.getDataRange().getValues();
    const activities = data.slice(1).map(row => ({
      id: row[0] || '',
      name: row[1] || 'ไม่ระบุ',
      date: row[2] ? formatDate(new Date(row[2])) : 'ไม่ระบุ',
      round: row[3] || 1,
      buddhistYear: row[4] || (new Date().getFullYear() + 543),
      status: (row[5] !== undefined ? row[5] : 'TRUE')
    }));
    
    // สร้าง filter ปี/รอบจาก record sheet
    const recordSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_RECORD);
    let years = [2567, 2568, 2569]; // ค่าเริ่มต้น ปี พ.ศ.
    
    if (recordSheet && recordSheet.getLastRow() > 1) {
      const recordData = recordSheet.getDataRange().getValues();
      // ดึงปีจากคอลัมน์ที่ 6 (ปี พ.ศ.)
      const recordYears = recordData.slice(1)
        .map(row => row[5])
        .filter(year => year && year !== '')
        .map(year => parseInt(year));
      
      if (recordYears.length > 0) {
        years = [...new Set(recordYears)].sort((a, b) => b - a);
      }
    }
    
    // ดึงรอบจากข้อมูลกิจกรรม
    let rounds = [1, 2];
    if (activities.length > 0) {
      rounds = [...new Set(activities.map(act => act.round))].sort();
    }
    
    // ดึงปี พ.ศ. จากข้อมูลกิจกรรม
    let buddhistYears = years;
    if (activities.length > 0) {
      const activityYears = [...new Set(activities.map(act => act.buddhistYear))];
      buddhistYears = [...new Set([...years, ...activityYears])].sort((a, b) => b - a);
    }
    
    return {
      status: 'success',
      activities,
      filters: { 
        years: years.map(y => y - 543), // ปี ค.ศ. สำหรับ backward compatibility
        rounds, 
        buddhistYears // ปี พ.ศ. สำหรับ activity และ record
      }
    };
  } catch (error) {
    Logger.log('getInitialData Error: ' + error.toString());
    return {
      status: 'success',
      activities: [
        {name: 'กิจกรรมตัวอย่าง 1', date: '15/01/2025', round: 1, buddhistYear: 2568},
        {name: 'กิจกรรมตัวอย่าง 2', date: '20/02/2025', round: 2, buddhistYear: 2568}
      ],
      filters: {
        years: [2567, 2568],
        rounds: [1, 2],
        buddhistYears: [2567, 2568]
      }
    };
  }
}

// Dashboard summary
function getDashboardData(params) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_RECORD);
    if (!sheet || sheet.getLastRow() <= 1) {
      return {
        status: 'success',
        data: { 
          summary: {},
          totalRecords: 0,
          filters: params
        }
      };
    }
    
    const data = sheet.getDataRange().getValues();
    let filtered = data.slice(1); // ข้าม header
    
    // Filter by year (ปี พ.ศ.)
    if (params.year && params.year !== '') {
      filtered = filtered.filter(row => String(row[5]) === String(params.year));
    }
    
    // Filter by round
    if (params.round && params.round !== '') {
      filtered = filtered.filter(row => String(row[4]) === String(params.round));
    }
    
    // สร้าง summary
    const summary = {};
    const userSummary = {};
    const activityDetails = [];
    
    filtered.forEach(row => {
      const activity = row[2] || 'ไม่ระบุ';
      const activityDate = row[3] || 'ไม่ระบุ';
      const round = row[4] || 'ไม่ระบุ';
      const year = row[5] || 'ไม่ระบุ';
      const userDisplay = row[8] || 'ไม่ระบุ';
      const user = userDisplay; // สำหรับ key แสดงผล
      const pscode = extractPscode(userDisplay);
      
      // นับตามกิจกรรม
      summary[activity] = (summary[activity] || 0) + 1;
      
      // นับตามผู้ใช้
      if (!userSummary[user]) {
        userSummary[user] = [];
      }
      userSummary[user].push(activity);
      
      // รายละเอียดกิจกรรม
      activityDetails.push({
        idrecord: row[0],
        timestamp: row[1] ? formatDate(new Date(row[1])) : 'ไม่ระบุ',
        activity: activity,
        activityDate: activityDate,
        round: round,
        year: year,
        image: row[6] || '',
        coordinates: row[7] || '',
        user: userDisplay,
        pscode: pscode
      });
    });
    
    return {
      status: 'success',
      data: { 
        summary,
        userSummary,
        activityDetails,
        totalRecords: filtered.length,
        filters: params
      }
    };
  } catch (error) {
    Logger.log('getDashboardData Error: ' + error.toString());
    return {status: 'error', message: error.toString()};
  }
}

// ฟังก์ชันช่วยในการจัดรูปแบบวันที่
function formatDate(date) {
  if (!(date instanceof Date)) return 'ไม่ระบุ';
  
  const day = date.getDate().toString().padStart(2, '0');
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const year = date.getFullYear();
  
  return `${day}/${month}/${year}`;
}

// ฟังก์ชันแปลงปี ค.ศ. เป็น พ.ศ.
function toBuddhistYear(christianYear) {
  return parseInt(christianYear) + 543;
}

// ฟังก์ชันแปลงปี พ.ศ. เป็น ค.ศ.
function toChristianYear(buddhistYear) {
  return parseInt(buddhistYear) - 543;
}

// แยก PScode จากรูปแบบ "ชื่อเต็ม(PScode)" หรือคืนค่าเดิมถ้าไม่มีวงเล็บ
function extractPscode(userCell) {
  const s = String(userCell || '').trim();
  const m = s.match(/\(([^)]+)\)\s*$/);
  return m ? m[1].trim() : s;
}

// ฟังก์ชันอัปโหลดภาพไปยัง Google Drive
function uploadImageToDrive(imageData, filename, recordId) {
  try {
    // สร้างโฟลเดอร์สำหรับเก็บภาพ (ถ้ายังไม่มี)
    let folder;
    const folders = DriveApp.getFoldersByName('ActivityImages');
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder('ActivityImages');
    }
    
    // แปลง base64 เป็น blob (รองรับทั้งแบบมี/ไม่มี prefix)
    if (!imageData) {
      return { status: 'error', message: 'No image data provided' };
    }

    let pureBase64 = String(imageData);
    // หากมี prefix เช่น data:image/jpeg;base64,xxx ให้ตัดส่วนหน้าออก
    if (pureBase64.indexOf(',') !== -1) {
      pureBase64 = pureBase64.split(',').pop();
    }
    // ลบช่องว่าง/ขึ้นบรรทัดใหม่ที่อาจปะปนมา
    pureBase64 = pureBase64.replace(/\s/g, '');

    let decodedBytes;
    try {
      decodedBytes = Utilities.base64Decode(pureBase64);
    } catch (decodeErr) {
      return { status: 'error', message: 'Invalid base64 image data' };
    }

    const safeFilename = filename || `activity_${recordId}.jpg`;
    const blob = Utilities.newBlob(decodedBytes, 'image/jpeg', safeFilename);
    
    // อัปโหลดไฟล์
    const file = folder.createFile(blob);
    
    // ทำให้ไฟล์สามารถเข้าถึงได้แบบสาธารณะ
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return {
      status: 'success',
      fileId: file.getId(),
      fileName: file.getName(),
      fileUrl: file.getUrl(),
      viewUrl: `https://drive.google.com/file/d/${file.getId()}/view`
    };
  } catch (error) {
    Logger.log('uploadImageToDrive Error: ' + error.toString());
    return {
      status: 'error',
      message: error.toString()
    };
  }
}

// ฟังก์ชันลบกิจกรรม
function deleteActivity(params) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ACTIVITY);
    if (!sheet) {
      return {status: 'error', message: `ไม่พบชีท ${SHEET_ACTIVITY}`};
    }
    
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex((row, index) => 
      index > 0 && row[1] === params.name // หาจากชื่อกิจกรรมในคอลัมน์ที่ 2
    );
    
    if (rowIndex === -1) {
      return {status: 'error', message: 'ไม่พบกิจกรรมที่ต้องการลบ'};
    }
    
    // ลบแถว (rowIndex + 1 เพราะ Google Sheets เริ่มนับจาก 1)
    sheet.deleteRow(rowIndex + 1);
    
    return {
      status: 'success', 
      message: `ลบกิจกรรม "${params.name}" เรียบร้อยแล้ว`
    };
  } catch (error) {
    Logger.log('deleteActivity Error: ' + error.toString());
    return {status: 'error', message: error.toString()};
  }
}

// อัปเดตกิจกรรม
function updateActivity(params) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ACTIVITY);
    if (!sheet) {
      return {status: 'error', message: `ไม่พบชีท ${SHEET_ACTIVITY}`};
    }

    if (!params.originalName) {
      return {status: 'error', message: 'ไม่พบชื่อกิจกรรมเดิม (originalName)'};
    }

    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex((row, index) => index > 0 && row[1] === params.originalName);
    if (rowIndex === -1) {
      return {status: 'error', message: 'ไม่พบกิจกรรมที่ต้องการแก้ไข'};
    }

    // ดึงค่าเดิม
    const existingRow = data[rowIndex];
    const currentId = existingRow[0];

    // เตรียมค่าที่จะอัปเดต (ไม่สร้างรหัสใหม่)
    const newName = params.activityName || existingRow[1];
    const newDate = params.activityDate ? new Date(params.activityDate) : existingRow[2];
    const newRound = params.round || existingRow[3] || 1;
    let newBuddhistYear = params.buddhistYear || existingRow[4];
    if (!newBuddhistYear && newDate) {
      newBuddhistYear = new Date(newDate).getFullYear() + 543;
    }

    // กำหนดสถานะ (ถ้ารับมาให้ใช้ ถ้าไม่ให้คงเดิม)
    const newStatus = (typeof params.status !== 'undefined') ? params.status : existingRow[5];

    sheet.getRange(rowIndex + 1, 1, 1, 6).setValues([[
      currentId,
      newName,
      newDate,
      newRound,
      newBuddhistYear,
      newStatus
    ]]);

    return {status: 'success', message: 'อัปเดตกิจกรรมสำเร็จ', activityId: currentId};
  } catch (error) {
    Logger.log('updateActivity Error: ' + error.toString());
    return {status: 'error', message: error.toString()};
  }
}

// ฟังก์ชันดึงข้อมูลตำแหน่งจากพิกัด (ใช้ reverse geocoding)
function getLocationFromCoordinates(latitude, longitude) {
  try {
    // ใช้ Google Maps Geocoding API (ต้อง enable API ก่อน)
    // const response = Maps.newGeocoder().reverseGeocode(latitude, longitude);
    // return response.results[0].formatted_address;
    
    // สำหรับตอนนี้ส่งคืนพิกัดเป็น string
    return `${latitude}, ${longitude}`;
  } catch (error) {
    Logger.log('getLocationFromCoordinates Error: ' + error.toString());
    return `${latitude}, ${longitude}`;
  }
}

// ฟังก์ชันดึงข้อมูลกิจกรรมตามรอบและปี
function getActivitiesByRoundAndYear(params) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ACTIVITY);
    if (!sheet || sheet.getLastRow() <= 1) {
      return {status: 'success', activities: []};
    }
    
    const data = sheet.getDataRange().getValues();
    let activities = data.slice(1).map(row => ({
      id: row[0] || '',
      name: row[1] || 'ไม่ระบุ',
      date: row[2] ? formatDate(new Date(row[2])) : 'ไม่ระบุ',
      round: row[3] || 1,
      buddhistYear: row[4] || (new Date().getFullYear() + 543),
      status: (row[5] !== undefined ? row[5] : 'TRUE')
    }));
  
    // กรองเฉพาะ active เป็นค่าเริ่มต้น (TRUE/true/1)
    if (!params || !params.includeInactive) {
      activities = activities.filter(act => {
        const v = (act.status !== undefined) ? act.status : 'TRUE';
        return v === true || v === 1 || String(v).toUpperCase() === 'TRUE';
      });
    }
    
    // กรองตามรอบ
    if (params.round && params.round !== '') {
      activities = activities.filter(act => String(act.round) === String(params.round));
    }
    
    // กรองตามปี พ.ศ.
    if (params.buddhistYear && params.buddhistYear !== '') {
      activities = activities.filter(act => String(act.buddhistYear) === String(params.buddhistYear));
    }
    
    return {status: 'success', activities};
  } catch (error) {
    Logger.log('getActivitiesByRoundAndYear Error: ' + error.toString());
    return {status: 'error', message: error.toString()};
  }
}

// ฟังก์ชันดึงรายงานการเข้าร่วมกิจกรรม
function getActivityReport(params) {
  try {
    const recordSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_RECORD);
    const userSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_USER);
    
    if (!recordSheet || !userSheet) {
      return {status: 'error', message: 'ไม่พบข้อมูลที่จำเป็น'};
    }
    
    const recordData = recordSheet.getDataRange().getValues();
    const userData = userSheet.getDataRange().getValues();
    
    // สร้าง mapping ของผู้ใช้
    const userMap = {};
    userData.slice(1).forEach(row => {
      userMap[row[0]] = {
        pscode: row[0],
        fullName: row[2],
        group: row[3],
        level: row[4],
        department: row[5]
      };
    });
    
    // ประมวลผลข้อมูล record
    let records = recordData.slice(1).map(row => {
      const userDisplay = row[8] || 'ไม่ระบุ';
      const pscode = extractPscode(userDisplay);
      return ({
        idrecord: row[0],
        timestamp: row[1] ? formatDate(new Date(row[1])) : 'ไม่ระบุ',
        activity: row[2] || 'ไม่ระบุ',
        activityDate: row[3] || 'ไม่ระบุ',
        round: row[4] || 'ไม่ระบุ',
        year: row[5] || 'ไม่ระบุ',
        image: row[6] || '',
        coordinates: row[7] || '',
        user: userDisplay,
        pscode: pscode,
        userInfo: userMap[pscode] || null
      });
    });
    
    // กรองตามเงื่อนไข
    if (params.year && params.year !== '') {
      records = records.filter(record => String(record.year) === String(params.year));
    }
    
    if (params.round && params.round !== '') {
      records = records.filter(record => String(record.round) === String(params.round));
    }
    
    if (params.activity && params.activity !== '') {
      records = records.filter(record => record.activity === params.activity);
    }
    
    return {
      status: 'success',
      records: records,
      totalRecords: records.length
    };
  } catch (error) {
    Logger.log('getActivityReport Error: ' + error.toString());
    return {status: 'error', message: error.toString()};
  }
}

// ฟังก์ชันสำหรับสร้างชีทเริ่มต้น (ถ้าต้องการ)
function setupSheets() {
  const ss = SpreadsheetApp.getActive();
  
  // สร้าง sheetrecord
  let recordSheet = ss.getSheetByName(SHEET_RECORD);
  if (!recordSheet) {
    recordSheet = ss.insertSheet(SHEET_RECORD);
    recordSheet.getRange(1, 1, 1, 9).setValues([[
      'idrecord', 'timestamp', 'กิจกรรม', 'วันที่จัด', 'รอบ', 'ปี', 'ภาพ', 'พิกัด', 'user'
    ]]);
    
    // จัดรูปแบบ header
    const headerRange = recordSheet.getRange(1, 1, 1, 9);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    recordSheet.setFrozenRows(1);
    
    // ตั้งความกว้างคอลัมน์
    recordSheet.setColumnWidth(1, 120); // idrecord
    recordSheet.setColumnWidth(2, 150); // timestamp
    recordSheet.setColumnWidth(3, 200); // กิจกรรม
    recordSheet.setColumnWidth(4, 120); // วันที่จัด
    recordSheet.setColumnWidth(5, 60);  // รอบ
    recordSheet.setColumnWidth(6, 60);  // ปี
    recordSheet.setColumnWidth(7, 150); // ภาพ
    recordSheet.setColumnWidth(8, 200); // พิกัด
    recordSheet.setColumnWidth(9, 80);  // user
  }
  
  // สร้าง sheetuser
  let userSheet = ss.getSheetByName(SHEET_USER);
  if (!userSheet) {
    userSheet = ss.insertSheet(SHEET_USER);
    userSheet.getRange(1, 1, 1, 8).setValues([[
      'PScode', 'ID 13 หลัก', 'ชื่อ-นามสกุล', 'กลุ่ม', 'ระดับ', 'หน่วยงาน', 'รหัสผ่าน', 'status'
    ]]);
    
    // เพิ่มข้อมูลจริง
    userSheet.getRange(2, 1, 14, 8).setValues([
      ['P01', '1234567890123', 'ทดสอบ ทดสอบ', 'เภสัชกร', 'supervisor', '', '@12345', 'TRUE'],
      ['P02', '3460700549570', 'ภญ.อาศิรา ภูศรีดาว', 'เภสัชกร', 'pharmacist', '', '@12345', 'TRUE'],
      ['P03', '3349900018143', 'ภญ.ชุติธนา ภัทรทิวานนท์', 'เภสัชกร', 'pharmacist', '', '@12345', 'TRUE'],
      ['P04', '3330600079866', 'ภญ.บงกช อินทร์พิมพ์', 'เภสัชกร', 'pharmacist', '', '@12345', 'TRUE'],
      ['P05', '3339900162368', 'ภก.นพพร บัวสี', 'เภสัชกร', 'pharmacist', '', '@12345', 'TRUE'],
      ['P06', '3329900242361', 'ภญ.วชิรา สุเมธิวิทย์', 'เภสัชกร', 'pharmacist', '', '@12345', 'TRUE'],
      ['P07', '3349900984163', 'ภญ.อารยา ถวิลหวัง', 'เภสัชกร', 'pharmacist', '', '@12345', 'TRUE'],
      ['P08', '3339900174951', 'ภญ.ณภัคอร ช่วงสกุล', 'เภสัชกร', 'pharmacist', '', '@12345', 'TRUE'],
      ['P09', '3339900155477', 'ภญ.อิสริยาภรณ์ บุญสังข์', 'เภสัชกร', 'pharmacist', 'คลัง', '@12345', 'TRUE'],
      ['P10', '3330900702038', 'ภญ.เกศสุภา พลพงษ์', 'เภสัชกร', 'pharmacist', 'DIS', '@12345', 'TRUE'],
      ['P11', '3350200021771', 'ภก.กฤษฎา บูราณ', 'เภสัชกร', 'pharmacist', 'จ่ายยาER', '@12345', 'TRUE'],
      ['P12', '3100200159376', 'ภญ.ภิญรัตน์ มหาลีวีรัศมี', 'เภสัชกร', 'pharmacist', 'จ่ายผู้ป่วยใน', '@12345', 'TRUE'],
      ['P13', '3320100275241', 'ภก.สุทธินันท์ เอิกเกริก', 'เภสัชกร', 'admin', 'จ่ายผู้ป่วยนอก', 'admin123', 'TRUE'],
      ['P14', '3440100150234', 'ภญ.แพรวพิมพ์อร ภูเฮืองแก้ว', 'เภสัชกร', 'pharmacist', 'จ่ายผู้ป่วยใน', '@12345', 'TRUE']
    ]);
    
    // จัดรูปแบบ header
    const headerRange = userSheet.getRange(1, 1, 1, 8);
    headerRange.setBackground('#34a853');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    userSheet.setFrozenRows(1);
    
    // ตั้งความกว้างคอลัมน์
    userSheet.setColumnWidth(1, 60);  // PScode
    userSheet.setColumnWidth(2, 140); // ID 13 หลัก
    userSheet.setColumnWidth(3, 200); // ชื่อ-นามสกุล
    userSheet.setColumnWidth(4, 80);  // กลุ่ม
    userSheet.setColumnWidth(5, 100); // ระดับ
    userSheet.setColumnWidth(6, 150); // หน่วยงาน
    userSheet.setColumnWidth(7, 80);  // รหัสผ่าน
    userSheet.setColumnWidth(8, 60);  // status
  }
  
  // สร้าง sheetactivity
  let activitySheet = ss.getSheetByName(SHEET_ACTIVITY);
  if (!activitySheet) {
    activitySheet = ss.insertSheet(SHEET_ACTIVITY);
    activitySheet.getRange(1, 1, 1, 6).setValues([[
      'idกิจกรรม', 'รายการกิจกรรม', 'วันที่จัดงาน', 'รอบ', 'ปี พ.ศ.', 'status'
    ]]);

    // เพิ่มข้อมูลตัวอย่าง
    activitySheet.getRange(2, 1, 8, 6).setValues([
      ['ACT001', 'การอบรมเภสัชกรรม', new Date('2025-01-15'), 1, 2568, 'TRUE'],
      ['ACT002', 'สัมมนาความปลอดภัยผู้ป่วย', new Date('2025-02-20'), 1, 2568, 'TRUE'],
      ['ACT003', 'ประชุมวิชาการประจำปี', new Date('2025-03-10'), 1, 2568, 'TRUE'],
      ['ACT004', 'อบรม CPR และการช่วยฟื้นคืนชีพ', new Date('2025-04-05'), 1, 2568, 'TRUE'],
      ['ACT005', 'การประเมินคุณภาพการดูแลผู้ป่วย', new Date('2025-05-12'), 1, 2568, 'TRUE'],
      ['ACT006', 'การอบรมเภสัชกรรมครั้งที่ 2', new Date('2025-07-15'), 2, 2568, 'TRUE'],
      ['ACT007', 'สัมมนาความปลอดภัยผู้ป่วยครั้งที่ 2', new Date('2025-08-20'), 2, 2568, 'TRUE'],
      ['ACT008', 'ประชุมวิชาการรอบ 2', new Date('2025-09-10'), 2, 2568, 'TRUE']
    ]);

    // จัดรูปแบบ header
    const headerRange = activitySheet.getRange(1, 1, 1, 6);
    headerRange.setBackground('#ff9800');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    activitySheet.setFrozenRows(1);

    // ตั้งความกว้างคอลัมน์
    activitySheet.setColumnWidth(1, 100); // idกิจกรรม
    activitySheet.setColumnWidth(2, 250); // รายการกิจกรรม
    activitySheet.setColumnWidth(3, 120); // วันที่จัดงาน
    activitySheet.setColumnWidth(4, 60);  // รอบ
    activitySheet.setColumnWidth(5, 80);  // ปี พ.ศ.
    activitySheet.setColumnWidth(6, 80);  // status
  }
  
  return 'สร้างชีทเรียบร้อยแล้ว: ' + SHEET_RECORD + ', ' + SHEET_USER + ', ' + SHEET_ACTIVITY;
}

// ฟังก์ชันทดสอบการทำงาน
function testFunctions() {
  // ทดสอบ getInitialData
  Logger.log('Testing getInitialData:');
  const initialData = getInitialData();
  Logger.log(initialData);
  
  // ทดสอบ getUserInfo
  Logger.log('Testing getUserInfo:');
  const userInfo = getUserInfo({pscode: 'P01'});
  Logger.log(userInfo);
  
  // ทดสอบ authenticateUser
  Logger.log('Testing authenticateUser:');
  const authResult = authenticateUser({pscode: 'P01', password: '@12345'});
  Logger.log(authResult);
  
  // ทดสอบ getAllUsers
  Logger.log('Testing getAllUsers:');
  const allUsers = getAllUsers({});
  Logger.log(allUsers);
  
  // ทดสอบ addRecord
  Logger.log('Testing addRecord:');
  const addRecordResult = addRecord({
    pscode: 'P01',
    activity: 'การอบรมเภสัชกรรม',
    round: 1,
    coordinates: '15.1161, 104.3203',
    image: {filename: 'test_image.jpg', data: 'base64data'},
    latitude: '15.1161',
    longitude: '104.3203'
  });
  Logger.log(addRecordResult);
  
  return 'ทดสอบเสร็จสิ้น ตรวจสอบ Logs';
}

// ฟังก์ชันลบข้อมูลทดสอบ
function clearTestData() {
  const sheets = [SHEET_RECORD, SHEET_USER, SHEET_ACTIVITY];
  
  sheets.forEach(sheetName => {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (sheet && sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
  });
  
  return 'ลบข้อมูลทดสอบเรียบร้อยแล้ว';
}
