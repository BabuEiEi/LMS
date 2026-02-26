// CONFIGURATION
var FOLDER_ID = '1em8kje-hcDCAgVAW8vqhLuze72SgI30B'; // <--- CHANGE THIS TO YOUR FOLDER ID

/**
 * Serve the HTML file
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('LMS Web App')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Setup the Google Sheet and Columns
 */
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Lessons');
  
  if (!sheet) {
    sheet = ss.insertSheet('Lessons');
    // Create Header Row: Added 'links' and 'files'
    sheet.appendRow(['id', 'topic', 'content', 'date_created', 'links', 'files']);
    // Style Header
    var headerRange = sheet.getRange(1, 1, 1, 6);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#26A69A'); // Teal theme color
    headerRange.setFontColor('white');
  } else {
    // Check if new columns exist, if not append them (Simple migration)
    var lastCol = sheet.getLastColumn();
    if (lastCol < 6) {
      sheet.getRange(1, 5).setValue('links');
      sheet.getRange(1, 6).setValue('files');
       var headerRange = sheet.getRange(1, 5, 1, 2);
       headerRange.setFontWeight('bold');
       headerRange.setBackground('#26A69A');
       headerRange.setFontColor('white');
    }
  }

  // Setup Admin Sheet
  var adminSheet = ss.getSheetByName('admin');
  if (!adminSheet) {
    adminSheet = ss.insertSheet('admin');
    adminSheet.appendRow(['username', 'password']);
    // Add default admin
    adminSheet.appendRow(['admin', 'admin123']);
    
    var adminHeader = adminSheet.getRange(1, 1, 1, 2);
    adminHeader.setFontWeight('bold');
    adminHeader.setBackground('#EF5350');
    adminHeader.setFontColor('white');
  }

  // Setup Students Sheet
  var studentsSheet = ss.getSheetByName('students');
  if (!studentsSheet) {
    studentsSheet = ss.insertSheet('students');
    studentsSheet.appendRow(['username', 'password', 'fullname']);
    var h = studentsSheet.getRange(1, 1, 1, 3);
    h.setFontWeight('bold'); h.setBackground('#4285F4'); h.setFontColor('white');
  }

  // Setup Exams Pre Sheet
  var examsPreSheet = ss.getSheetByName('exams_pre');
  if (!examsPreSheet) {
    examsPreSheet = ss.insertSheet('exams_pre');
    examsPreSheet.appendRow(['id', 'question', 'options', 'correct_answer', 'image_url']);
    var h = examsPreSheet.getRange(1, 1, 1, 5);
    h.setFontWeight('bold'); h.setBackground('#FBBC05'); h.setFontColor('white');
  }

  // Setup Exams Post Sheet
  var examsPostSheet = ss.getSheetByName('exams_post');
  if (!examsPostSheet) {
    examsPostSheet = ss.insertSheet('exams_post');
    examsPostSheet.appendRow(['id', 'question', 'options', 'correct_answer', 'image_url']);
    var h = examsPostSheet.getRange(1, 1, 1, 5);
    h.setFontWeight('bold'); h.setBackground('#EA4335'); h.setFontColor('white');
  }

  // Setup Scores Sheet
  var scoresSheet = ss.getSheetByName('scores');
  if (!scoresSheet) {
    scoresSheet = ss.insertSheet('scores');
    scoresSheet.appendRow(['student_username', 'exam_type', 'score', 'timestamp']);
    var h = scoresSheet.getRange(1, 1, 1, 4);
    h.setFontWeight('bold'); h.setBackground('#34A853'); h.setFontColor('white');
  }
}

/**
 * Upload File to Drive
 * @param {String} dataBase64 Base64 encoded file data
 * @param {String} filename Name of the file
 * @param {String} mimeType Mime type of the file
 * @returns {Object} { url, type, name }
 */
function uploadFile(dataBase64, filename, mimeType) {
  try {
    var folder;
    if (FOLDER_ID && FOLDER_ID !== 'YOUR_GOOGLE_DRIVE_FOLDER_ID') {
      try {
        folder = DriveApp.getFolderById(FOLDER_ID);
      } catch(e) {
        // Fallback to root if ID is invalid
        folder = DriveApp.getRootFolder();
      }
    } else {
      folder = DriveApp.getRootFolder();
    }

    var decoded = Utilities.base64Decode(dataBase64);
    var blob = Utilities.newBlob(decoded, mimeType, filename);
    var file = folder.createFile(blob);
    
    // Set permissions to anyone with link (optional, but often needed for public web apps)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var url = file.getUrl();
    
    // As explicitly requested: Force the use of 'lh3.googleusercontent.com/d/' for images.
    // Note: This endpoint is undocumented and may lead to 404s if permissions aren't fully propagated,
    // but we are following the user's strict requirement here.
    if (mimeType.indexOf('image') > -1) {
       url = 'https://lh3.googleusercontent.com/d/' + file.getId();
    }

    return {
      success: true,
      url: url,
      name: filename,
      type: mimeType
    };

  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Get all lessons from the sheet
 * @returns {Array} Array of lesson objects
 */
function fetchLessonList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Lessons');
  
  if (!sheet) {
    setupSheet(); 
    sheet = ss.getSheetByName('Lessons');
  }
  
  var data = sheet.getDataRange().getValues();
  var lessons = [];
  
  // Skip header row (index 0)
  for (var i = 1; i < data.length; i++) {
    var rawDate = data[i][3];
    var formattedDate = '';
    
    if (rawDate instanceof Date) {
      formattedDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      formattedDate = rawDate;
    }
    
    // Parse JSON for links and files if they exist
    var links = [];
    var files = [];
    try { links = data[i][4] ? JSON.parse(data[i][4]) : []; } catch(e) {}
    try { files = data[i][5] ? JSON.parse(data[i][5]) : []; } catch(e) {}

    lessons.push({
      id: data[i][0],
      topic: data[i][1],
      content: data[i][2],
      date_created: formattedDate,
      links: links,
      files: files
    });
  }
  
  return lessons;
}

/**
 * Save a new lesson
 * @param {Object} lessonData {topic, content, links, files}
 * @returns {Object} Result status
 */
function saveLesson(lessonData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Lessons');
    
    var id = Utilities.getUuid();
    var date = new Date();
    var linksJson = JSON.stringify(lessonData.links || []);
    var filesJson = JSON.stringify(lessonData.files || []);
    
    sheet.appendRow([id, lessonData.topic, lessonData.content, date, linksJson, filesJson]);
    
    return { success: true, message: 'Lesson saved successfully' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Update an existing lesson
 * @param {Object} lessonData {id, topic, content, links, files}
 * @returns {Object} Result status
 */
function updateLesson(lessonData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Lessons');
    var data = sheet.getDataRange().getValues();
    
    var rowIndex = -1;
    
    // Find row by ID
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == lessonData.id) {
        rowIndex = i + 1; // 1-based index
        break;
      }
    }
    
    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 2).setValue(lessonData.topic);
      sheet.getRange(rowIndex, 3).setValue(lessonData.content);
      // Update links and files
      sheet.getRange(rowIndex, 5).setValue(JSON.stringify(lessonData.links || []));
      sheet.getRange(rowIndex, 6).setValue(JSON.stringify(lessonData.files || []));
      
      return { success: true, message: 'Lesson updated successfully' };
    } else {
      return { success: false, message: 'Lesson ID not found' };
    }
    
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Delete a lesson
 * @param {String} id Lesson ID
 * @returns {Object} Result status
 */
function deleteLesson(id) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Lessons');
    var data = sheet.getDataRange().getValues();
    
    var rowIndex = -1;
    
    // Find row by ID
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        rowIndex = i + 1; // 1-based index
        break;
      }
    }
    
    if (rowIndex > 0) {
      sheet.deleteRow(rowIndex);
      return { success: true, message: 'Lesson deleted successfully' };
    } else {
      return { success: false, message: 'Lesson ID not found' };
    }
    
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Check Admin Credentials against generic 'admin' sheet
 * @param {String} username
 * @param {String} password
 * @returns {Object} success boolean
 */
function checkAdmin(username, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('admin');
  
  if (!sheet) {
    // If sheet doesn't exist, try setup or fail safe
    setupSheet();
    sheet = ss.getSheetByName('admin');
  }
  
  var data = sheet.getDataRange().getValues();
  
  // Skip header (row 0)
  for (var i = 1; i < data.length; i++) {
    var storedUser = data[i][0];
    var storedPass = data[i][1];
    
    // Simple string comparison (In production, hash passwords!)
    if (String(storedUser) === String(username) && String(storedPass) === String(password)) {
      return { success: true };
    }
  }
  
  return { success: false };
}

/**
 * Get Admin Profile (Username only for security, or both if needed for edit)
 */
function getAdminProfile() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('admin');
  var data = sheet.getDataRange().getValues();
  // Assume first admin for simplicity
  if(data.length > 1) {
    return { success: true, username: data[1][0], password: data[1][1] };
  }
  return { success: false };
}

/**
 * Update Admin Profile
 */
function updateAdminProfile(oldUsername, newUsername, newPassword) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('admin');
  var data = sheet.getDataRange().getValues();
  
  for(var i=1; i<data.length; i++) {
    if(String(data[i][0]) === String(oldUsername)) {
      sheet.getRange(i+1, 1).setValue(newUsername);
      sheet.getRange(i+1, 2).setValue(newPassword);
      return { success: true, message: 'บันทึกข้อมูลเรียบร้อยแล้ว' };
    }
  }
  return { success: false, message: 'ไม่พบผู้ใช้' };
}

// =========================================
// STUDENT MANAGEMENT
// =========================================

function loginStudent(username, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('students');
  if(!sheet) return { success: false, message: 'System not ready' };
  
  var data = sheet.getDataRange().getValues();
  for(var i=1; i<data.length; i++) {
    if(String(data[i][0]) === String(username) && String(data[i][1]) === String(password)) {
      return { success: true, fullname: data[i][2] };
    }
  }
  return { success: false, message: 'Invalid credentials' };
}

function getStudents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('students');
  if(!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var students = [];
  for(var i=1; i<data.length; i++) {
    students.push({
      username: data[i][0],
      password: data[i][1],
      fullname: data[i][2]
    });
  }
  return students;
}

function saveStudent(studentData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('students');
  // Check duplicate
  var data = sheet.getDataRange().getValues();
  for(var i=1; i<data.length; i++) {
    if(data[i][0] == studentData.username) {
      return { success: false, message: 'Username already exists' };
    }
  }
  sheet.appendRow([studentData.username, studentData.password, studentData.fullname]);
  return { success: true };
}

function updateStudent(studentData) { // Re-using save/update logic pattern if needed, but usually Delete+Add or Edit
   // For simplicity in this turn, we'll assume Edit finds by username
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var sheet = ss.getSheetByName('students');
   var data = sheet.getDataRange().getValues();
   for(var i=1; i<data.length; i++) {
     if(String(data[i][0]) === String(studentData.origin_username || studentData.username)) {
       sheet.getRange(i+1, 1).setValue(studentData.username);
       sheet.getRange(i+1, 2).setValue(studentData.password);
       sheet.getRange(i+1, 3).setValue(studentData.fullname);
       return { success: true };
     }
   }
   return { success: false, message: 'Student not found' };
}

function deleteStudent(username) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('students');
  var data = sheet.getDataRange().getValues();
  for(var i=1; i<data.length; i++) {
    if(String(data[i][0]) === String(username)) {
      sheet.deleteRow(i+1);
      return { success: true };
    }
  }
  return { success: false, message: 'Not found' };
}

// =========================================
// EXAM MANAGEMENT (PRE/POST)
// =========================================

function getExams(type) {
  // type = 'pre' or 'post'
  var sheetName = type === 'pre' ? 'exams_pre' : 'exams_post';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if(!sheet) { setupSheet(); sheet = ss.getSheetByName(sheetName); }
  
  var data = sheet.getDataRange().getValues();
  var exams = [];
  for(var i=1; i<data.length; i++) {
    exams.push({
      id: data[i][0],
      question: data[i][1],
      options: JSON.parse(data[i][2] || '[]'),
      correct_answer: data[i][3],
      image_url: data[i][4]
    });
  }
  return exams;
}

function saveExam(type, examData) {
  var sheetName = type === 'pre' ? 'exams_pre' : 'exams_post';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  var id = Utilities.getUuid();
  sheet.appendRow([
    id, 
    examData.question, 
    JSON.stringify(examData.options), 
    examData.correct_answer, 
    examData.image_url
  ]);
  return { success: true };
}

function updateExam(type, examData) {
  var sheetName = type === 'pre' ? 'exams_pre' : 'exams_post';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  
  for(var i=1; i<data.length; i++) {
    if(String(data[i][0]) === String(examData.id)) {
      sheet.getRange(i+1, 2).setValue(examData.question);
      sheet.getRange(i+1, 3).setValue(JSON.stringify(examData.options));
      sheet.getRange(i+1, 4).setValue(examData.correct_answer);
      sheet.getRange(i+1, 5).setValue(examData.image_url);
      return { success: true };
    }
  }
  return { success: false };
}

function deleteExam(type, id) {
  var sheetName = type === 'pre' ? 'exams_pre' : 'exams_post';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  
  for(var i=1; i<data.length; i++) {
    if(String(data[i][0]) === String(id)) {
      sheet.deleteRow(i+1);
      return { success: true };
    }
  }
  return { success: false };
}

// =========================================
// SCORING SYSTEM
// =========================================

function submitExam(studentUsername, type, studentAnswers) {
  // studentAnswers = { examId: selectedOption, ... }
  var sheetName = type === 'pre' ? 'exams_pre' : 'exams_post';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  
  var score = 0;
  var total = 0;
  
  // Create a map of Correct Answers
  var keyMap = {};
  for(var i=1; i<data.length; i++) {
    keyMap[data[i][0]] = data[i][3]; // ID -> Correct Answer
    total++;
  }
  
  // Calculate Score
  for (var qId in studentAnswers) {
    if (keyMap[qId] && keyMap[qId] == studentAnswers[qId]) {
      score++;
    }
  }
  
  // Save Score
  var scoreSheet = ss.getSheetByName('scores');
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  scoreSheet.appendRow([studentUsername, type, score, timestamp]);
  
  return { success: true, score: score, total: total };
}

function getExamResults(type) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('scores');
  if(!sheet) return [];
  
  // Get all scores
  var data = sheet.getDataRange().getValues();
  var results = [];
  
  // Get Student Names map
  var studentSheet = ss.getSheetByName('students');
  var studentData = studentSheet ? studentSheet.getDataRange().getValues() : [];
  var studentNameMap = {};
  for(var j=1; j<studentData.length; j++) {
    studentNameMap[studentData[j][0]] = studentData[j][2];
  }
  
  for(var i=1; i<data.length; i++) {
    var rowType = String(data[i][1] || '').trim().toLowerCase();
    var targetType = String(type || '').trim().toLowerCase();
    
    if(rowType === targetType) {
      var ts = data[i][3];
      var formattedTs = ts instanceof Date ? Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : String(ts || '');
      
      results.push({
        username: data[i][0],
        fullname: studentNameMap[data[i][0]] || data[i][0],
        score: Number(data[i][2]) || 0,
        timestamp: formattedTs
      });
    }
  }
  
  // Sort by Score Descending
  results.sort(function(a, b) { return b.score - a.score; });
  
  // Add Rank
  return results.map(function(item, index) { 
    item.rank = index + 1; 
    return item;
  });
}
