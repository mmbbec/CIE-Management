// =====================================================
// ADMIN.GS – Subject sync, semester/exam management
// =====================================================

/****************************************************
 * SYNC SUBJECT REGISTRATION → SUBJECT STRENGTH
 ****************************************************/
function syncSubjectStrength() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Subject_Registration");
  const targetSheet = ss.getSheetByName("Subject_Strength");
  if (!sourceSheet) throw new Error("Subject_Registration sheet not found.");
  if (!targetSheet) throw new Error("Subject_Strength sheet not found.");

  const sourceData = sourceSheet.getDataRange().getValues();
  if (sourceData.length <= 1) {
    SpreadsheetApp.getUi().alert("No subject registration data found.");
    return;
  }
  const headers = sourceData[0];
  const registrationIDCol = headers.indexOf("Registration_ID");
  const academicYearCol = headers.indexOf("Academic_Year");
  const semesterCol = headers.indexOf("Semester");
  const departmentCol = headers.indexOf("Department");
  const subjectCodeCol = headers.indexOf("Subject_Code");
  const subjectNameCol = headers.indexOf("Subject_Name");
  const sectionCol = headers.indexOf("Section");
  const strengthCol = headers.indexOf("Strength");
  const statusCol = headers.indexOf("Status");

  const requiredColumns = [
    ["Registration_ID", registrationIDCol],
    ["Academic_Year", academicYearCol],
    ["Semester", semesterCol],
    ["Department", departmentCol],
    ["Subject_Code", subjectCodeCol],
    ["Subject_Name", subjectNameCol],
    ["Section", sectionCol],
    ["Strength", strengthCol],
    ["Status", statusCol]
  ];
  requiredColumns.forEach(function(item) {
    if (item[1] === -1) throw new Error("Required column missing: " + item[0]);
  });

  const output = [];
  let counter = 1;
  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i];
    const status = String(row[statusCol]).trim().toLowerCase();
    if (status !== "active") continue;
    const strength = Number(row[strengthCol]) || 0;
    output.push([
      "STR" + String(counter).padStart(4, "0"),
      row[semesterCol],
      row[departmentCol],
      row[subjectCodeCol],
      row[subjectNameCol],
      row[sectionCol],
      strength,
      "Subject Registration",
      "Yes"
    ]);
    counter++;
  }

  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clearContent();
  }
  if (output.length > 0) {
    targetSheet.getRange(2, 1, output.length, output[0].length).setValues(output);
  }
  targetSheet.autoResizeColumns(1, targetSheet.getLastColumn());
  SpreadsheetApp.getUi().alert("Subject Strength Updated Successfully.\n\nActive Subject Records: " + output.length);
}

/****************************************************
 * RESET ADMIN (in case of password issues)
 ****************************************************/
function resetAdmin() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) { SpreadsheetApp.getUi().alert("Users sheet not found."); return; }
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf("Email");
  const passwordCol = headers.indexOf("Password");
  const statusCol = headers.indexOf("Status");
  const roleCol = headers.indexOf("Role");
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][emailCol]).trim().toLowerCase() === "mmbbec@gmail.com") {
      sheet.getRange(i+1, passwordCol+1).setValue("mahadevmb");
      sheet.getRange(i+1, statusCol+1).setValue("ACTIVE");
      sheet.getRange(i+1, roleCol+1).setValue("ADMIN");
      SpreadsheetApp.getUi().alert("Admin reset. Email: mmbbec@gmail.com, Password: mahadevmb");
      return;
    }
  }
  SpreadsheetApp.getUi().alert("Admin user not found. Run setupAdmin() first.");
}

/****************************************************
 * ENSURE EXAMS SHEET EXISTS
 ****************************************************/
function getExamsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Exams");
  if (!sheet) {
    sheet = ss.insertSheet("Exams");
    const headers = ["Exam_ID", "Academic_Year", "Exam_Name", "Semester_Type", "Status"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

/****************************************************
 * ADD SEMESTER
 ****************************************************/
function addSemester(academicYear, semesterNumber, semesterType) {
  if (!academicYear || !semesterNumber || !semesterType)
    throw new Error("Academic Year, Semester Number and Type are required.");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Semesters");
  if (!sheet) throw new Error("Semesters sheet not found. Run setupExaminationERP first.");
  
  const data = sheet.getDataRange().getValues();
  const count = data.length;
  const semesterId = "SEM" + Utilities.formatString("%02d", count);
  
  // Check duplicate
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(semesterNumber) &&
        String(data[i][2]).toUpperCase() === String(semesterType).toUpperCase() &&
        String(data[i][3]) === String(academicYear)) {
      throw new Error("This semester already exists.");
    }
  }
  
  sheet.appendRow([semesterId, semesterNumber, semesterType, academicYear]);
  return { success: true, message: "Semester added successfully." };
}

/****************************************************
 * ADD EXAM
 ****************************************************/
function addExam(academicYear, examName, semesterType) {
  if (!academicYear || !examName || !semesterType)
    throw new Error("Academic Year, Exam Name and Semester Type are required.");
  
  const sheet = getExamsSheet_();
  const data = sheet.getDataRange().getValues();
  const count = data.length;
  const examId = "EXAM" + Utilities.formatString("%04d", count);
  
  // Check duplicate
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(academicYear) &&
        String(data[i][2]).toUpperCase() === String(examName).toUpperCase() &&
        String(data[i][3]).toUpperCase() === String(semesterType).toUpperCase()) {
      throw new Error("This exam already exists for this year and semester type.");
    }
  }
  
  sheet.appendRow([examId, academicYear, examName, semesterType, "Active"]);
  return { success: true, message: "Exam added successfully." };
}
