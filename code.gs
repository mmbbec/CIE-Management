/****************************************************
 * EXAMINATION ERP
 * DATABASE SETUP
 * Academic Year: 2026-27
 ****************************************************/
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('CIE Examination Management')
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}

function setupExaminationERP() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {
    "Exam_Settings": ["Setting", "Value"],
    "Semesters": ["Semester_ID", "Semester", "Semester_Type", "Year"],
    "Departments": ["Department_ID", "Department_Code", "Department_Name", "Status"],
    "Subject_Registration": ["Registration_ID", "Academic_Year", "Semester", "Department", "Subject_Code", "Subject_Name", "Section", "Strength", "Status"],
    "Subject_Strength": ["Strength_ID", "Semester", "Department", "Subject_Code", "Subject_Name", "Section", "Strength", "Source", "Verified"],
    "Exam_Slots": ["Slot_ID", "Day", "Slot_No", "Start_Time", "End_Time", "Display_Time", "Status"],
    "CIE_Timetable": ["Exam_ID", "Academic_Year", "Exam_Name", "Semester", "Department", "Subject_Code", "Subject_Name", "Strength", "Slot_ID", "Day", "Time", "Room_Status", "Timetable_Status"],
    "Semester_Timetable": ["Timetable_ID", "Exam_ID", "Semester", "Day", "Slot_No", "Time", "Subject_Code", "Subject_Name", "Department"],
    "Staff_Master": ["Staff_ID", "Staff_Name", "Department", "Designation", "Experience_Years", "Email", "Mobile", "Eligible_Invigilation", "Eligible_Supervisor", "Eligible_Squad", "Status"],
    "Designation_Rules": ["Rule_ID", "Designation", "Min_Experience", "Invigilation", "Supervisor", "Squad", "Priority"],
    "Room_Master": ["Room_ID", "Room_Name", "Building", "Floor", "Capacity", "Department", "Available", "Remarks"],
    "Question_Paper_Status": ["QP_ID", "Exam_ID", "Semester", "Department", "Subject_Code", "Subject_Name", "Faculty_Name", "Submission_Deadline", "Submitted_Date", "Drive_File_ID", "Status", "Verified_By"],
    "Question_Paper_Pockets": ["Pocket_ID", "Exam_ID", "Day", "Slot", "Time", "Semester", "Subject_Code", "Subject_Name", "Strength", "Question_Paper_Status", "Prepared_By", "Prepared_Date", "Pocket_Status"],
    "Duty_Rules": ["Rule_ID", "Rule_Name", "Value", "Description", "Active"],
    "Duty_Allocation": ["Duty_ID", "Exam_ID", "Day", "Slot_ID", "Time", "Staff_ID", "Staff_Name", "Department", "Designation", "Duty_Type", "Room_ID", "Status", "Notification_Status"],
    "Notifications": ["Notification_ID", "Duty_ID", "Staff_ID", "Staff_Name", "Notification_Type", "Recipient", "Message", "Sent_Date", "Status", "Error"],
    "Exam_Logs": ["Log_ID", "Date_Time", "User", "Action", "Module", "Record_ID", "Details"]
  };

  Object.keys(sheets).forEach(function(sheetName) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    if (sheet.getLastRow() === 0) {
      const headers = sheets[sheetName];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      formatHeader(sheet, headers.length);
    }
  });

  addMasterData(ss);
  SpreadsheetApp.getUi().alert(
    "Examination ERP setup completed successfully.\n\n" +
    "17 Examination ERP sheets are ready."
  );
}

function formatHeader(sheet, columnCount) {
  sheet.getRange(1, 1, 1, columnCount).setFontWeight("bold");
  sheet.getRange(1, 1, 1, columnCount).setHorizontalAlignment("center");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, columnCount);
}

function addMasterData(ss) {
  // Exam Settings
  const settings = ss.getSheetByName("Exam_Settings");
  if (settings.getLastRow() === 1) {
    settings.getRange(2, 1, 7, 2).setValues([
      ["Academic_Year", "2026-27"],
      ["Exam_Type", "CIE"],
      ["Semester_Type", "Odd"],
      ["Exam_Name", "CIE-I"],
      ["Start_Date", ""],
      ["Number_of_Days", 3],
      ["Status", "Draft"]
    ]);
  }

  // Semesters
  const semesters = ss.getSheetByName("Semesters");
  if (semesters.getLastRow() === 1) {
    semesters.getRange(2, 1, 8, 4).setValues([
      ["SEM01", 1, "Odd", 1],
      ["SEM02", 2, "Even", 1],
      ["SEM03", 3, "Odd", 2],
      ["SEM04", 4, "Even", 2],
      ["SEM05", 5, "Odd", 3],
      ["SEM06", 6, "Even", 3],
      ["SEM07", 7, "Odd", 4],
      ["SEM08", 8, "Even", 4]
    ]);
  }

  // Departments
  const departments = ss.getSheetByName("Departments");
  if (departments.getLastRow() === 1) {
    departments.getRange(2, 1, 7, 4).setValues([
      ["DEPT01", "CSE", "Computer Science & Engineering", "Active"],
      ["DEPT02", "ECE", "Electronics & Communication Engineering", "Active"],
      ["DEPT03", "ME", "Mechanical Engineering", "Active"],
      ["DEPT04", "CV", "Civil Engineering", "Active"],
      ["DEPT05", "EEE", "Electrical & Electronics Engineering", "Active"],
      ["DEPT06", "AIML", "Artificial Intelligence & Machine Learning", "Active"],
      ["DEPT07", "MATH", "Mathematics", "Active"]
    ]);
  }

  // Exam Slots
  const slots = ss.getSheetByName("Exam_Slots");
  if (slots.getLastRow() === 1) {
    const slotData = [];
    const times = [
      ["08:00", "09:30", "8:00–9:30"],
      ["10:00", "11:30", "10:00–11:30"],
      ["12:00", "13:30", "12:00–1:30"],
      ["14:00", "15:30", "2:00–3:30"],
      ["16:00", "17:30", "4:00–5:30"]
    ];
    for (let day = 1; day <= 3; day++) {
      for (let slot = 1; slot <= 5; slot++) {
        const t = times[slot - 1];
        slotData.push([
          "D" + day + "S" + slot,
          "Day " + day,
          slot,
          t[0],
          t[1],
          t[2],
          "Active"
        ]);
      }
    }
    slots.getRange(2, 1, slotData.length, 7).setValues(slotData);
  }

  // Designation Rules
  const rules = ss.getSheetByName("Designation_Rules");
  if (rules.getLastRow() === 1) {
    rules.getRange(2, 1, 6, 7).setValues([
      ["R001", "Principal", 0, "No", "Yes", "No", 1],
      ["R002", "Dean", 0, "No", "Yes", "Yes", 2],
      ["R003", "HOD", 0, "No", "Yes", "Yes", 3],
      ["R004", "Professor", 10, "Yes", "Yes", "Yes", 4],
      ["R005", "Associate Professor", 5, "Yes", "Yes", "Yes", 5],
      ["R006", "Assistant Professor", 0, "Yes", "No", "Yes", 6]
    ]);
  }

  // Duty Rules
  const dutyRules = ss.getSheetByName("Duty_Rules");
  if (dutyRules.getLastRow() === 1) {
    dutyRules.getRange(2, 1, 6, 5).setValues([
      ["DR001", "Max_Duties_Per_Day", 2, "Maximum duties for one staff member per day", "Yes"],
      ["DR002", "Prevent_Time_Conflict", "Yes", "Staff cannot have two duties in the same slot", "Yes"],
      ["DR003", "Prevent_Same_Room", "Yes", "Prevent duplicate room allocation", "Yes"],
      ["DR004", "Balance_Duties", "Yes", "Distribute duties among eligible staff", "Yes"],
      ["DR005", "Supervisor_Required", "Yes", "Assign supervisor for each examination session", "Yes"],
      ["DR006", "Squad_Required", "Yes", "Assign squad according to examination rules", "Yes"]
    ]);
  }
}

/****************************************************
 * DESIGNATION SENIORITY HELPER
 ****************************************************/
function getDesignationRank(designation) {
  const d = String(designation).trim().toLowerCase();
  if (d === "dean") return 1;
  if (d === "hod") return 2;
  if (d === "professor") return 3;
  if (d === "associate professor") return 4;
  if (d === "assistant professor") return 5;
  return 99;
}

/****************************************************
 * SETUP ADMIN (for mmbbec@gmail.com)
 ****************************************************/
function setupAdmin() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Users");
  if (!sheet) {
    sheet = ss.insertSheet("Users");
    const headers = [
      "User_ID", "Name", "Email", "Mobile", "Password",
      "Role", "Department", "Designation", "Employee_ID",
      "Status", "Created_Date"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf("Email");
  const statusCol = headers.indexOf("Status");
  const roleCol = headers.indexOf("Role");
  let adminExists = false;
  for (let i = 1; i < data.length; i++) {
    const rowEmail = String(data[i][emailCol] || "").trim().toLowerCase();
    if (rowEmail === "mmbbec@gmail.com") {
      adminExists = true;
      if (statusCol !== -1) sheet.getRange(i+1, statusCol+1).setValue("ACTIVE");
      if (roleCol !== -1) sheet.getRange(i+1, roleCol+1).setValue("ADMIN");
      break;
    }
  }

  if (!adminExists) {
    const newRow = [];
    headers.forEach(function(header) {
      switch (header) {
        case "User_ID": newRow.push("U" + Utilities.formatString("%04d", data.length)); break;
        case "Name": newRow.push("Admin"); break;
        case "Email": newRow.push("mmbbec@gmail.com"); break;
        case "Mobile": newRow.push("9035246580"); break;
        case "Password": newRow.push("mahadevmb"); break;
        case "Role": newRow.push("ADMIN"); break;
        case "Department": newRow.push("Administration"); break;
        case "Designation": newRow.push("Administrator"); break;
        case "Status": newRow.push("ACTIVE"); break;
        case "Created_Date": newRow.push(new Date()); break;
        default: newRow.push("");
      }
    });
    sheet.appendRow(newRow);
  }

  SpreadsheetApp.getUi().alert(
    "Admin user setup completed.\n\n" +
    "Email: mmbbec@gmail.com\n" +
    "Password: mahadevmb\n" +
    "Status: ACTIVE\n" +
    "Role: ADMIN"
  );
}

/****************************************************
 * CREATE TEST HOD USER
 ****************************************************/
function createTestHOD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Users");
  
  if (!sheet) {
    sheet = ss.insertSheet("Users");
    const headers = [
      "User_ID", "Name", "Email", "Mobile", "Password",
      "Role", "Department", "Designation", "Employee_ID",
      "Status", "Created_Date"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf("Email");
  
  for (let i = 1; i < data.length; i++) {
    const email = String(data[i][emailCol] || "").trim().toLowerCase();
    if (email === "hod@test.edu") {
      SpreadsheetApp.getUi().alert("Test HOD already exists.\n\nEmail: hod@test.edu\nPassword: hod123\nDepartment: CSE");
      return;
    }
  }

  const newRow = [];
  headers.forEach(function(header) {
    switch (header) {
      case "User_ID": newRow.push("U" + Utilities.formatString("%04d", data.length)); break;
      case "Name": newRow.push("Test HOD"); break;
      case "Email": newRow.push("hod@test.edu"); break;
      case "Mobile": newRow.push("9000000001"); break;
      case "Password": newRow.push("hod123"); break;
      case "Role": newRow.push("HOD"); break;
      case "Department": newRow.push("CSE"); break;
      case "Designation": newRow.push("HOD"); break;
      case "Employee_ID": newRow.push("HOD001"); break;
      case "Status": newRow.push("ACTIVE"); break;
      case "Created_Date": newRow.push(new Date()); break;
      default: newRow.push("");
    }
  });
  sheet.appendRow(newRow);

  SpreadsheetApp.getUi().alert(
    "✅ Test HOD created successfully!\n\n" +
    "Email: hod@test.edu\n" +
    "Password: hod123\n" +
    "Department: CSE\n\n" +
    "You can now log in as HOD to test the system."
  );
}
