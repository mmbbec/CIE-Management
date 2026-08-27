// =====================================================
// HOD.GS – Complete HOD Backend Functions
// =====================================================

/****************************************************
 * ADD SUBJECT (HOD only)
 ****************************************************/
function addSubjectByHOD(data) {
  console.log("addSubjectByHOD called with data:", data);
  if (!data) throw new Error("No data received.");
  
  const department = String(data.Department || "").trim();
  const academicYear = String(data.Academic_Year || "").trim();
  const semester = String(data.Semester || "").trim();
  const subjectCode = String(data.Subject_Code || "").trim();
  const subjectName = String(data.Subject_Name || "").trim();
  const section = String(data.Section || "").trim();
  const strength = Number(data.Strength) || 0;

  if (!department || !academicYear || !semester || !subjectCode || !subjectName)
    throw new Error("All fields (Department, Academic Year, Semester, Subject Code, Subject Name) are required.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Subject_Registration");
  if (!sheet) throw new Error("Subject_Registration sheet not found. Run setupExaminationERP first.");

  const dataRange = sheet.getDataRange().getValues();
  const regId = "REG" + Utilities.formatString("%04d", dataRange.length);

  // Check duplicate
  for (let i = 1; i < dataRange.length; i++) {
    const row = dataRange[i];
    if (String(row[3]) === department &&
        String(row[4]) === subjectCode &&
        String(row[2]) === semester &&
        String(row[1]) === academicYear) {
      throw new Error("This subject already exists for the department.");
    }
  }

  sheet.appendRow([
    regId,
    academicYear,
    semester,
    department,
    subjectCode,
    subjectName,
    section,
    strength,
    "Active"
  ]);

  return { success: true, message: "Subject added successfully." };
}

/****************************************************
 * GET SUBJECTS FOR DEPARTMENT
 ****************************************************/
function getDepartmentSubjects(department) {
  console.log("getDepartmentSubjects called for:", department);
  if (!department) throw new Error("Department is required.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Subject_Registration");
  if (!sheet) throw new Error("Subject_Registration sheet not found.");

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const deptCol = headers.indexOf("Department");
  const statusCol = headers.indexOf("Status");
  const subjects = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[deptCol]).trim().toUpperCase() === department.toUpperCase() &&
        String(row[statusCol]).trim().toLowerCase() === "active") {
      subjects.push({
        Registration_ID: row[headers.indexOf("Registration_ID")],
        Academic_Year: row[headers.indexOf("Academic_Year")],
        Semester: row[headers.indexOf("Semester")],
        Subject_Code: row[headers.indexOf("Subject_Code")],
        Subject_Name: row[headers.indexOf("Subject_Name")],
        Section: row[headers.indexOf("Section")],
        Strength: row[headers.indexOf("Strength")]
      });
    }
  }
  return subjects;
}

/****************************************************
 * ADD FACULTY (HOD only)
 ****************************************************/
function addFacultyByHOD(data) {
  console.log("addFacultyByHOD called with data:", data);
  if (!data) throw new Error("No data received.");

  const name = String(data.Staff_Name || "").trim();
  const designation = String(data.Designation || "").trim();
  const department = String(data.Department || "").trim();
  const experience = Number(data.Experience) || 0;
  const email = String(data.Email || "").trim();
  const mobile = String(data.Mobile || "").trim();

  if (!name || !designation || !department || !email)
    throw new Error("Name, Designation, Department, and Email are required.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Staff_Master");
  if (!sheet) throw new Error("Staff_Master sheet not found. Run setupExaminationERP first.");

  const dataRange = sheet.getDataRange().getValues();
  const staffId = "STF" + Utilities.formatString("%04d", dataRange.length);

  // Check duplicate email
  const headers = dataRange[0];
  const emailCol = headers.indexOf("Email");
  for (let i = 1; i < dataRange.length; i++) {
    if (String(dataRange[i][emailCol]).trim().toLowerCase() === email.toLowerCase()) {
      throw new Error("A staff member with this email already exists.");
    }
  }

  // Determine eligibility based on designation
  const canInvigilate = (designation.toLowerCase().includes("professor") || designation.toLowerCase().includes("hod")) ? "Yes" : "No";
  const canSupervise = (designation.toLowerCase().includes("professor") || designation.toLowerCase().includes("hod") || designation.toLowerCase().includes("dean")) ? "Yes" : "No";
  const canSquad = "Yes";

  sheet.appendRow([
    staffId,
    name,
    department,
    designation,
    experience,
    email,
    mobile,
    canInvigilate,
    canSupervise,
    canSquad,
    "Active"
  ]);

  return { success: true, message: "Faculty added successfully." };
}

/****************************************************
 * GET FACULTY FOR DEPARTMENT
 ****************************************************/
function getDepartmentFaculty(department) {
  console.log("getDepartmentFaculty called for:", department);
  if (!department) throw new Error("Department is required.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Staff_Master");
  if (!sheet) throw new Error("Staff_Master sheet not found.");

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const deptCol = headers.indexOf("Department");
  const statusCol = headers.indexOf("Status");
  const faculty = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[deptCol]).trim().toUpperCase() === department.toUpperCase() &&
        String(row[statusCol]).trim().toLowerCase() === "active") {
      faculty.push({
        Staff_ID: row[headers.indexOf("Staff_ID")],
        Staff_Name: row[headers.indexOf("Staff_Name")],
        Designation: row[headers.indexOf("Designation")],
        Experience_Years: row[headers.indexOf("Experience_Years")],
        Email: row[headers.indexOf("Email")],
        Mobile: row[headers.indexOf("Mobile")]
      });
    }
  }
  return faculty;
}

/****************************************************
 * ADD SEMESTER TIMETABLE (HOD)
 ****************************************************/
function addSemesterTimetableByHOD(data) {
  console.log("addSemesterTimetableByHOD called with data:", data);
  if (!data) throw new Error("No data received.");

  const semester = String(data.Semester || "").trim();
  const day = String(data.Day || "").trim();
  const time = String(data.Time || "").trim();
  const subjectCode = String(data.Subject_Code || "").trim();
  const department = String(data.Department || "").trim();
  const examID = String(data.Exam_ID || "").trim();

  if (!semester || !day || !time || !subjectCode || !department)
    throw new Error("Semester, Day, Time, Subject Code, and Department are required.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Semester_Timetable");
  if (!sheet) throw new Error("Semester_Timetable sheet not found.");

  const dataRange = sheet.getDataRange().getValues();
  const ttId = "ST" + Utilities.formatString("%04d", dataRange.length);

  // Check duplicate
  for (let i = 1; i < dataRange.length; i++) {
    const row = dataRange[i];
    if (String(row[2]) === semester &&
        String(row[3]) === day &&
        String(row[5]) === time &&
        String(row[6]) === subjectCode &&
        String(row[8]) === department) {
      throw new Error("This timetable entry already exists.");
    }
  }

  const dayNum = parseInt(day.replace(/\D/g, "")) || 0;
  const timeNum = parseInt(time.replace(/\D/g, "")) || 0;
  const slotNo = dayNum * 10 + timeNum;

  sheet.appendRow([
    ttId,
    examID || "",
    semester,
    day,
    slotNo,
    time,
    subjectCode,
    "",
    department
  ]);

  return { success: true, message: "Timetable entry added successfully." };
}

/****************************************************
 * GET SEMESTER TIMETABLE FOR DEPARTMENT
 ****************************************************/
function getSemesterTimetable(semester, department) {
  console.log("getSemesterTimetable called with semester:", semester, "department:", department);
  if (!semester || !department) throw new Error("Semester and Department are required.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Semester_Timetable");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const semCol = headers.indexOf("Semester");
  const deptCol = headers.indexOf("Department");
  const dayCol = headers.indexOf("Day");
  const timeCol = headers.indexOf("Time");
  const subjectCodeCol = headers.indexOf("Subject_Code");
  const results = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[semCol]).trim() === String(semester).trim() &&
        String(row[deptCol]).trim().toUpperCase() === department.toUpperCase()) {
      results.push({
        Day: row[dayCol],
        Time: row[timeCol],
        Subject_Code: row[subjectCodeCol]
      });
    }
  }
  return results;
}

/****************************************************
 * SAVE TIMETABLE SLOT (Add or Update)
 ****************************************************/
function saveTimetableSlot(semester, day, time, subjectCode, department) {
  console.log("saveTimetableSlot called:", {semester, day, time, subjectCode, department});
  if (!semester || !day || !time || !department) {
    throw new Error("Semester, Day, Time, and Department are required.");
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Semester_Timetable");
  if (!sheet) throw new Error("Semester_Timetable sheet not found.");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const semCol = headers.indexOf("Semester");
  const deptCol = headers.indexOf("Department");
  const dayCol = headers.indexOf("Day");
  const timeCol = headers.indexOf("Time");
  const subjectCodeCol = headers.indexOf("Subject_Code");
  const examIDCol = headers.indexOf("Exam_ID");
  const ttIdCol = headers.indexOf("Timetable_ID");
  const slotNoCol = headers.indexOf("Slot_No");
  const subjectNameCol = headers.indexOf("Subject_Name");

  // Check if entry exists
  let existingRow = -1;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[semCol]).trim() === String(semester).trim() &&
        String(row[dayCol]).trim() === day.trim() &&
        String(row[timeCol]).trim() === time.trim() &&
        String(row[deptCol]).trim().toUpperCase() === department.toUpperCase()) {
      existingRow = i;
      break;
    }
  }

  if (existingRow !== -1) {
    // Update existing
    if (subjectCode && subjectCode.trim() !== "") {
      sheet.getRange(existingRow+1, subjectCodeCol+1).setValue(subjectCode.trim());
      return { success: true, message: "Timetable updated successfully." };
    } else {
      // Delete the row
      sheet.deleteRow(existingRow+1);
      return { success: true, message: "Timetable entry removed." };
    }
  } else {
    // Insert new
    if (!subjectCode || subjectCode.trim() === "") {
      throw new Error("Subject Code is required for new entry.");
    }
    const ttId = "ST" + Utilities.formatString("%04d", data.length);
    const dayNum = parseInt(day.replace(/\D/g, "")) || 0;
    const timeNum = parseInt(time.replace(/\D/g, "")) || 0;
    const slotNo = dayNum * 10 + timeNum;
    const newRow = [];
    headers.forEach(function(header) {
      switch (header) {
        case "Timetable_ID": newRow.push(ttId); break;
        case "Exam_ID": newRow.push(""); break;
        case "Semester": newRow.push(semester); break;
        case "Day": newRow.push(day); break;
        case "Slot_No": newRow.push(slotNo); break;
        case "Time": newRow.push(time); break;
        case "Subject_Code": newRow.push(subjectCode.trim()); break;
        case "Subject_Name": newRow.push(""); break;
        case "Department": newRow.push(department); break;
        default: newRow.push("");
      }
    });
    sheet.appendRow(newRow);
    return { success: true, message: "Timetable added successfully." };
  }
}

/****************************************************
 * DELETE TIMETABLE SLOT
 ****************************************************/
function deleteTimetableSlot(semester, day, time, department) {
  return saveTimetableSlot(semester, day, time, "", department);
}

/****************************************************
 * GENERATE DEPARTMENT TIMETABLE PRINTOUT
 ****************************************************/
function generateDepartmentTimetablePrintout(semester, department) {
  if (!semester || !department) throw new Error("Semester and Department are required.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Semester_Timetable");
  if (!sheet) throw new Error("Semester_Timetable sheet not found.");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const semCol = headers.indexOf("Semester");
  const deptCol = headers.indexOf("Department");
  const dayCol = headers.indexOf("Day");
  const timeCol = headers.indexOf("Time");
  const subjectCodeCol = headers.indexOf("Subject_Code");
  const subjectNameCol = headers.indexOf("Subject_Name");

  // Build lookup for subjects: code -> name
  const regSheet = ss.getSheetByName("Subject_Registration");
  const subjectNameLookup = {};
  if (regSheet) {
    const regData = regSheet.getDataRange().getValues();
    if (regData.length > 1) {
      const regHeaders = regData[0];
      const codeIdx = regHeaders.indexOf("Subject_Code");
      const nameIdx = regHeaders.indexOf("Subject_Name");
      const deptIdx = regHeaders.indexOf("Department");
      for (let i = 1; i < regData.length; i++) {
        const row = regData[i];
        if (String(row[deptIdx]).trim().toUpperCase() === department.toUpperCase()) {
          subjectNameLookup[String(row[codeIdx]).trim()] = String(row[nameIdx]).trim();
        }
      }
    }
  }

  const days = ["Day 1", "Day 2", "Day 3"];
  const times = ["8:00–9:30", "10:00–11:30", "12:00–1:30", "2:00–3:30", "4:00–5:30"];
  const lookup = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[semCol]).trim() === String(semester).trim() &&
        String(row[deptCol]).trim().toUpperCase() === department.toUpperCase()) {
      const key = String(row[dayCol]).trim() + "|" + String(row[timeCol]).trim();
      lookup[key] = String(row[subjectCodeCol]).trim();
    }
  }

  const sheetName = "Timetable_Print_" + department + "_Sem" + semester;
  let printSheet = ss.getSheetByName(sheetName);
  if (printSheet) ss.deleteSheet(printSheet);
  printSheet = ss.insertSheet(sheetName);

  printSheet.getRange("A1").setValue("DEPARTMENT TIMETABLE");
  printSheet.getRange("A1").setFontSize(16).setFontWeight("bold");
  printSheet.getRange("A2").setValue("Department: " + department + "  |  Semester: " + semester);
  printSheet.getRange("A2").setFontWeight("bold");

  const startRow = 4;
  let row = startRow;
  printSheet.getRange(row, 1).setValue("Day / Time");
  times.forEach(function(t, idx) {
    printSheet.getRange(row, idx+2).setValue(t);
  });
  printSheet.getRange(row, 1, 1, times.length+1).setFontWeight("bold").setHorizontalAlignment("center");
  row++;

  days.forEach(function(day) {
    printSheet.getRange(row, 1).setValue(day);
    times.forEach(function(time, idx) {
      const key = day + "|" + time;
      const code = lookup[key] || "";
      const name = subjectNameLookup[code] || "";
      printSheet.getRange(row, idx+2).setValue(code + (name ? " " + name : ""));
    });
    row++;
  });

  printSheet.getRange(startRow, 1, 3, times.length+1).setBorder(true, true, true, true, true, true);
  printSheet.autoResizeColumns(1, times.length+1);
  printSheet.setFrozenRows(startRow);

  SpreadsheetApp.getUi().alert(
    "Timetable printout generated!\n\n" +
    "Open the sheet: " + sheetName
  );
}

/****************************************************
 * GET DEPARTMENT DUTIES (Optional)
 ****************************************************/
function getDepartmentDuties(department) {
  if (!department) throw new Error("Department is required.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CIE_Final_Duty_List");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const deptCol = headers.indexOf("Staff_Department");
  if (deptCol === -1) return [];
  const duties = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[deptCol]).trim().toUpperCase() === department.toUpperCase()) {
      duties.push({
        Day: row[headers.indexOf("Day")],
        Slot_ID: row[headers.indexOf("Slot_ID")],
        Time: row[headers.indexOf("Time")],
        Duty_Type: row[headers.indexOf("Duty_Type")],
        Room_ID: row[headers.indexOf("Room_ID")],
        Room_Name: row[headers.indexOf("Room_Name")],
        Staff_Name: row[headers.indexOf("Staff_Name")],
        Designation: row[headers.indexOf("Designation")],
        Status: row[headers.indexOf("Status")]
      });
    }
  }
  return duties;
}

/****************************************************
 * GENERATE DEPARTMENT DUTY ORDER (Optional)
 ****************************************************/
function generateDepartmentDutyOrder(department) {
  if (!department) throw new Error("Department is required.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CIE_Final_Duty_List");
  if (!sheet) throw new Error("CIE_Final_Duty_List sheet not found.");
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error("No duty data found.");
  const headers = data[0];
  const deptCol = headers.indexOf("Staff_Department");
  if (deptCol === -1) throw new Error("Staff_Department column not found.");
  const filtered = data.filter((row, idx) => idx === 0 || String(row[deptCol]).trim().toUpperCase() === department.toUpperCase());
  if (filtered.length <= 1) throw new Error("No duties found for this department.");
  const sheetName = "HOD_Duty_Order_" + department;
  let outputSheet = ss.getSheetByName(sheetName);
  if (outputSheet) ss.deleteSheet(outputSheet);
  outputSheet = ss.insertSheet(sheetName);
  outputSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  outputSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  outputSheet.getRange(2, 1, filtered.length - 1, headers.length).setValues(filtered.slice(1));
  outputSheet.autoResizeColumns(1, headers.length);
  outputSheet.setFrozenRows(1);
  SpreadsheetApp.getUi().alert("Duty order generated for " + department + "\n\nOpen the sheet: " + sheetName);
}

/****************************************************
 * GET DEPARTMENT STATS
 ****************************************************/
function getDepartmentStats(department) {
  if (!department) throw new Error("Department is required.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const facultySheet = ss.getSheetByName("Staff_Master");
  const subjectSheet = ss.getSheetByName("Subject_Registration");
  
  let facultyCount = 0;
  if (facultySheet) {
    const data = facultySheet.getDataRange().getValues();
    if (data.length > 1) {
      const headers = data[0];
      const deptCol = headers.indexOf("Department");
      const statusCol = headers.indexOf("Status");
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][deptCol]).trim().toUpperCase() === department.toUpperCase() &&
            String(data[i][statusCol]).trim().toLowerCase() === "active") {
          facultyCount++;
        }
      }
    }
  }
  
  let subjectCount = 0;
  if (subjectSheet) {
    const data = subjectSheet.getDataRange().getValues();
    if (data.length > 1) {
      const headers = data[0];
      const deptCol = headers.indexOf("Department");
      const statusCol = headers.indexOf("Status");
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][deptCol]).trim().toUpperCase() === department.toUpperCase() &&
            String(data[i][statusCol]).trim().toLowerCase() === "active") {
          subjectCount++;
        }
      }
    }
  }
  
  return {
    department: department,
    facultyCount: facultyCount,
    subjectCount: subjectCount
  };
}
