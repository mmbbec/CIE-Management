// =====================================================
// AUTH.GS – User signup, login, and admin creation
// =====================================================

/****************************************************
 * ENSURE USERS SHEET EXISTS
 ****************************************************/
function getUsersSheet_() {
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
  return sheet;
}

/****************************************************
 * SIGNUP USER
 ****************************************************/
function signupUser(data) {
  if (!data || typeof data !== 'object')
    throw new Error("No signup data received.");
  const sheet = getUsersSheet_();
  const name = String(data.name || "").trim();
  const email = String(data.email || "").trim().toLowerCase();
  const mobile = String(data.mobile || "").trim();
  const password = String(data.password || "");
  const role = String(data.role || "FACULTY").trim().toUpperCase();
  const department = String(data.department || "").trim();
  const designation = String(data.designation || "").trim();
  const employeeId = String(data.employeeId || "").trim();

  if (!name || !email || !password || !department)
    throw new Error("Name, Email, Password and Department are required.");

  if (role === "ADMIN")
    throw new Error("ADMIN account cannot be created through signup.");

  const values = sheet.getDataRange().getValues();
  if (values.length === 0) throw new Error("Users sheet has no header row.");
  const headers = values[0];
  const emailCol = headers.indexOf("Email");
  if (emailCol === -1) throw new Error("Users sheet must contain an Email column.");

  for (let i = 1; i < values.length; i++) {
    const existingEmail = String(values[i][emailCol] || "").trim().toLowerCase();
    if (existingEmail === email)
      throw new Error("An account with this email already exists.");
  }

  const userId = "U" + Utilities.formatString("%04d", Math.max(values.length, 1));
  const row = [];
  headers.forEach(function(header) {
    switch (header) {
      case "User_ID": row.push(userId); break;
      case "Name": row.push(name); break;
      case "Email": row.push(email); break;
      case "Mobile": row.push(mobile); break;
      case "Password": row.push(password); break;
      case "Role": row.push(role); break;
      case "Department": row.push(department); break;
      case "Designation": row.push(designation); break;
      case "Employee_ID": row.push(employeeId); break;
      case "Status": row.push("Pending"); break;
      case "Created_Date": row.push(new Date()); break;
      default: row.push("");
    }
  });
  sheet.appendRow(row);
  return {
    success: true,
    message: "Registration successful. Your account is waiting for Admin approval."
  };
}

/****************************************************
 * LOGIN
 ****************************************************/
function loginUser(email, password) {
  const sheet = getUsersSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) throw new Error("No users found.");
  const headers = values[0];
  const col = {};
  headers.forEach(function(header, index) { col[header] = index; });
  if (col.Email === undefined || col.Password === undefined || col.Role === undefined)
    throw new Error("Users sheet must contain Email, Password and Role columns.");

  const loginEmail = String(email || "").trim().toLowerCase();
  const loginPassword = String(password || "");

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const userEmail = String(row[col.Email] || "").trim().toLowerCase();
    const userPassword = String(row[col.Password] || "");
    if (userEmail === loginEmail && userPassword === loginPassword) {
      const status = String(row[col.Status] || "").trim().toUpperCase();
      if (status !== "ACTIVE") {
        if (status === "PENDING")
          throw new Error("Your account is waiting for Admin approval.");
        throw new Error("Your account is not active.");
      }
      return {
        success: true,
        userId: row[col.User_ID] || "",
        name: row[col.Name] || "",
        email: row[col.Email] || "",
        role: String(row[col.Role] || "").trim().toUpperCase(),
        department: row[col.Department] || "",
        designation: row[col.Designation] || ""
      };
    }
  }
  throw new Error("Invalid email or password.");
}
