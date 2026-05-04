// ---- CONFIG ----
const ADMIN_EMAILS = "abdullah@axicom.my,idris.ahad@axicom.my,muhammad.idzham@axicom.my";
const CC_EMAILS = "salahuddin@axicom.my";
const SHEET_ID = "12dmZ3Yx9C5z_xNqU-B4xz389wsZYqe9RSADSBM4fhCw"; 
const BOOKING_SHEET_NAME = "Logs";
const MAINT_SHEET_NAME = "Maintenance";
const TIMEZONE = "Asia/Kuala_Lumpur";

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    var result = processApproval(e.parameter.id, e.parameter.action);
    var color = result.includes("Approved") ? "#10B981" : (result.includes("Rejected") ? "#EF4444" : "#6B7280");
    return HtmlService.createHtmlOutput(
      `<div style="font-family:sans-serif; text-align:center; margin-top:50px; padding:20px;">
         <h1 style="color:${color}">${result}</h1>
         <p>The requester has been notified. You can close this window now.</p>
         <script>setTimeout(function(){window.close()},3000);</script>
       </div>`
    ).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  
  // Single file serves everything
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("AxiCom Fleet")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function openSheet(name) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === BOOKING_SHEET_NAME) {
      sheet.appendRow(['ID', 'Timestamp', 'Form Type', 'Status', 'Processed By', 'Processed Time', 'name', 'email', 'phone', 'department', 'vehicle', 'departure', 'return', 'destination', 'purpose', 'health']);
    } else if (name === MAINT_SHEET_NAME) {
      sheet.appendRow(['Timestamp', 'Vehicle', 'Odometer', 'Fuel', 'Clean Ext', 'Clean Int', 'Lights', 'Tyres', 'Issues']);
    }
  }
  return sheet;
}

function getColMap(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  headers.forEach((h, i) => map[h] = i);
  return map;
}

function formatPhone(raw) {
  if (!raw) return "";
  var str = String(raw).replace(/\D/g, ''); 
  if (str.length >= 10) return str.substring(0, 3) + "-" + str.substring(3);
  return str;
}

function getNextRequestId(sheet) {
  var data = sheet.getDataRange().getValues();
  var maxNum = 0;
  for (var i = 1; i < data.length; i++) {
    var id = data[i][0];
    if (typeof id === 'string' && id.startsWith('ASB')) {
      var num = parseInt(id.replace('ASB', ''), 10);
      if (!isNaN(num) && num > maxNum) maxNum = num;
    }
  }
  return 'ASB' + ('000' + (maxNum + 1)).slice(-3);
}

function isVehicleAvailable(sheet, targetVehicle, start, end, excludeId) {
  var data = sheet.getDataRange().getValues();
  var map = getColMap(sheet);
  var targetStart = new Date(start);
  var targetEnd = new Date(end);
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowId = String(row[map['ID']]);
    var status = String(row[map['Status']] || "").toUpperCase();
    if (excludeId && rowId === String(excludeId)) continue;
    if (status === "CANCELLED" || status === "REJECTED") continue;
    var vehicle = row[map['vehicle']];
    var dep = new Date(row[map['departure']]);
    var ret = new Date(row[map['return']]);
    if (vehicle === targetVehicle && !isNaN(dep) && !isNaN(ret)) {
      if (targetStart < ret && targetEnd > dep) return false;
    }
  }
  return true;
}

function submitForm(formData) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return {success: false, message: "Server busy."}; }
  try {
    var sheet = openSheet(BOOKING_SHEET_NAME);
    var data = formData.data;
    if (new Date(data.return) <= new Date(data.departure)) return {success: false, message: "Return time must be after Departure."};
    if (!isVehicleAvailable(sheet, data.vehicle, data.departure, data.return, null)) return {success: false, message: "❌ " + data.vehicle + " is busy."};
    var id = getNextRequestId(sheet);
    var now = new Date();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var row = headers.map(h => {
      if (h === 'ID') return id;
      if (h === 'Timestamp') return now;
      if (h === 'Form Type') return "VBF";
      if (h === 'Status') return "PENDING";
      if (h === 'phone') return formatPhone(data[h]);
      var val = data[h] !== undefined ? data[h] : "";
      return (typeof val === 'string') ? val.toUpperCase() : val;
    });
    sheet.appendRow(row);
    var emailData = {};
    for (var key in data) {
      if (key === 'phone') emailData[key] = formatPhone(data[key]);
      else emailData[key] = (typeof data[key] === 'string') ? data[key].toUpperCase() : data[key];
    }
    sendApprovalEmail(id, emailData);
    return {success: true, message: 'Booking Submitted: ' + id};
  } catch (err) { return {success: false, message: 'Error: ' + err.message}; } 
  finally { lock.releaseLock(); }
}

function submitMaintenance(formObj) {
  try {
    var sheet = openSheet(MAINT_SHEET_NAME);
    var d = formObj.data;
    sheet.appendRow([
      new Date(), String(d.check_vehicle || "").toUpperCase(), d.odometer, d.fuel,
      d.clean_ext ? "OK" : "CHECK", d.clean_int ? "OK" : "CHECK",
      d.lights ? "OK" : "CHECK", d.tyres ? "OK" : "CHECK",
      String(d.issues || "NONE").toUpperCase()
    ]);
    return {success: true, message: "Inspection Logged."};
  } catch (e) { return {success: false, message: e.message}; }
}

function editBooking(id, newData) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return {success: false, message: "System busy."}; }
  try {
    var sheet = openSheet(BOOKING_SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    var map = getColMap(sheet);
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][map['ID']]) === String(id)) { rowIndex = i; break; }
    }
    if (rowIndex === -1) return {success: false, message: "ID not found."};
    var row = data[rowIndex];
    if (row[map['Status']] !== "PENDING") return {success: false, message: "Can only edit Pending requests."};
    if (!isVehicleAvailable(sheet, newData.vehicle, newData.departure, newData.return, id)) return {success: false, message: "❌ New time/vehicle conflicts."};
    sheet.getRange(rowIndex + 1, map['vehicle'] + 1).setValue(String(newData.vehicle).toUpperCase());
    sheet.getRange(rowIndex + 1, map['departure'] + 1).setValue(newData.departure);
    sheet.getRange(rowIndex + 1, map['return'] + 1).setValue(newData.return);
    return {success: true, message: "✅ Booking Updated."};
  } catch(e) { return {success: false, message: e.message}; } finally { lock.releaseLock(); }
}

function cancelBooking(id) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return {success: false, message: "System busy."}; }
  try {
    var sheet = openSheet(BOOKING_SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    var map = getColMap(sheet);
    var rowIndex = -1;
    
    // Find the booking
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][map['ID']]) === String(id)) {
        rowIndex = i;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return {success: false, message: "ID not found."};
    }
    
    var currentStatus = String(data[rowIndex][map['Status']] || "").toUpperCase();
    
    // Only allow cancellation of PENDING or APPROVED bookings
    if (currentStatus !== "PENDING" && currentStatus !== "APPROVED") {
      return {success: false, message: "Cannot cancel booking with status: " + currentStatus};
    }
    
    // Check if booking has already passed
    var returnDate = new Date(data[rowIndex][map['return']]);
    var now = new Date();
    
    if (returnDate < now) {
      return {success: false, message: "Cannot cancel booking that has already passed."};
    }
    
    // Mark as CANCELLED instead of deleting
    sheet.getRange(rowIndex + 1, map['Status'] + 1).setValue("CANCELLED");
    sheet.getRange(rowIndex + 1, map['Processed By'] + 1).setValue("User (Self-Cancelled)");
    sheet.getRange(rowIndex + 1, map['Processed Time'] + 1).setValue(new Date());
    
    return {success: true, message: "✅ Booking cancelled successfully."};
  } catch(e) { 
    return {success: false, message: "Error: " + e.message}; 
  } finally { 
    lock.releaseLock(); 
  }
}

function getBookingsTableData() {
  var sheet = openSheet(BOOKING_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var map = getColMap(sheet);
  var upcoming = [], past = [];
  var now = new Date();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[map['ID']]) continue; 
    var dep = new Date(row[map['departure']]);
    var ret = new Date(row[map['return']]);
    var item = {
      ID: row[map['ID']],
      name: String(row[map['name']] || "").toUpperCase(),
      vehicle: String(row[map['vehicle']] || "").toUpperCase(),
      destination: String(row[map['destination']] || "").toUpperCase(),
      departure: Utilities.formatDate(dep, TIMEZONE, "dd/MM/yyyy HH:mm"),
      return: Utilities.formatDate(ret, TIMEZONE, "dd/MM/yyyy HH:mm"),
      status: String(row[map['Status']] || "PENDING").toUpperCase(),
      _ts: dep.getTime(), _retTs: ret.getTime()    
    };
    if (item._retTs < now.getTime()) past.push(item);
    else upcoming.push(item);
  }
  upcoming.sort((a,b) => a._ts - b._ts);
  past.sort((a,b) => b._retTs - a._retTs);
  return { upcoming: upcoming, past: past };
}

function sendApprovalEmail(id, data) {
  var baseUrl = ScriptApp.getService().getUrl();
  var approveLink = baseUrl + "?action=approve&id=" + encodeURIComponent(id);
  var rejectLink = baseUrl + "?action=reject&id=" + encodeURIComponent(id);
  var fmtDep = Utilities.formatDate(new Date(data.departure), TIMEZONE, "dd/MM/yyyy HH:mm");
  var fmtRet = Utilities.formatDate(new Date(data.return), TIMEZONE, "dd/MM/yyyy HH:mm");
  var html = `
    <div style="font-family:sans-serif; padding:20px; border:1px solid #ddd; border-radius:10px; max-width:600px;">
      <h2 style="color:#4338CA;">Fleet Request: ${id}</h2>
      <p><b>Vehicle:</b> ${data.vehicle}</p>
      <p><b>Requester:</b> ${data.name} (${data.department})</p>
      <p><b>Phone:</b> ${data.phone}</p>
      <p><b>Destination:</b> ${data.destination}</p>
      <p><b>Start:</b> ${fmtDep}</p>
      <p><b>End:</b> ${fmtRet}</p>
      <p><b>Purpose:</b> ${data.purpose}</p>
      <hr style="border:0; border-top:1px solid #eee; margin:20px 0;">
      <div>
        <a href="${approveLink}" style="background:#10B981;color:white;padding:12px 20px;border-radius:5px;text-decoration:none;margin-right:10px;">✅ Approve</a>
        <a href="${rejectLink}" style="background:#EF4444;color:white;padding:12px 20px;border-radius:5px;text-decoration:none;">❌ Reject</a>
      </div>
    </div>`;
  GmailApp.sendEmail(ADMIN_EMAILS, "Action Required: " + id, "", {htmlBody: html, cc: CC_EMAILS});
}

const ADMIN_SHEET_NAME = "Admins";

// ---------- helpers ----------

function _hashPasswordWithSalt(password, salt) {
  // Iterated SHA-256 with a per-user salt (PBKDF2-style)
  var input = salt + password;
  for (var i = 0; i < 10000; i++) {
    var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input);
    input = bytes.map(function(b) { return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0'); }).join('');
  }
  return input;
}

function _getAdminSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(ADMIN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ADMIN_SHEET_NAME);
    // Columns: Email | Salt | PasswordHash | Name | RegisteredAt
    sheet.appendRow(["Email", "Salt", "PasswordHash", "Name", "RegisteredAt"]);
  }
  return sheet;
}

function _generateToken() {
  return Utilities.getUuid();
}

// ---------- admin auth ----------

function adminRegister(email, password, name) {
  if (!email || !password || !name) return {success: false, message: "All fields are required."};
  email = email.trim().toLowerCase();
  var sheet = _getAdminSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === email) {
      return {success: false, message: "An account with this email already exists."};
    }
  }
  var salt = _generateToken(); // unique per user
  var hash = _hashPasswordWithSalt(password, salt);
  sheet.appendRow([email, salt, hash, name.trim(), new Date()]);
  return {success: true, message: "Registration successful. You can now log in."};
}

function adminLoginWithPassword(email, password) {
  if (!email || !password) return {success: false, message: "Email and password are required."};
  email = email.trim().toLowerCase();
  var sheet = _getAdminSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === email) {
      var salt = String(data[i][1]);
      var storedHash = String(data[i][2]);
      if (_hashPasswordWithSalt(password, salt) === storedHash) {
        var token = _generateToken();
        CacheService.getScriptCache().put("admin_token_" + token, email, 7200); // 2 hours
        return {success: true, token: token, name: String(data[i][3])};
      }
      break;
    }
  }
  return {success: false, message: "Incorrect email or password."};
}

function adminLogoutSession(token) {
  if (token) CacheService.getScriptCache().remove("admin_token_" + token);
  return {success: true};
}

function verifyAdminSession(token) {
  if (!token) return null;
  return CacheService.getScriptCache().get("admin_token_" + token);
}

function getCurrentUserEmail() {
  var email = Session.getActiveUser().getEmail();
  return { email: email };
}

function approveRequestInApp(id, action, sessionToken) {
  var adminEmail = verifyAdminSession(sessionToken);
  if (!adminEmail) {
    return {success: false, message: "Session expired or unauthorized. Please log in again."};
  }
  var result = processApproval(id, action, adminEmail);
  var succeeded = result && (result.indexOf("APPROVED") !== -1 || result.indexOf("REJECTED") !== -1);
  return {success: succeeded, message: result};
}

function processApproval(id, action, processedBy) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return "System busy.";
  try {
    var sheet = openSheet(BOOKING_SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    var map = getColMap(sheet);
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][map['ID']]) === String(id)) {
        var status = data[i][map['Status']];
        if (status !== "PENDING") return "Request " + id + " is already " + status;
        
        var newStatus = (action === "approve") ? "APPROVED" : "REJECTED";
        var byLabel = processedBy ? "Admin (" + processedBy + ")" : "Admin (Email Link)";
        sheet.getRange(i+1, map['Status']+1).setValue(newStatus);
        sheet.getRange(i+1, map['Processed By']+1).setValue(byLabel);
        sheet.getRange(i+1, map['Processed Time']+1).setValue(new Date());

        // --- REQUESTER NOTIFICATION ---
        var email = data[i][map['email']];
        var name = data[i][map['name']];
        if (email) {
          var subject = newStatus === "APPROVED" ? "✅ Booking Approved" : "❌ Booking Rejected";
          var msg = "Dear " + name + ",\n\nYour booking request " + id + " has been " + newStatus + ".\n\nThank you.";
          GmailApp.sendEmail(email, subject + ": " + id, msg);
        }

        return "Request " + id + " has been " + newStatus;
      }
    }
    return "Error: Request ID " + id + " not found.";
  } catch (e) { return "Error processing request: " + e.message; } finally { lock.releaseLock(); }
}
