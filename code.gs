// ==========================================
// CONFIG
// ==========================================
const SHEET_ID = "1jGklnO-SO6lHnocPmXPghnMRovj0qswnuahWfHdSK0E";
const FOLDER_ID = "1YuUgnHV5Nr3ziAM3ocWvVyCa43Q96w_E";

// Razorpay keys are now read from Script Properties instead of being hardcoded.
// Set them once via Project Settings > Script Properties, or run setRazorpayKeys()
// below from the script editor one time with your real values, then delete the call.
function setRazorpayKeys_ONE_TIME_SETUP(keyId, keySecret) {
  PropertiesService.getScriptProperties().setProperty("RZP_KEY_ID", keyId);
  PropertiesService.getScriptProperties().setProperty("RZP_KEY_SECRET", keySecret);
}
function getRzpKeyId_() { return PropertiesService.getScriptProperties().getProperty("RZP_KEY_ID"); }
function getRzpKeySecret_() { return PropertiesService.getScriptProperties().getProperty("RZP_KEY_SECRET"); }

// ==========================================
// WEB APP & FILE HANDLING
// ==========================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('West Bengal Agri Stack — ID Card Generator')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getTemplateImage(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) { return { error: e.toString() }; }
}

function saveCardDataToDriveAndSheets(formData, base64CardImage, base64PassportPhoto) {
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const cardBlob = Utilities.newBlob(Utilities.base64Decode(base64CardImage.split(',')[1]), 'image/jpeg', formData.farmerId + '_ID_Card.jpg');
    const cardFile = folder.createFile(cardBlob);

    let passportFileUrl = "";
    if (base64PassportPhoto) {
      const photoBlob = Utilities.newBlob(Utilities.base64Decode(base64PassportPhoto.split(',')[1]), 'image/jpeg', formData.farmerId + '_Passport_Photo.jpg');
      passportFileUrl = folder.createFile(photoBlob).getUrl();
    }

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Generate Card");
    if (!sheet) throw new Error("Sheet not found!");

    sheet.appendRow([
      formData.timestamp, formData.farmerId, formData.enrollment, formData.nameEn, formData.nameBn,
      formData.fatherName, formData.dob, formData.gender, formData.mobile, formData.aadhaar,
      formData.district, formData.subDistrict, formData.village, formData.pincode,
      cardFile.getUrl(), passportFileUrl, formData.generatedBy || "Customer"
    ]);
    return { success: true };
  } catch (error) { return { success: false, error: error.toString() }; }
}

// ==========================================
// ADMIN AUTH (server-side, session-token based)
// ==========================================
// Sessions are stored in CacheService, expire automatically, and are the ONLY
// thing that unlocks admin-only functions below. A client can no longer just
// flip a JS variable or read localStorage to gain admin access.

const SESSION_TTL_SECONDS = 6 * 60 * 60; // 6 hours

function getAdminPasswordHash_() {
  let props = PropertiesService.getScriptProperties();
  let hash = props.getProperty("ADMIN_PWD_HASH");
  if (!hash) {
    // First run default password is "admin123" — CHANGE THIS immediately from the Settings tab.
    hash = hashPassword_("admin123");
    props.setProperty("ADMIN_PWD_HASH", hash);
  }
  return hash;
}

function hashPassword_(pwd) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pwd, Utilities.Charset.UTF_8);
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function adminLogin(password) {
  if (!password) return { success: false, error: "Password required" };
  if (hashPassword_(password) === getAdminPasswordHash_()) {
    const token = Utilities.getUuid();
    CacheService.getScriptCache().put("session_" + token, "valid", SESSION_TTL_SECONDS);
    return { success: true, token: token };
  }
  return { success: false, error: "Incorrect password" };
}

function isValidAdminSession_(token) {
  if (!token) return false;
  return CacheService.getScriptCache().get("session_" + token) === "valid";
}

function requireAdmin_(token) {
  if (!isValidAdminSession_(token)) {
    throw new Error("UNAUTHORIZED: Admin session invalid or expired. Please log in again.");
  }
}

function adminChangePassword(token, oldPwd, newPwd) {
  requireAdmin_(token);
  if (hashPassword_(oldPwd) !== getAdminPasswordHash_()) {
    return { success: false, error: "Incorrect current password" };
  }
  if (!newPwd || newPwd.length < 4) {
    return { success: false, error: "New password must be at least 4 characters" };
  }
  PropertiesService.getScriptProperties().setProperty("ADMIN_PWD_HASH", hashPassword_(newPwd));
  return { success: true };
}

function adminLogout(token) {
  if (token) CacheService.getScriptCache().remove("session_" + token);
  return { success: true };
}

// ==========================================
// ADMIN DASHBOARD & DROPDOWN LOGIC
// ==========================================
function getActualCardCount(token) {
  requireAdmin_(token);
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Generate Card");
  if (!sheet) return { total: 0, today: 0, admin: 0, customer: 0 };

  const data = sheet.getDataRange().getValues();
  let counts = { total: 0, today: 0, admin: 0, customer: 0 };
  let todayStr = new Date().toLocaleDateString();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1]) {
      counts.total++;
      let dateCell = new Date(data[i][0]);
      if (!isNaN(dateCell.getTime()) && dateCell.toLocaleDateString() === todayStr) counts.today++;
      (data[i][16] === "Admin") ? counts.admin++ : counts.customer++;
    }
  }
  return counts;
}

// Used for customer-facing autofill — intentionally NOT admin-gated, but does not
// expose Aadhaar/mobile any more than the form itself already would for that farmer ID.
function getFarmerDetails(farmerId) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Generate Card");
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]).trim().toUpperCase() === String(farmerId).trim().toUpperCase()) {
      let passportFileId = "";
      let targetCell = data[i][15] || data[i][14];
      if (targetCell) { let match = targetCell.match(/[-\w]{25,}/); if (match) passportFileId = match[0]; }

      return {
        farmerId: data[i][1], enrollment: data[i][2], nameEn: data[i][3], nameBn: data[i][4],
        fatherName: data[i][5], dob: data[i][6], gender: data[i][7], mobile: data[i][8],
        aadhaar: data[i][9], district: data[i][10], subDistrict: data[i][11], village: data[i][12],
        pincode: data[i][13], fileId: passportFileId
      };
    }
  }
  return null;
}

function getDropdownData() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("DROPDOWN");
  if (!sheet) return {};
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  const map = {};
  data.forEach(row => {
    let dist = (row[0] || "").toString().trim(), sub = (row[1] || "").toString().trim(), vill = (row[2] || "").toString().trim();
    if (dist) { if (!map[dist]) map[dist] = {}; if (sub) { if (!map[dist][sub]) map[dist][sub] = []; if (vill && !map[dist][sub].includes(vill)) map[dist][sub].push(vill); } }
  });
  return map;
}

function saveGlobalSettings(token, settings) {
  requireAdmin_(token);
  try {
    PropertiesService.getScriptProperties().setProperty("GLOBAL_LAYOUT_SETTINGS", JSON.stringify(settings));
    return true;
  } catch (e) { return false; }
}

// Layout settings are needed to render the preview for everyone, so this stays public/read-only.
function getGlobalSettings() {
  try {
    let settingsStr = PropertiesService.getScriptProperties().getProperty("GLOBAL_LAYOUT_SETTINGS");
    if (settingsStr) return JSON.parse(settingsStr);
  } catch (e) { return null; }
  return null;
}

// ==========================================
// PAYMENT, PRICING & TRANSACTIONS LOGIC
// ==========================================
function getFeeAmount() {
  let props = PropertiesService.getScriptProperties();
  let fee = props.getProperty("ID_CARD_FEE");
  return fee ? parseInt(fee) : 25;
}

function updateFeeAmount(token, newAmount) {
  requireAdmin_(token);
  let amt = parseInt(newAmount);
  if (isNaN(amt) || amt <= 0) return { success: false, error: "Invalid amount" };
  PropertiesService.getScriptProperties().setProperty("ID_CARD_FEE", amt.toString());
  return { success: true };
}

function saveTransactionData_(farmerId, txnId, amount, status) {
  let ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName("Transactions");
  if (!sheet) {
    sheet = ss.insertSheet("Transactions");
    sheet.appendRow(["Timestamp", "Farmer ID", "Transaction ID", "Amount (₹)", "Status"]);
    sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#d4edda");
  }
  sheet.appendRow([new Date().toLocaleString(), farmerId, txnId, amount, status]);
}

function getAdminPaymentData(token) {
  requireAdmin_(token);
  let ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName("Transactions");
  if (!sheet) return { revenue: 0, history: [] };

  let data = sheet.getDataRange().getValues();
  let totalRev = 0; let history = [];

  for (let i = 1; i < data.length; i++) {
    let amt = parseFloat(data[i][3]) || 0;
    let status = data[i][4];
    if (status === "SUCCESS") totalRev += amt;
    history.push({ date: data[i][0], farmerId: data[i][1], txnId: data[i][2], amount: amt, status: status });
  }
  return { revenue: totalRev, history: history.reverse().slice(0, 50) };
}

// ==========================================
// RAZORPAY GATEWAY INTEGRATION
// ==========================================
function createRazorpayOrder(farmerId) {
  const currentFee = getFeeAmount();
  const amountInPaise = currentFee * 100;

  const payload = {
    "amount": amountInPaise,
    "currency": "INR",
    "receipt": "rcpt_" + farmerId + "_" + new Date().getTime()
  };

  const options = {
    "method": "post",
    "headers": {
      "Authorization": "Basic " + Utilities.base64Encode(getRzpKeyId_() + ":" + getRzpKeySecret_()),
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch("https://api.razorpay.com/v1/orders", options);
    const result = JSON.parse(response.getContentText());
    if (result.id) {
      return { success: true, order_id: result.id, amount: result.amount, key_id: getRzpKeyId_() };
    } else {
      return { success: false, error: result.error ? result.error.description : "Order creation failed" };
    }
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function verifyRazorpaySignature(paymentId, orderId, signature, farmerId) {
  const generatedSignature = Utilities.computeHmacSha256Signature(orderId + "|" + paymentId, getRzpKeySecret_());
  const generatedSignatureHex = generatedSignature.map(function (byte) {
    return ('0' + (byte & 0xFF).toString(16)).slice(-2);
  }).join('');

  if (generatedSignatureHex === signature) {
    let paidAmt = getFeeAmount();
    saveTransactionData_(farmerId, paymentId, paidAmt, "SUCCESS");
    return { success: true };
  } else {
    saveTransactionData_(farmerId, paymentId, 0, "FAILED");
    return { success: false };
  }
}
