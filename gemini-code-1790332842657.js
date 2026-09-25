function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Smart Zone | Online Store')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ⚠️ IMPORTANT: Change these Sheet & Folder IDs to your own Google Drive IDs
const STORE_SHEET_ID = "1jGklnO-SO6lHnocPmXPghnMRovj0qswnuahWfHdSK0E"; 
const PRODUCT_IMAGES_FOLDER_ID = "1YuUgnHV5Nr3ziAM3ocWvVyCa43Q96w_E";

// ==========================================
// PRODUCT MANAGEMENT (ADMIN)
// ==========================================
function getProducts() {
  const ss = SpreadsheetApp.openById(STORE_SHEET_ID);
  let sheet = ss.getSheetByName("Products");
  
  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet("Products");
    sheet.appendRow(["Timestamp", "Product Name", "Price (₹)", "Image URL"]);
    sheet.getRange("A1:D1").setFontWeight("bold").setBackground("#cbd5e1");
    return []; // return empty if just created
  }
  
  const data = sheet.getDataRange().getValues();
  let products = [];
  for (let i = 1; i < data.length; i++) {
    if(data[i][1]) { // If product name exists
      products.push({
        name: data[i][1],
        price: data[i][2],
        image: data[i][3] // URL of image
      });
    }
  }
  return products.reverse(); // Newest first
}

function addStoreProduct(name, price, base64Image) {
  try {
    const folder = DriveApp.getFolderById(PRODUCT_IMAGES_FOLDER_ID);
    // Decode base64 and save as image file
    const imageBlob = Utilities.newBlob(Utilities.base64Decode(base64Image.split(',')[1]), 'image/jpeg', name + '_product.jpg');
    const imageFile = folder.createFile(imageBlob);
    
    // Set file sharing to anyone with link can view (so image loads on website)
    imageFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Modify URL for direct image rendering
    let imageUrl = imageFile.getUrl().replace("file/d/", "uc?export=view&id=").replace("/view?usp=drivesdk", "");

    const ss = SpreadsheetApp.openById(STORE_SHEET_ID);
    let sheet = ss.getSheetByName("Products");
    if (!sheet) {
      sheet = ss.insertSheet("Products");
      sheet.appendRow(["Timestamp", "Product Name", "Price (₹)", "Image URL"]);
    }
    
    sheet.appendRow([new Date().toLocaleString(), name, price, imageUrl]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// ORDER MANAGEMENT & RAZORPAY API
// ==========================================
const RZP_KEY_ID = "rzp_test_TfrnxuSLpFJKoZ";
const RZP_KEY_SECRET = "n84nBKLgcsvx6gD2glF40Avy";

function createEcommerceOrder(totalAmount) {
  const amountInPaise = totalAmount * 100; // INR to Paise

  const payload = {
    "amount": amountInPaise,
    "currency": "INR",
    "receipt": "store_order_" + new Date().getTime()
  };

  const options = {
    "method": "post",
    "headers": {
      "Authorization": "Basic " + Utilities.base64Encode(RZP_KEY_ID + ":" + RZP_KEY_SECRET),
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch("https://api.razorpay.com/v1/orders", options);
    const result = JSON.parse(response.getContentText());
    if (result.id) {
      return { success: true, order_id: result.id, amount: result.amount, key_id: RZP_KEY_ID };
    } else {
      return { success: false, error: result.error ? result.error.description : "Order creation failed" };
    }
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function verifyAndSaveOrder(paymentId, orderId, signature, orderData) {
  // Signature Verification
  const generatedSignature = Utilities.computeHmacSha256Signature(orderId + "|" + paymentId, RZP_KEY_SECRET);
  const generatedSignatureHex = generatedSignature.map(function(byte) { return ('0' + (byte & 0xFF).toString(16)).slice(-2); }).join('');

  if (generatedSignatureHex === signature) {
    // Save to Google Sheets
    const ss = SpreadsheetApp.openById(STORE_SHEET_ID);
    let sheet = ss.getSheetByName("Orders");
    if (!sheet) {
      sheet = ss.insertSheet("Orders"); 
      sheet.appendRow(["Timestamp", "Customer Name", "Phone", "Address", "Items Ordered", "Total Amount (₹)", "Payment ID", "Status"]);
      sheet.getRange("A1:H1").setFontWeight("bold").setBackground("#d4edda");
    }
    
    sheet.appendRow([
      new Date().toLocaleString(), 
      orderData.name, 
      orderData.phone, 
      orderData.address, 
      orderData.items, 
      orderData.amount, 
      paymentId, 
      "SUCCESS"
    ]);
    
    return { success: true };
  } else {
    return { success: false, error: "Payment verification failed due to invalid signature." };
  }
}

function getAdminOrders() {
  const ss = SpreadsheetApp.openById(STORE_SHEET_ID);
  let sheet = ss.getSheetByName("Orders");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  let history = [];
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][1]) {
      history.push({ 
        date: data[i][0], 
        name: data[i][1], 
        phone: data[i][2], 
        address: data[i][3],
        items: data[i][4],
        amount: data[i][5],
        status: data[i][7]
      });
    }
  }
  return history.reverse(); // Show latest orders first
}