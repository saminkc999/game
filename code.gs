// =============================================
// SHOPIE GAMING — Google Apps Script
// Paste this in Extensions > Apps Script
// =============================================

var NOTIFY_EMAIL = "anniegaming934@gmail.com"; // Your Gmail

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // ── 1. Save photo to Google Drive ──────────
    var photoUrl = "";
    if (data.photoB64) {
      var folder = getOrCreateFolder("Shopie Gaming Photos");
      var bytes  = Utilities.base64Decode(data.photoB64);
      var blob   = Utilities.newBlob(bytes, data.photoMime || "image/jpeg", data.photoName || "photo.jpg");
      var file   = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoUrl = file.getUrl();
    }

    // ── 2. Save row to Google Sheet ────────────
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Add header row if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp", "Name", "Gmail", "Game Name", "Message", "Photo Link"
      ]);
      // Bold the header
      sheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#1a0533").setFontColor("#ffd700");
    }

    sheet.appendRow([
      new Date(),
      data.name,
      data.email,
      data.gameName,
      data.message,
      photoUrl ? '=HYPERLINK("' + photoUrl + '","View Photo")' : "No photo"
    ]);

    // ── 3. Send Gmail notification ─────────────
    var subject = "📸 New photo from " + data.name + " | " + data.gameName + " — Shopie Gaming";
    var body =
      "New submission received!\n\n" +
      "Name:    " + data.name + "\n" +
      "Gmail:   " + data.email + "\n" +
      "Game:    " + data.gameName + "\n" +
      "Message: " + data.message + "\n\n" +
      (photoUrl ? "Photo: " + photoUrl : "No photo attached") + "\n\n" +
      "Reply to: " + data.email;

    GmailApp.sendEmail(NOTIFY_EMAIL, subject, body, {
      replyTo: data.email,
      name:    "Shopie Gaming Form"
    });

    return ok();

  } catch (err) {
    return error(err.message);
  }
}

// ── Helpers ──────────────────────────────────
function getOrCreateFolder(name) {
  var folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function ok() {
  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function error(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}