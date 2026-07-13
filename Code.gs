// ==== SETUP ====
// 1. Buka https://sheets.google.com, buat Sheet baru (misal "Absen KKN")
// 2. Buka Extensions > Apps Script, hapus isi default, tempel semua kode ini
// 3. Ganti SECRET di bawah — HARUS SAMA PERSIS dengan SECRET di display.html
// 4. Deploy > New deployment > Web app
//      - Execute as: Me
//      - Who has access: Anyone
// 5. Salin URL yang muncul (diakhiri /exec), tempel ke SCRIPT_URL
//    di display.html DAN absen.html

const SECRET = "GANTI_RAHASIA_INI"; // samakan dengan display.html
const SHEET_NAME = "Absen";
const BUCKET_MS = 5 * 60 * 1000;
const TIMEZONE = "GMT+7";

function bucketFor(ts) {
  return Math.floor(ts / BUCKET_MS);
}

function makeCode(bucket, secret) {
  var str = secret + ':' + bucket;
  var hash = 0;
  for (var i = 0; i < str.length; i++) {
    hash = (hash * 31 + str.charCodeAt(i)) >>> 0;
  }
  return hash.toString(36).toUpperCase().padStart(7, '0').slice(-6);
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Tanggal", "Nama", "Jam", "Kode"]);
  }
  return sheet;
}

function doPost(e) {
  var result;
  try {
    var body = JSON.parse(e.postData.contents);
    var name = (body.name || "").trim();
    var code = (body.code || "").trim().toUpperCase();

    var now = new Date();
    var nowBucket = bucketFor(now.getTime());
    var validCodes = [
      makeCode(nowBucket, SECRET),
      makeCode(nowBucket - 1, SECRET) // toleransi kode yang baru saja ganti
    ];

    if (!name || validCodes.indexOf(code) === -1) {
      result = { status: "invalid" };
    } else {
      var dateStr = Utilities.formatDate(now, TIMEZONE, "yyyy-MM-dd");
      var timeStr = Utilities.formatDate(now, TIMEZONE, "HH:mm");
      var sheet = getSheet();
      var data = sheet.getDataRange().getValues();

      var already = null;
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === dateStr && data[i][1] === name) {
          already = data[i][2];
          break;
        }
      }

      if (already) {
        result = { status: "already", time: already };
      } else {
        sheet.appendRow([dateStr, name, timeStr, code]);
        result = { status: "ok", time: timeStr };
      }
    }
  } catch (err) {
    result = { status: "invalid", error: String(err) };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
