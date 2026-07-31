var WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbxx0g8iFoeEusDpMNaaGQKCKXlSmGD_XwPY2xyCmVIcoTn-MFw/exec";
var WEBHOOK_SECRET = "2fef9ece-5a0b-43f5-b0e3-d020940e0c7b";

function onChangeLamMoiCache(e) {
  try {
    UrlFetchApp.fetch(WEBHOOK_URL + "?action=lam_moi_cache&secret=" + WEBHOOK_SECRET, { muteHttpExceptions: true });
  } catch (err) {
    // bo qua loi mang - khong chan thao tac cua nguoi dung tren file nay
  }
}
