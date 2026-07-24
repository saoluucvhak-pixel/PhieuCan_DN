/** * HỆ THỐNG QUẢN LÝ HAKGROUP - PHÂN TÁCH BIỆT LẬP IMPORT & CÁC CHỨC NĂNG BÁO CÁO (2026) - ĐÃ FIX LỖI TRÙNG BATCH */
const CONFIG = {
  FOLDER_INPUT: "1sybSo9vSdSq_puu1LQR2m1zT59ZopgrL",
  FOLDER_DONE: "1bAp97Lwrpq6N8z4-2oXSaieszSL2roca",
  SPREADSHEET_ID: "1vqMVxccBA7zlAMHrGsVBydGFwZJ6QuDZW10zJ74V29g",
  URL_BAO_GIA: "https://docs.google.com/spreadsheets/d/1SIhfjP5-6ouRPDj265lAMmI5yWs1XcnedjqpzDwaIC0/edit",
  DATA_SHEET: "PhieuCan_DN",
  SRC_FILE_ID: "1cv11ORWuAF3Sit4f-kA0xrP6-ab4SF-7LEdkCvGi_gI",
  MISA_DST_ID: "1vkeu2YxME6fsp9ed8DokdtV1jxla5pA-H7heHBt-BRs",
  DNTT_FILE_ID: "1oUm87_gbDbnuPc_We0dyZ_e4kHXBHXs95AQAxp5okYo",
  SHEET_BAO_GIA: "Baogia_DN_SAVE"
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('HỆ THỐNG QUẢN LÝ CÂN HAKGROUP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/*********************************************************
 * PHẦN 1: TIẾN TRÌNH IMPORT & ĐỐI SOÁT KIỂM TRA DATA FILE
 *********************************************************/

function step1_PreviewDraft(fileData) {
  try {
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_INPUT);
    if (fileData && fileData.base64) {
      folder.createFile(Utilities.newBlob(Utilities.base64Decode(fileData.base64), fileData.mimeType, fileData.name));
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const dataSheet = ss.getSheetByName(CONFIG.DATA_SHEET);
    const lastRow = dataSheet.getLastRow();

    const duplicateMap = new Map();
    if (lastRow > 1) {
      // FIX: phải đọc tối thiểu 25 cột (không phải 23) vì trạng thái "OK" nằm ở cột 25 (chỉ số 24).
      // Đọc thiếu cột khiến row[24] luôn undefined -> status luôn rỗng -> không bao giờ
      // nhận diện đúng dòng "Đã khóa OK" ở bước xem trước.
      const existingData = dataSheet.getRange(2, 1, lastRow - 1, 25).getValues();
      existingData.forEach((row, index) => {
        const keyMaCT = String(row[21] || "").trim();
        if (keyMaCT) duplicateMap.set(keyMaCT, { rowNum: index + 2, status: String(row[24] || "").trim() });
      });
    }

    const files = folder.getFiles();
    let previewRows = [];
    // FIX: theo dõi các key đã xuất hiện TRONG CHÍNH lượt xem trước này, để cảnh báo
    // sớm nếu có phiếu trùng số giữa các file/dòng cùng batch, tránh lỗi khi Xác nhận.
    const seenInThisBatch = new Set();

    while (files.hasNext()) {
      const file = files.next();
      if (!file.getName().match(/\.xls[x]?$/i)) continue;

      const tempFile = Drive.Files.insert({ title: "TMP_" + file.getName() }, file.getBlob(), { convert: true });
      const values = SpreadsheetApp.openById(tempFile.id).getSheets()[0].getDataRange().getValues();

      let hIdx = values.findIndex(r => r.some(c => String(c).toLowerCase().includes("số phiếu")));
      if (hIdx === -1) { Drive.Files.remove(tempFile.id); continue; }

      const rowsToProcess = values.slice(hIdx + 1);

      for (let r of rowsToProcess) {
        const spRaw = String(r[0] || "").trim();
        if (!spRaw || spRaw.toLowerCase().includes("ngày") || spRaw.toLowerCase().includes("tổng") || spRaw.length > 20) continue;

        const valA_Dich = String(r[1] || "").trim();
        const valF_Dich = String(r[4] || "").trim();

        let hGocValue = parseFloat(String(r[7] || "").replace(/[^0-9.]/g, ""));
        if (!r[7] || isNaN(hGocValue) || hGocValue === 0) continue;

        let dateC = toDateObj(r[2]);
        let dateD = toDateObj(r[3]);

        let now = new Date();
        let nam = (dateC) ? dateC.getFullYear() : now.getFullYear();
        let currentMaChungTu = valA_Dich + "/" + nam + "/NK";

        // Logic khối lượng độc lập từng cột (< 70 nhân 1000)
        let rawCan1 = parseFloat(String(r[5] || "").replace(/[^0-9.]/g, "")) || 0;
        let rawCan2 = parseFloat(String(r[6] || "").replace(/[^0-9.]/g, "")) || 0;
        let rawHang = parseFloat(String(r[7] || "").replace(/[^0-9.]/g, "")) || 0;

        let previewCan1 = rawCan1 < 70 ? rawCan1 * 1000 : rawCan1;
        let previewCan2 = rawCan2 < 70 ? rawCan2 * 1000 : rawCan2;
        let previewHang = rawHang < 70 ? rawHang * 1000 : rawHang;

        let isError = false; let errorMsg = "";
        if (!dateC || !(dateC instanceof Date) || isNaN(dateC.getTime())) { isError = true; errorMsg += "Lỗi Ngày Cân 1. "; }
        if (!dateD || !(dateD instanceof Date) || isNaN(dateD.getTime())) { isError = true; errorMsg += "Lỗi Ngày Cân 2. "; }

        // FIX: đánh dấu lỗi nếu số chứng từ đã xuất hiện trong chính lượt xem trước này
        if (seenInThisBatch.has(currentMaChungTu)) {
          isError = true;
          errorMsg += "Trùng Số Chứng Từ ngay trong dữ liệu đang nạp (" + currentMaChungTu + "). ";
        } else {
          seenInThisBatch.add(currentMaChungTu);
        }

        let typeImport = "Mới";
        if (duplicateMap.has(currentMaChungTu)) {
          const info = duplicateMap.get(currentMaChungTu);
          typeImport = (info.status === "OK") ? "Bỏ qua (Đã khóa OK)" : "Cập nhật dòng cũ";
        }

        previewRows.push({
          isError: isError, errorMsg: errorMsg.trim(), typeImport: typeImport, uniqueKey: currentMaChungTu,
          soPhieu: valA_Dich,
          ngayCan1: dateC ? Utilities.formatDate(dateC, "GMT+7", "MM/dd/yyyy") : "Lỗi định dạng ngày",
          gioCan1: dateC ? Utilities.formatDate(dateC, "GMT+7", "HH:mm:ss") : "",
          ngayCan2: dateD ? Utilities.formatDate(dateD, "GMT+7", "MM/dd/yyyy") : "Lỗi định dạng ngày",
          gioCan2: dateD ? Utilities.formatDate(dateD, "GMT+7", "HH:mm:ss") : "",
          soXe: valF_Dich, klCan1: previewCan1, klCan2: previewCan2, klHangGoc: previewHang, rawRowData: r
        });
      }
      Drive.Files.remove(tempFile.id);
    }

    // Khởi tạo file Sheet Draft vật lý đối soát thực tế
    let draftUrl = "";
    if (previewRows.length > 0) {
      const draftSS = SpreadsheetApp.create("DRAFT_XemTruoc_PhieuCan_" + Utilities.formatDate(new Date(), "GMT+7", "ddMM_HHmm"));
      const draftSheet = draftSS.getSheets()[0];
      draftSheet.setName("XemTruoc_Data");
      const draftHeaders = ["Trạng Thái", "Số Chứng Từ", "Số Phiếu", "Số Xe", "Ngày Cân 1", "Giờ Cân 1", "Ngày Cân 2", "Giờ Cân 2", "KL Cân 1", "KL Cân 2", "KL Hàng (Kg)", "Ghi Chú Lỗi"];
      draftSheet.getRange(1, 1, 1, draftHeaders.length).setValues([draftHeaders]).setFontWeight("bold").setBackground("#cfe2ff");
      const draftValues = previewRows.map(row => [
        row.typeImport, row.uniqueKey, row.soPhieu, row.soXe, row.ngayCan1, row.gioCan1, row.ngayCan2, row.gioCan2, row.klCan1, row.klCan2, row.klHangGoc, row.isError ? "❌ " + row.errorMsg : "✔️ Hợp lệ"
      ]);
      draftSheet.getRange(2, 1, draftValues.length, draftHeaders.length).setValues(draftValues);
      draftSheet.getRange(2, 9, draftValues.length, 3).setNumberFormat("#,##0");
      draftSheet.getRange(2, 2, draftValues.length, 3).setNumberFormat("@");
      const draftFileObj = DriveApp.getFileById(draftSS.getId());
      DriveApp.getFolderById(CONFIG.FOLDER_DONE).addFile(draftFileObj);
      DriveApp.getRootFolder().removeFile(draftFileObj);
      draftUrl = draftSS.getUrl();
    }
    return { status: "success", data: previewRows, draftUrl: draftUrl };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

function step1_ConfirmImport(confirmedDataList) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const dataSheet = ss.getSheetByName(CONFIG.DATA_SHEET);
    const lastRow = dataSheet.getLastRow();
    const NUM_COLUMNS = 23; const now = new Date();

    const duplicateMap = new Map();
    if (lastRow > 1) {
      const existingData = dataSheet.getRange(2, 1, lastRow - 1, 25).getValues();
      existingData.forEach((row, index) => {
        const keyMaCT = String(row[21] || "").trim();
        if (keyMaCT) duplicateMap.set(keyMaCT, { rowNum: index + 2, status: String(row[24] || "").trim() });
      });
    }

    let countNew = 0, countUpdate = 0, countSkip = 0; const batchNew = [];

    for (let item of confirmedDataList) {
      if (item.isError) continue;
      const r = item.rawRowData; const valA_Dich = item.soPhieu; const valF_Dich = item.soXe; const uniqueKeyMaCT = item.uniqueKey;
      let valL_Dich = r[9]; let valN_Dich = String(r[13] || "").toUpperCase(); let valO_Dich = String(r[12] || "").toUpperCase();
      let valK_Dich = valN_Dich + "_" + valO_Dich; let valQ_Dich = valN_Dich + "_" + valO_Dich + "_Y";

      if (duplicateMap.has(uniqueKeyMaCT)) {
        const info = duplicateMap.get(uniqueKeyMaCT);

        if (info.status === "OK") { countSkip++; continue; }

        // ---- FIX QUAN TRỌNG (giống lỗi ở bản DH) ----
        // Nếu dòng trùng này là dòng MỚI vừa thêm trong CHÍNH batch hiện tại (chưa
        // có mặt thật trên sheet -> không có rowNum), phải cập nhật thẳng vào mảng
        // batchNew đang ở RAM, KHÔNG được gọi getRange lên sheet bằng info.rowNum
        // (lúc đó info.rowNum sẽ là undefined -> lỗi "Tham số (null,number)...").
        if (info.rowNum === undefined && info.batchIndex !== undefined) {
          const row = batchNew[info.batchIndex];
          row[10] = valK_Dich;
          row[11] = valL_Dich;
          row[13] = valN_Dich;
          row[14] = valO_Dich;
          row[16] = valQ_Dich;
          row[18] = now;
          countUpdate++;
          continue;
        }

        dataSheet.getRange(info.rowNum, 11).setValue(valK_Dich);
        dataSheet.getRange(info.rowNum, 12).setValue(valL_Dich);
        dataSheet.getRange(info.rowNum, 14, 1, 2).setValues([[valN_Dich, valO_Dich]]);
        dataSheet.getRange(info.rowNum, 17).setValue(valQ_Dich);
        dataSheet.getRange(info.rowNum, 19).setValue(now);
        countUpdate++; continue;
      }

      let newRow = new Array(NUM_COLUMNS).fill("");
      let dateC = toDateObj(r[2]); let dateD = toDateObj(r[3]);
      newRow[0] = valA_Dich;
      if (dateC) { newRow[1] = Utilities.formatDate(dateC, "GMT+7", "MM/dd/yyyy"); newRow[2] = Utilities.formatDate(dateC, "GMT+7", "HH:mm:ss"); }
      if (dateD) { newRow[3] = Utilities.formatDate(dateD, "GMT+7", "MM/dd/yyyy"); newRow[4] = Utilities.formatDate(dateD, "GMT+7", "HH:mm:ss"); }
      newRow[5] = valF_Dich; newRow[7] = item.klCan1; newRow[8] = item.klCan2; newRow[9] = item.klHangGoc;
      newRow[10] = valK_Dich; newRow[11] = valL_Dich; newRow[12] = "GK"; newRow[13] = valN_Dich; newRow[14] = valO_Dich;
      newRow[15] = "Y"; newRow[16] = valQ_Dich; newRow[17] = 0; newRow[18] = now; newRow[21] = uniqueKeyMaCT; newRow[22] = uniqueKeyMaCT;
      batchNew.push(newRow);
      // FIX: lưu batchIndex thay vì object rỗng, để nếu gặp trùng lặp ngay trong
      // cùng batch (dòng chưa ghi lên sheet) thì cập nhật đúng dòng RAM, không undefined.
      duplicateMap.set(uniqueKeyMaCT, { status: "", batchIndex: batchNew.length - 1 });
      countNew++;
    }

    if (batchNew.length > 0) dataSheet.getRange(dataSheet.getLastRow() + 1, 1, batchNew.length, NUM_COLUMNS).setValues(batchNew);

    const folder = DriveApp.getFolderById(CONFIG.FOLDER_INPUT); const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next(); if (!file.getName().match(/\.xls[x]?$/i)) continue;
      const meta = Drive.Files.get(file.getId(), { fields: "parents" });
      Drive.Files.update({}, file.getId(), null, { addParents: CONFIG.FOLDER_DONE, removeParents: meta.parents[0].id });
    }

    let priceMsg = "";
    if (countNew > 0 || countUpdate > 0) { const priceResult = runCalculatePrice(); priceMsg = " | " + priceResult.message; }
    return { status: "success", message: `Mới: ${countNew}, Cập nhật: ${countUpdate}, Bỏ qua: ${countSkip}${priceMsg}` };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

/*********************************************************
 * PHẦN 2: BIỆT LẬP CÁC CHỨC NĂNG XUẤT BÁO CÁO (CHẠY ĐỘC LẬP)
 *********************************************************/

function runCreateMisaData(fromDate, toDate) {
  try {
    const res = copyDataWithFinalLookup(fromDate, toDate);
    if(res.status === "success") {
      return { status: "success", message: "Đã xử lý trích xuất dữ liệu hạch toán MISA thành công!" };
    } else {
      return res;
    }
  } catch (e) { return { status: "error", message: e.toString() }; }
}

function downloadMisaExcel() {
  try {
    const ssMisa = SpreadsheetApp.openById(CONFIG.MISA_DST_ID);
    const sheet = ssMisa.getSheetByName('Update_MiSa_PC');
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: "error", message: "Bảng dữ liệu MISA trống! Vui lòng bấm tạo data trước." };

    const data = sheet.getRange(1, 1, lastRow, 31).getValues();
    const tempSS = SpreadsheetApp.create("Misa_Export_" + Utilities.formatDate(new Date(), "GMT+7", "ddMM_HHmm"));
    const tempSheet = tempSS.getSheets()[0];

    tempSheet.getRange(1, 10, data.length, 2).setNumberFormat("@");
    tempSheet.getRange(1, 1, data.length, 31).setValues(data);
    tempSheet.getRange(1, 1, 1, 31).setBackground("#B7B7B7").setFontWeight("bold").setHorizontalAlignment("center");

    const tempFile = DriveApp.getFileById(tempSS.getId());
    DriveApp.getFolderById(CONFIG.FOLDER_DONE).addFile(tempFile);
    DriveApp.getRootFolder().removeFile(tempFile);

    return { status: "success", url: "https://docs.google.com/spreadsheets/d/" + tempSS.getId() + "/export?format=xlsx" };
  } catch (e) { return { status: "error", message: "Lỗi tạo file tải: " + e.toString() }; }
}

function exportToExcel(fromDate, toDate) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.DATA_SHEET);
    const fullData = sheet.getDataRange().getValues();
    if(fullData.length <= 1) return { status: "error", message: "Sheet dữ liệu tổng hợp trống!" };

    const header = fullData[0];
    const skipIndexes = [13, 14, 16, 18, 20, 21, 26, 27];
    const keepIndexes = header.map((_, i) => i).filter(i => !skipIndexes.includes(i));

    const start = fromDate ? new Date(fromDate + "T00:00:00+07:00") : null;
    const end = toDate ? new Date(toDate + "T23:59:59+07:00") : null;

    const filtered = fullData.slice(1).filter(row => {
      let d = parseDate(row[1]); if (!start || !end) return true;
      return d && d >= start && d <= end;
    }).map(row => {
      return keepIndexes.map(i => { if (i === 7) return '="' + String(row[i] || "").trim() + '"'; return row[i]; });
    });

    if (filtered.length === 0) return { status: "error", message: "Không tìm thấy dữ liệu nào trong khoảng thời gian đã chọn!" };

    const newHeader = keepIndexes.map(i => header[i]);
    const tempSS = SpreadsheetApp.create("Bao_Cao_Tong_Hop_" + (fromDate || "All"));
    const tempSheet = tempSS.getSheets()[0];
    tempSheet.getRange(1, 1, 1, newHeader.length).setValues([newHeader]);
    tempSheet.getRange(2, 1, filtered.length, newHeader.length).setValues(filtered);

    const numRows = filtered.length; const getColPos = (oldIdx) => keepIndexes.indexOf(oldIdx) + 1;
    [8, 9, 19, 23, 25].forEach(oldIdx => { let col = getColPos(oldIdx); if (col > 0) tempSheet.getRange(2, col, numRows, 1).setNumberFormat("#,##0"); });
    [2, 4].forEach(oldIdx => { let col = getColPos(oldIdx); if (col > 0) tempSheet.getRange(2, col, numRows, 1).setNumberFormat("HH:mm:ss"); });
    let colB = getColPos(1); if (colB > 0) tempSheet.getRange(2, colB, numRows, 1).setNumberFormat("dd/MM/yyyy");
    let colH = getColPos(7); if (colH > 0) tempSheet.getRange(2, colH, numRows, 1).setNumberFormat("@");

    tempSheet.getRange(1, 1, 1, newHeader.length).setFontWeight("bold").setBackground("#B7B7B7").setHorizontalAlignment("center");
    const tempFile = DriveApp.getFileById(tempSS.getId());
    DriveApp.getFolderById(CONFIG.FOLDER_DONE).addFile(tempFile);
    DriveApp.getRootFolder().removeFile(tempFile);

    return { status: "success", url: "https://docs.google.com/spreadsheets/d/" + tempSS.getId() + "/export?format=xlsx" };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

/*********************************************************
 * PHẦN KỸ THUẬT NGẦM CHỐNG TREO & SỬA LỖI ĐỊNH DẠNG NGÀY
 *********************************************************/
function copyDataWithFinalLookup(fromDate, toDate) {
  try {
    const ssSourcePC = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.DATA_SHEET);
    const ssTarget   = SpreadsheetApp.openById(CONFIG.MISA_DST_ID).getSheetByName('Update_MiSa_PC');
    const ssRef      = SpreadsheetApp.openById(CONFIG.SRC_FILE_ID);
    const ssDNTT     = SpreadsheetApp.openById(CONFIG.DNTT_FILE_ID);
    let start = fromDate ? new Date(fromDate + "T00:00:00+07:00") : new Date(new Date().getFullYear(), new Date().getMonth(), 1);
    let end = toDate ? new Date(toDate + "T23:59:59+07:00") : new Date();
    const mapNG = createSimpleMap_(ssRef.getSheetByName('DM_NG'), 1, 3);
    const dataKH = ssRef.getSheetByName('DM_KH').getDataRange().getValues();
    let mapKH = {}; for (let i = 1; i < dataKH.length; i++) { let k = String(dataKH[i][1]).trim(); if (k) mapKH[k] = { colC: dataKH[i][2], colD: dataKH[i][3], colE: dataKH[i][4] }; }
    const dataHDNCC = ssRef.getSheetByName('HD_NCC').getDataRange().getValues();
    let mapHDNCC = {}; for (let i = 1; i < dataHDNCC.length; i++) { let k = String(dataHDNCC[i][2]).trim(); if (k) mapHDNCC[k] = { colG: dataHDNCC[i][6], colE: dataHDNCC[i][4] }; }
    const dataDNTT = ssDNTT.getSheetByName('DNTT_GK_DN_CT').getDataRange().getValues();
    let mapDNTT = {}; for (let i = 1; i < dataDNTT.length; i++) { let k = String(dataDNTT[i][11] || "").trim(); if (k) mapDNTT[k] = { colT: dataDNTT[i][19], colV: dataDNTT[i][21] }; }
    const sourceData = ssSourcePC.getDataRange().getValues(); let targetData = [];
    for (let i = 1; i < sourceData.length; i++) {
      let row = sourceData[i]; let ngayPhieu = parseDate(row[1]); if (ngayPhieu && (ngayPhieu < start || ngayPhieu > end)) continue;
      let newRow = new Array(33).fill(""); let soHopDongGoc = String(row[22] || "").trim(); let khoiLuong = (parseFloat(row[9]) || 0) / 1000; let donGia = parseFloat(row[23]) || 0; let thanhTienTuCotZ = parseFloat(row[25]) || 0;
      newRow[0] = row[1]; newRow[1] = row[1]; newRow[2] = soHopDongGoc; newRow[3] = row[5]; newRow[4] = khoiLuong; newRow[5] = donGia; newRow[6] = thanhTienTuCotZ;
      let soHopDongMoi = ""; if (soHopDongGoc && mapDNTT[soHopDongGoc]) { soHopDongMoi = String(mapDNTT[soHopDongGoc].colT).trim(); newRow[9] = '="' + soHopDongMoi + '"'; newRow[8] = mapDNTT[soHopDongGoc].colV || "31/12/2050"; } else { newRow[8] = "31/12/2050"; }
      newRow[16] = "2"; newRow[21] = "VND"; newRow[23] = "621A.001"; newRow[24] = "Gỗ tròn keo lai CW"; newRow[26] = "621"; newRow[27] = "33111"; newRow[28] = "Tấn"; newRow[29] = "62111"; newRow[30] = "155A.001";
      newRow[31] = row[13]; newRow[32] = row[14]; newRow[12] = mapNG[String(row[14]).trim()] || ""; newRow[7] = mapKH[String(row[13]).trim()] ? mapKH[String(row[13]).trim()].colE : "";
      // FIX: bỏ gán lồng String(maNCC = String(...)) thừa, giữ đúng logic gán 1 lần
      let maNCC = ""; let tenNCC = ""; if (soHopDongMoi && mapHDNCC[soHopDongMoi]) { maNCC = String(mapHDNCC[soHopDongMoi].colG).trim(); tenNCC = mapHDNCC[soHopDongMoi].colE; }
      if (!maNCC) { let keyKH = String(row[13]).trim(); if (mapKH[keyKH]) { maNCC = String(mapKH[keyKH].colC).trim(); tenNCC = mapKH[keyKH].colD; } }
      newRow[10] = maNCC ? '="' + maNCC + '"' : ""; newRow[11] = tenNCC; targetData.push(newRow);
    }
    if (ssTarget.getLastRow() > 1) ssTarget.getRange(2, 1, ssTarget.getLastRow(), 33).clearContent();
    if (targetData.length > 0) { ssTarget.getRange(2, 10, targetData.length, 2).setNumberFormat("@"); ssTarget.getRange(2, 1, targetData.length, 33).setValues(targetData); ssTarget.getRange(2, 1, targetData.length, 2).setNumberFormat("dd/MM/yyyy"); ssTarget.getRange(2, 9, targetData.length, 1).setNumberFormat("dd/MM/yyyy"); }
    return { status: "success" };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

function runCalculatePrice() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); const sheet = ss.getSheetByName(CONFIG.DATA_SHEET); const data = sheet.getDataRange().getValues();
    const ssBG = SpreadsheetApp.openByUrl(CONFIG.URL_BAO_GIA); const sheetBG = ssBG.getSheetByName(CONFIG.SHEET_BAO_GIA); const rawDataBG = sheetBG.getDataRange().getValues();
    const dataBG = rawDataBG.slice(1).map(bg => { return { start: new Date(bg[1]).getTime(), end: new Date(bg[2]).getTime(), keyQ: String(bg[3] || "").trim().toUpperCase(), minKl: parseFloat(bg[4]) || 0, maxKl: parseFloat(bg[5]) || 0, price: parseFloat(bg[6]) || 0 }; });
    const resT = []; const resX = []; const resY = []; const resZ = [];
    for (let i = 1; i < data.length; i++) {
      let r = data[i]; let currentStatusY = String(r[24] || "").trim(); if (currentStatusY === "OK") { resT.push([r[19]]); resX.push([r[23]]); resY.push([r[24]]); resZ.push([r[25]]); continue; }
      let valQ = String(r[16] || "").trim().toUpperCase(); let rVal = parseFloat(r[17]) || 0; let klJ = parseFloat(r[9]) || 0; let klSoSanh = klJ / 1000;
      let dt = new Date(r[1]); if (r[2]) { let t = r[2]; let h = (t instanceof Date) ? t.getHours() : parseInt(String(t).split(":")[0]) || 0; let m = (t instanceof Date) ? t.getMinutes() : parseInt(String(t).split(":")[1]) || 0; dt.setHours(h, m, 0, 0); }
      let ts = dt.getTime(); let giaFound = 0;
      if (valQ !== "" && klSoSanh > 0) { let listCungMa = dataBG.filter(bg => bg.keyQ === valQ); if (listCungMa.length > 0) { for (let bg of listCungMa) { if (ts >= bg.start && ts < bg.end && klSoSanh > bg.minKl && klSoSanh <= bg.maxKl) { giaFound = bg.price; break; } } } }
      resT.push([giaFound]); let hieuSo = Math.round(giaFound + rVal); resX.push([hieuSo]); let thanhTienRaw = Math.round(klSoSanh * hieuSo); let thanhTienLamTron = Math.floor(thanhTienRaw / 1000) * 1000; resZ.push([thanhTienLamTron]);
      resY.push((giaFound > 0 && thanhTienLamTron > 0) ? ["Test giá"] : ["Lỗi ĐK/Báo giá"]);
    }
    if (resT.length > 0) { sheet.getRange(2, 20, resT.length, 1).setValues(resT); sheet.getRange(2, 24, resX.length, 1).setValues(resX); sheet.getRange(2, 25, resY.length, 1).setValues(resY); sheet.getRange(2, 26, resZ.length, 1).setValues(resZ); sheet.getRange(2, 24, resX.length, 1).setNumberFormat("#,##0"); sheet.getRange(2, 26, resZ.length, 1).setNumberFormat("#,##0"); }
    return { status: "success", message: "Đã tính lại giá giật cấp!" };
  } catch (e) { return { status: "error", message: "Lỗi: " + e.toString() }; }
}

function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  if (typeof v === "number") return new Date(Math.round((v - 25569) * 86400 * 1000));
  if (typeof v === "string") {
    let p = v.includes("-") ? v.split("-") : v.split("/");
    if (p.length === 3) return p[0].length === 4 ? new Date(p[0], p[1]-1, p[2]) : new Date(p[2], p[1]-1, p[0]);
  }
  return null;
}

function toDateObj(val) { if (val instanceof Date) return val; if (typeof val === 'string' && val.trim() !== "") { const parts = val.match(/(\d+)/g); if (parts && parts.length >= 3) return new Date(parts[2], parts[1]-1, parts[0], parts[3]||0, parts[4]||0, parts[5]||0); } return null; }
function createSimpleMap_(s,k,v) { const d=s.getDataRange().getValues(); let m={}; for(let i=1;i<d.length;i++){ let key=String(d[i][k]).trim(); if(key) m[key]=d[i][v]; } return m; }
function onOpen() { SpreadsheetApp.getUi().createMenu('🚀 HỆ THỐNG HAKGROUP').addItem('👉 Xuất Misa thủ công', 'copyDataWithFinalLookup_Manual').addToUi(); }
function copyDataWithFinalLookup_Manual() { const ui = SpreadsheetApp.getUi(); const res = copyDataWithFinalLookup(null, null); ui.alert("Đã chạy đồng bộ xong!"); }

function clearInputFolderFiles() {
  try {
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_INPUT);
    const files = folder.getFiles();
    let count = 0;

    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().match(/\.xls[x]?$/i)) {
        file.setTrashed(true);
        count++;
      }
    }
    return { status: "success", message: `Đã dọn dẹp sạch sẽ, hủy bỏ và xóa thành công ${count} file tạm rác!` };
  } catch (e) {
    return { status: "error", message: "Lỗi dọn dẹp thư mục tạm: " + e.toString() };
  }
}
