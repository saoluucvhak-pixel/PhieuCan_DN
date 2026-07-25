/** * HỆ THỐNG QUẢN LÝ HAKGROUP - PHÂN TÁCH BIỆT LẬP IMPORT & CÁC CHỨC NĂNG BÁO CÁO (2026)
 *  - ĐÃ FIX LỖI TRÙNG BATCH
 *  - ĐÃ FIX #1: Ghi Date object thật thay vì chuỗi format sẵn (tránh lệch ngày/tháng do locale)
 *  - ĐÃ FIX #2: Bổ sung LockService cho các hàm ghi dữ liệu (chống race condition đa người dùng)
 *  - ĐÃ FIX #4: Sửa lỗi ngày bị sai lệch khi Xác nhận Import (do google.script.run tự
 *    chuyển Date thành chuỗi ISO khi đi qua lại giữa client/server)
 *
 *  LƯU Ý: Toàn bộ hằng số cấu hình (CONFIG, BAOGIA_CONFIG, COMPANY_NAME...) đã
 *  được chuyển sang file Config.gs riêng để dễ bảo trì - xem file đó nếu cần
 *  đổi ID Spreadsheet/Folder hoặc tên Sheet.
 */

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
      try {
        const values = SpreadsheetApp.openById(tempFile.id).getSheets()[0].getDataRange().getValues();

        let hIdx = values.findIndex(r => r.some(c => String(c).toLowerCase().includes("số phiếu")));
        if (hIdx === -1) { continue; }

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
            // Lưu ý: các chuỗi ngày/giờ dưới đây CHỈ dùng để HIỂN THỊ trong bảng xem trước
            // (draft sheet) cho người dùng đọc. Dữ liệu ghi thật vào PhieuCan_DN ở bước
            // Xác nhận (step1_ConfirmImport) dùng Date object gốc (dateC/dateD), không
            // dùng các chuỗi này, nên không bị ảnh hưởng bởi lỗi định dạng locale.
            ngayCan1: dateC ? Utilities.formatDate(dateC, "GMT+7", "MM/dd/yyyy") : "Lỗi định dạng ngày",
            gioCan1: dateC ? Utilities.formatDate(dateC, "GMT+7", "HH:mm:ss") : "",
            ngayCan2: dateD ? Utilities.formatDate(dateD, "GMT+7", "MM/dd/yyyy") : "Lỗi định dạng ngày",
            gioCan2: dateD ? Utilities.formatDate(dateD, "GMT+7", "HH:mm:ss") : "",
            soXe: valF_Dich, klCan1: previewCan1, klCan2: previewCan2, klHangGoc: previewHang, rawRowData: r
          });
        }
      } finally {
        // FIX: dùng try/finally để đảm bảo file tạm luôn bị xóa dù có lỗi xảy ra
        // trong lúc đọc (tránh rò rỉ file rác trong Drive khi 1 file lỗi làm hỏng vòng lặp)
        Drive.Files.remove(tempFile.id);
      }
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
  // FIX #2: khóa toàn bộ quá trình xác nhận ghi dữ liệu. Nếu 2 người dùng bấm
  // "Xác nhận" gần như đồng thời, nếu không khóa thì cả hai sẽ đọc cùng
  // getLastRow() và ghi đè lên CÙNG một dải hàng, làm mất dữ liệu của một bên.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
  } catch (e) {
    return { status: "error", message: "Hệ thống đang bận xử lý một yêu cầu khác, vui lòng thử lại sau ít giây." };
  }

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
        // Đây là lớp bảo vệ dự phòng: về lý thuyết step1_PreviewDraft() đã đánh dấu
        // isError=true cho mọi dòng trùng mã chứng từ NGAY TRONG CÙNG batch
        // (xem seenInThisBatch ở trên), nên các dòng như vậy thường đã bị "continue"
        // ở đầu vòng lặp này rồi. Nhánh dưới đây chỉ kích hoạt nếu client gọi thẳng
        // Confirm mà bỏ qua bước Preview, hoặc gửi sai cờ isError — khi đó phải
        // cập nhật thẳng vào mảng batchNew đang ở RAM, KHÔNG được gọi getRange lên
        // sheet bằng info.rowNum (lúc đó info.rowNum sẽ là undefined -> lỗi
        // "Tham số (null,number)...").
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

      // FIX #3: Ghi đúng Ngày THUẦN vào cột "Ngày cân X" và Giờ THUẦN vào cột
      // "Giờ cân X", khớp với đúng quy ước dữ liệu thật đang có trong hệ thống
      // (thay vì ghi 1 datetime đầy đủ vào cả 2 cột như trước — vẫn hiển thị
      // đúng nhưng giá trị gốc lưu trong ô bị lệch chuẩn kiểu dữ liệu).
      newRow[1] = dateC ? toDateOnly_(dateC) : "";
      newRow[2] = dateC ? toTimeOnly_(dateC) : "";
      newRow[3] = dateD ? toDateOnly_(dateD) : "";
      newRow[4] = dateD ? toTimeOnly_(dateD) : "";

      newRow[5] = valF_Dich; newRow[7] = item.klCan1; newRow[8] = item.klCan2; newRow[9] = item.klHangGoc;
      newRow[10] = valK_Dich; newRow[11] = valL_Dich; newRow[12] = "GK"; newRow[13] = valN_Dich; newRow[14] = valO_Dich;
      newRow[15] = "Y"; newRow[16] = valQ_Dich; newRow[17] = 0; newRow[18] = now; newRow[21] = uniqueKeyMaCT; newRow[22] = uniqueKeyMaCT;
      batchNew.push(newRow);
      // FIX: lưu batchIndex thay vì object rỗng, để nếu gặp trùng lặp ngay trong
      // cùng batch (dòng chưa ghi lên sheet) thì cập nhật đúng dòng RAM, không undefined.
      duplicateMap.set(uniqueKeyMaCT, { status: "", batchIndex: batchNew.length - 1 });
      countNew++;
    }

    if (batchNew.length > 0) {
      const startRow = dataSheet.getLastRow() + 1;
      dataSheet.getRange(startRow, 1, batchNew.length, NUM_COLUMNS).setValues(batchNew);

      // Định dạng hiển thị khớp với giá trị THẬT đã ghi (Ngày thuần / Giờ thuần)
      dataSheet.getRange(startRow, 2, batchNew.length, 1).setNumberFormat("dd/MM/yyyy"); // Ngày Cân 1
      dataSheet.getRange(startRow, 3, batchNew.length, 1).setNumberFormat("HH:mm:ss");   // Giờ Cân 1
      dataSheet.getRange(startRow, 4, batchNew.length, 1).setNumberFormat("dd/MM/yyyy"); // Ngày Cân 2
      dataSheet.getRange(startRow, 5, batchNew.length, 1).setNumberFormat("HH:mm:ss");   // Giờ Cân 2
    }

    const folder = DriveApp.getFolderById(CONFIG.FOLDER_INPUT); const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next(); if (!file.getName().match(/\.xls[x]?$/i)) continue;
      const meta = Drive.Files.get(file.getId(), { fields: "parents" });
      if (meta.parents && meta.parents.length > 0) {
        Drive.Files.update({}, file.getId(), null, { addParents: CONFIG.FOLDER_DONE, removeParents: meta.parents[0].id });
      }
    }

    let priceMsg = "";
    if (countNew > 0 || countUpdate > 0) {
      // FIX #2: gọi thẳng bản "_core" KHÔNG khóa, vì ta đang giữ khóa của
      // step1_ConfirmImport rồi. Nếu gọi runCalculatePrice() (bản có khóa)
      // ở đây sẽ bị TREO VĨNH VIỄN (deadlock) vì cùng 1 lượt thực thi lại tự
      // chờ khóa mà chính nó đang giữ.
      const priceResult = runCalculatePrice_core();
      priceMsg = " | " + priceResult.message;
    }
    const finalMsg = `Mới: ${countNew}, Cập nhật: ${countUpdate}, Bỏ qua: ${countSkip}${priceMsg}`;
    logAudit_('IMPORT_PHIEUCAN', 'OK', finalMsg);
    return { status: "success", message: finalMsg };
  } catch (e) {
    logAudit_('IMPORT_PHIEUCAN', 'ERROR', e.toString());
    return { status: "error", message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/*********************************************************
 * PHẦN 1B: NHẬP LIỆU THỦ CÔNG PHIẾU CÂN (MENU 1 - TAB "NHẬP TAY")
 *********************************************************/

// fields: {soPhieu, soXe, soXe2, ngayCan1, gioCan1, ngayCan2, gioCan2, klCan1, klCan2, klHang, khachHang, maKH, maNG}
// LƯU Ý QUAN TRỌNG: "maKH" ở đây thực chất là Mã ĐẠI LÝ (ghi vào cột N "ĐL" của
// PhieuCan_DN), KHÔNG PHẢI Mã Khách Hàng — tên biến giữ nguyên theo code gốc để
// khỏi phải đổi cả chuỗi logic valK_Dich/valQ_Dich, nhưng nhãn hiển thị trên
// giao diện đã sửa thành "Mã Đại Lý - ĐL" để tránh nhầm lẫn khi nhập liệu thật.
// "maNG" là Mã Nguồn Gốc (cột O "NG"), tên gọi khớp đúng ý nghĩa.
function addManualPhieuCan(fields) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
  } catch (e) {
    return { status: "error", message: "Hệ thống đang bận xử lý một yêu cầu khác, vui lòng thử lại sau ít giây." };
  }

  try {
    if (!fields || !String(fields.soPhieu || "").trim()) return { status: "error", message: "Vui lòng nhập Số phiếu." };
    if (!String(fields.soXe || "").trim()) return { status: "error", message: "Vui lòng nhập Số xe." };

    const dateC = combineDateTime_(fields.ngayCan1, fields.gioCan1);
    const dateD = combineDateTime_(fields.ngayCan2, fields.gioCan2);
    if (!dateC) return { status: "error", message: "Ngày/giờ cân 1 không hợp lệ." };
    if (!dateD) return { status: "error", message: "Ngày/giờ cân 2 không hợp lệ." };

    const rawCan1 = parseFloat(fields.klCan1) || 0;
    const rawCan2 = parseFloat(fields.klCan2) || 0;
    const rawHang = parseFloat(fields.klHang) || 0;
    if (rawHang <= 0) return { status: "error", message: "Khối lượng hàng phải lớn hơn 0." };

    // Đồng bộ quy ước với step1_PreviewDraft: nếu số nhập vào < 70 thì hiểu là đơn vị Tấn -> quy đổi ra Kg
    const klCan1 = rawCan1 < 70 ? rawCan1 * 1000 : rawCan1;
    const klCan2 = rawCan2 < 70 ? rawCan2 * 1000 : rawCan2;
    const klHang = rawHang < 70 ? rawHang * 1000 : rawHang;

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const dataSheet = ss.getSheetByName(CONFIG.DATA_SHEET);
    const lastRow = dataSheet.getLastRow();
    const nam = dateC.getFullYear();
    const uniqueKeyMaCT = String(fields.soPhieu).trim() + "/" + nam + "/NK";

    if (lastRow > 1) {
      // Chỉ cần đọc cột V (22) để so trùng Mã Chứng Từ, không cần load cả 25 cột
      const existingKeys = dataSheet.getRange(2, 22, lastRow - 1, 1).getValues();
      for (let row of existingKeys) {
        if (String(row[0] || "").trim() === uniqueKeyMaCT) {
          return { status: "error", message: "Số chứng từ " + uniqueKeyMaCT + " đã tồn tại. Vui lòng kiểm tra lại Số phiếu / năm." };
        }
      }
    }

    const NUM_COLUMNS = 23; const now = new Date();
    let newRow = new Array(NUM_COLUMNS).fill("");
    const maKH = String(fields.maKH || "").toUpperCase().trim();
    const maNG = String(fields.maNG || "").toUpperCase().trim();
    const valK_Dich = maKH + "_" + maNG;
    const valQ_Dich = maKH + "_" + maNG + "_Y";

    newRow[0] = String(fields.soPhieu).trim();
    // FIX #3: Ghi đúng Ngày THUẦN vào cột "Ngày cân X" và Giờ THUẦN vào cột
    // "Giờ cân X", khớp đúng quy ước dữ liệu thật (xem toDateOnly_/toTimeOnly_).
    newRow[1] = toDateOnly_(dateC); newRow[2] = toTimeOnly_(dateC);
    newRow[3] = toDateOnly_(dateD); newRow[4] = toTimeOnly_(dateD);
    newRow[5] = String(fields.soXe).trim();
    newRow[6] = String(fields.soXe2 || "").trim(); // Cột G - Biển số 2 (tùy chọn, cho xe kéo/rơ-moóc)
    newRow[7] = klCan1; newRow[8] = klCan2; newRow[9] = klHang;
    newRow[10] = valK_Dich;
    newRow[11] = String(fields.khachHang || "").trim(); // Cột L - Khách hàng
    newRow[12] = "GK"; newRow[13] = maKH; newRow[14] = maNG;
    newRow[15] = "Y"; newRow[16] = valQ_Dich; newRow[17] = 0; newRow[18] = now;
    newRow[21] = uniqueKeyMaCT; newRow[22] = uniqueKeyMaCT;

    const startRow = dataSheet.getLastRow() + 1;
    dataSheet.getRange(startRow, 1, 1, NUM_COLUMNS).setValues([newRow]);
    dataSheet.getRange(startRow, 2, 1, 1).setNumberFormat("dd/MM/yyyy");
    dataSheet.getRange(startRow, 3, 1, 1).setNumberFormat("HH:mm:ss");
    dataSheet.getRange(startRow, 4, 1, 1).setNumberFormat("dd/MM/yyyy");
    dataSheet.getRange(startRow, 5, 1, 1).setNumberFormat("HH:mm:ss");

    // Gọi bản _core (không khóa) vì đang giữ khóa của addManualPhieuCan rồi
    const priceResult = runCalculatePrice_core();
    const finalMsg = "Đã thêm phiếu cân " + uniqueKeyMaCT + " | " + priceResult.message;
    logAudit_('MANUAL_ENTRY', 'OK', finalMsg);
    return { status: "success", message: finalMsg };
  } catch (e) {
    logAudit_('MANUAL_ENTRY', 'ERROR', e.toString());
    return { status: "error", message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function combineDateTime_(dateStr, timeStr) {
  if (!dateStr) return null;
  const d = new Date(dateStr + "T" + (timeStr || "00:00") + ":00");
  if (isNaN(d.getTime())) return null;
  return d;
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
    const sheet = ssMisa.getSheetByName(CONFIG.MISA_DST_SHEET);
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
  // FIX #2: khóa vì hàm này CLEAR + GHI ĐÈ toàn bộ sheet Update_MiSa_PC.
  // Nếu 2 người cùng bấm "Tạo data Misa" cùng lúc, không khóa có thể dẫn tới
  // xung đột đọc/ghi hoặc một phiên xóa dữ liệu ngay khi phiên kia đang đọc.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
  } catch (e) {
    return { status: "error", message: "Hệ thống đang bận xử lý một yêu cầu khác, vui lòng thử lại sau ít giây." };
  }

  try {
    const ssSourcePC = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.DATA_SHEET);
    const ssTarget   = SpreadsheetApp.openById(CONFIG.MISA_DST_ID).getSheetByName(CONFIG.MISA_DST_SHEET);
    const ssRef      = SpreadsheetApp.openById(CONFIG.SRC_FILE_ID);
    const ssDNTT     = SpreadsheetApp.openById(CONFIG.DNTT_FILE_ID);
    let start = fromDate ? new Date(fromDate + "T00:00:00+07:00") : new Date(new Date().getFullYear(), new Date().getMonth(), 1);
    let end = toDate ? new Date(toDate + "T23:59:59+07:00") : new Date();
    const mapNG = createSimpleMap_(ssRef.getSheetByName('DM_NG'), 1, 3);
    const dataKH = ssRef.getSheetByName('DM_KH').getDataRange().getValues();
    let mapKH = {}; for (let i = 1; i < dataKH.length; i++) { let k = String(dataKH[i][1]).trim(); if (k) mapKH[k] = { colC: dataKH[i][2], colD: dataKH[i][3], colE: dataKH[i][4] }; }
    const dataHDNCC = ssRef.getSheetByName('HD_NCC').getDataRange().getValues();
    let mapHDNCC = {}; for (let i = 1; i < dataHDNCC.length; i++) { let k = String(dataHDNCC[i][2]).trim(); if (k) mapHDNCC[k] = { colG: dataHDNCC[i][6], colE: dataHDNCC[i][4] }; }
    const dataDNTT = ssDNTT.getSheetByName(CONFIG.DNTT_SHEET).getDataRange().getValues();
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
  } catch (e) {
    return { status: "error", message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// FIX #2: bản CÓ KHÓA — dùng khi gọi ĐỘC LẬP (menu thủ công, trigger theo giờ...).
function runCalculatePrice() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
  } catch (e) {
    return { status: "error", message: "Hệ thống đang bận xử lý một yêu cầu khác, vui lòng thử lại sau ít giây." };
  }
  try {
    return runCalculatePrice_core();
  } finally {
    lock.releaseLock();
  }
}

// FIX #2: bản LÕI, KHÔNG khóa — dùng khi được gọi từ bên trong một hàm khác
// (như step1_ConfirmImport) mà đã tự khóa từ trước. Gọi runCalculatePrice()
// (bản có khóa) từ trong 1 hàm đang giữ khóa sẽ gây deadlock (tự chờ chính mình).
function runCalculatePrice_core() {
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

/*********************************************************
 * PHẦN 4: BÁO CÁO TỔNG HỢP CÂN & BÁO CÁO MISA (MENU 2)
 * - Bộ lọc dùng chung: fromDate, toDate, xe, khachHang, trangThai ('', 'da', 'chua')
 * - trangThai đối chiếu cột AA (ID_DNTT) trong PhieuCan_DN: rỗng = "Chưa lập ĐNTT"
 *********************************************************/

// Trả về danh sách xe / khách hàng / NCC để đổ vào dropdown filter phía client
function getFilterOptions() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.DATA_SHEET);
    const lastRow = sheet.getLastRow();
    let xeSet = new Set(); let khSet = new Set(); let daiLySet = new Set(); let nguonGocSet = new Set(); let maDonGiaSet = new Set();
    if (lastRow > 1) {
      const data = sheet.getRange(2, 1, lastRow - 1, 17).getValues(); // A..Q (cần tới cột Q=Mã ĐG, index16)
      data.forEach(row => {
        const xe = String(row[5] || "").trim(); if (xe) xeSet.add(xe);
        const kh = String(row[11] || "").trim(); if (kh) khSet.add(kh); // Cột L
        const dl = String(row[13] || "").trim(); if (dl) daiLySet.add(dl); // Cột N - ĐL (Đại lý)
        const ng = String(row[14] || "").trim(); if (ng) nguonGocSet.add(ng); // Cột O - NG (Nguồn gốc)
        const madg = String(row[16] || "").trim(); if (madg) maDonGiaSet.add(madg); // Cột Q - Mã ĐG
      });
    }
    let ncSet = new Set();
    try {
      const ssMisa = SpreadsheetApp.openById(CONFIG.MISA_DST_ID);
      const sheetMisa = ssMisa.getSheetByName(CONFIG.MISA_DST_SHEET);
      const lastRowMisa = sheetMisa.getLastRow();
      if (lastRowMisa > 1) {
        const dataMisa = sheetMisa.getRange(2, 12, lastRowMisa - 1, 1).getValues(); // Cột L (Tên NCC)
        dataMisa.forEach(row => { const nc = String(row[0] || "").trim(); if (nc) ncSet.add(nc); });
      }
    } catch (e) { /* bỏ qua nếu sheet Misa chưa có dữ liệu / chưa tồn tại */ }

    // Mã khối lượng lấy từ danh mục Ma_KL bên hệ thống Báo giá (không lưu trực tiếp
    // trên từng dòng PhieuCan_DN, mà được suy ra bằng cách so khớp KL hàng thực tế
    // với dải Min-Max của từng mã tại thời điểm xem báo cáo).
    let maKLList = [];
    try {
      const sheetKL = BG_ss_().getSheetByName(BAOGIA_CONFIG.MAKL_SHEET);
      const lastRowKL = sheetKL.getLastRow();
      if (lastRowKL > 1) {
        const dataKL = sheetKL.getRange(2, 2, lastRowKL - 1, 1).getValues(); // Cột B - Mã khối lượng
        maKLList = dataKL.map(r => String(r[0] || "").trim()).filter(Boolean).sort();
      }
    } catch (e) { /* bỏ qua nếu chưa có dữ liệu Ma_KL */ }

    return {
      status: "success",
      xeList: Array.from(xeSet).sort(),
      khachHangList: Array.from(khSet).sort(),
      ncList: Array.from(ncSet).sort(),
      daiLyList: Array.from(daiLySet).sort(),
      nguonGocList: Array.from(nguonGocSet).sort(),
      maDonGiaList: Array.from(maDonGiaSet).sort(),
      maKLList: maKLList
    };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

// Map: Mã Chứng Từ (cột V) -> đã lập ĐNTT hay chưa (dựa vào cột AA - ID_DNTT có rỗng hay không)
function buildDNTTStatusMap_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.DATA_SHEET);
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow > 1) {
    // Đọc từ cột V(22) đến cột AA(27): index0=V(MãCT) ... index5=AA(ID_DNTT)
    const data = sheet.getRange(2, 22, lastRow - 1, 6).getValues();
    data.forEach(row => {
      const key = String(row[0] || "").trim();
      const idDntt = String(row[5] || "").trim();
      if (key) map[key] = !!idDntt; // true = Đã lập ĐNTT, false = Chưa lập ĐNTT
    });
  }
  return map;
}

// filters = {fromDate, toDate, xe, khachHang, trangThai}
function getBaoCaoTongHop(filters) {
  try {
    filters = filters || {};
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.DATA_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: "success", data: [], summary: { soLuong: 0, tongKL: 0, tongTien: 0 } };

    const data = sheet.getRange(2, 1, lastRow - 1, 27).getValues(); // A..AA
    const start = filters.fromDate ? new Date(filters.fromDate + "T00:00:00+07:00") : null;
    const end = filters.toDate ? new Date(filters.toDate + "T23:59:59+07:00") : null;
    const xeFilter = String(filters.xe || "").trim();
    const khFilter = String(filters.khachHang || "").trim();
    const daiLyFilter = String(filters.daiLy || "").trim();
    const nguonGocFilter = String(filters.nguonGoc || "").trim();
    const trangThaiFilter = String(filters.trangThai || "").trim(); // '', 'da', 'chua'

    let result = []; let tongKL = 0; let tongTien = 0;

    data.forEach(row => {
      const ngayCan1 = row[1]; // Cột B - đã là Date object thật nhờ FIX #1
      if (!(ngayCan1 instanceof Date) || isNaN(ngayCan1.getTime())) return;
      if (start && ngayCan1 < start) return;
      if (end && ngayCan1 > end) return;

      const xe = String(row[5] || "").trim();
      if (xeFilter && xe !== xeFilter) return;

      const khachHang = String(row[11] || "").trim(); // Cột L
      if (khFilter && khachHang !== khFilter) return;

      const daiLy = String(row[13] || "").trim(); // Cột N - ĐL (Đại lý)
      if (daiLyFilter && daiLy !== daiLyFilter) return;

      const nguonGoc = String(row[14] || "").trim(); // Cột O - NG (Nguồn gốc)
      if (nguonGocFilter && nguonGoc !== nguonGocFilter) return;

      const idDntt = String(row[26] || "").trim(); // Cột AA
      const daLap = !!idDntt;
      if (trangThaiFilter === "da" && !daLap) return;
      if (trangThaiFilter === "chua" && daLap) return;

      const klHang = parseFloat(row[9]) || 0;
      const donGia = parseFloat(row[23]) || 0; // Cột X
      const thanhTien = parseFloat(row[25]) || 0; // Cột Z
      const trangThaiGia = String(row[24] || "").trim(); // Cột Y

      tongKL += klHang; tongTien += thanhTien;

      result.push({
        maChungTu: String(row[21] || "").trim(),
        soPhieu: row[0],
        ngayCan1: Utilities.formatDate(ngayCan1, "GMT+7", "dd/MM/yyyy"),
        gioCan1: (row[2] instanceof Date) ? Utilities.formatDate(row[2], "GMT+7", "HH:mm:ss") : "",
        ngayCan2: (row[3] instanceof Date) ? Utilities.formatDate(row[3], "GMT+7", "dd/MM/yyyy") : "",
        gioCan2: (row[4] instanceof Date) ? Utilities.formatDate(row[4], "GMT+7", "HH:mm:ss") : "",
        soXe: xe,
        soXe2: String(row[6] || "").trim(), // Cột G - Biển số 2 (xe kéo/rơ-moóc, nếu có)
        khachHang: khachHang,
        daiLy: daiLy,
        nguonGoc: nguonGoc,
        klCan1: parseFloat(row[7]) || 0,
        klCan2: parseFloat(row[8]) || 0,
        klHang: klHang,
        donGia: donGia,
        thanhTien: thanhTien,
        trangThaiGia: trangThaiGia,
        idDntt: idDntt, // Giá trị thô của cột AA (ID_DNTT), vd: "Đóng TT", hoặc rỗng
        // Hiển thị đúng nội dung thật trong cột AA khi có (vd "Đóng TT") thay vì
        // chỉ ghi chung chung "Đã lập ĐNTT", giúp thấy rõ tình trạng thanh toán thực tế.
        trangThaiThanhToan: daLap ? idDntt : "Chưa lập ĐNTT"
      });
    });

    result.sort((a, b) => (a.maChungTu > b.maChungTu ? 1 : -1));
    return { status: "success", data: result, summary: { soLuong: result.length, tongKL: tongKL, tongTien: tongTien } };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

/* ---------- Bảng tổng hợp cân THEO BÁO GIÁ (xem đơn giá áp dụng thế nào) ---------- */

// Map: Mã Báo Giá (vd "DT_ĐL_Y") -> Nội dung diễn giải (từ danh mục Ma_BaoGia)
function getMaBaoGiaLookup_() {
  const map = {};
  try {
    const sheet = BG_ss_().getSheetByName(BAOGIA_CONFIG.MA_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const data = sheet.getRange(2, 2, lastRow - 1, 5).getValues(); // B..F: maBaoGia, daiLy, nguonGoc, hinhAnh, noiDung
      data.forEach(row => {
        const ma = String(row[0] || "").trim();
        if (ma) map[ma] = String(row[4] || "").trim(); // Cột F - Nội dung
      });
    }
  } catch (e) { /* bỏ qua nếu chưa có dữ liệu Ma_BaoGia */ }
  return map;
}

// Danh sách dải khối lượng (Tấn) từ danh mục Ma_KL, dùng để suy ra "Mã KL" phù hợp
// với KL hàng thực tế của từng phiếu cân (PhieuCan_DN không lưu trực tiếp mã này).
function getMaKLBands_() {
  const bands = [];
  try {
    const sheet = BG_ss_().getSheetByName(BAOGIA_CONFIG.MAKL_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
      data.forEach(row => {
        const maKL = String(row[1] || "").trim();
        if (!maKL) return;
        bands.push({ maKL: maKL, minTan: (parseFloat(row[2]) || 0) / 1000, maxTan: (parseFloat(row[3]) || 0) / 1000 });
      });
    }
  } catch (e) { /* bỏ qua nếu chưa có dữ liệu Ma_KL */ }
  return bands;
}

// Khớp đúng quy ước đang dùng trong runCalculatePrice_core: min < KL <= max
function findMaKLForTan_(bands, klTan) {
  for (let i = 0; i < bands.length; i++) {
    if (klTan > bands[i].minTan && klTan <= bands[i].maxTan) return bands[i].maKL;
  }
  return "";
}

// filters = {fromDate, toDate, xe, khachHang, daiLy, nguonGoc, trangThai, maDonGia, maKL}
function getBaoCaoDonGia(filters) {
  try {
    filters = filters || {};
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.DATA_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: "success", data: [], summary: { soLuong: 0, tongTien: 0 } };

    const data = sheet.getRange(2, 1, lastRow - 1, 27).getValues(); // A..AA
    const start = filters.fromDate ? new Date(filters.fromDate + "T00:00:00+07:00") : null;
    const end = filters.toDate ? new Date(filters.toDate + "T23:59:59+07:00") : null;
    const xeFilter = String(filters.xe || "").trim();
    const khFilter = String(filters.khachHang || "").trim();
    const daiLyFilter = String(filters.daiLy || "").trim();
    const nguonGocFilter = String(filters.nguonGoc || "").trim();
    const trangThaiFilter = String(filters.trangThai || "").trim();
    const maDonGiaFilter = String(filters.maDonGia || "").trim();
    const maKLFilter = String(filters.maKL || "").trim();

    const noiDungMap = getMaBaoGiaLookup_();
    const klBands = getMaKLBands_();

    let result = []; let tongTien = 0;

    data.forEach(row => {
      const ngayCan1 = row[1];
      if (!(ngayCan1 instanceof Date) || isNaN(ngayCan1.getTime())) return;
      if (start && ngayCan1 < start) return;
      if (end && ngayCan1 > end) return;

      const xe = String(row[5] || "").trim();
      if (xeFilter && xe !== xeFilter) return;

      const khachHang = String(row[11] || "").trim();
      if (khFilter && khachHang !== khFilter) return;

      const daiLy = String(row[13] || "").trim();
      if (daiLyFilter && daiLy !== daiLyFilter) return;

      const nguonGoc = String(row[14] || "").trim();
      if (nguonGocFilter && nguonGoc !== nguonGocFilter) return;

      const idDntt = String(row[26] || "").trim();
      const daLap = !!idDntt;
      if (trangThaiFilter === "da" && !daLap) return;
      if (trangThaiFilter === "chua" && daLap) return;

      const maDonGia = String(row[16] || "").trim(); // Cột Q - Mã ĐG
      if (maDonGiaFilter && maDonGia !== maDonGiaFilter) return;

      const klHang = parseFloat(row[9]) || 0;
      const klTan = klHang / 1000;
      const maKLMatched = findMaKLForTan_(klBands, klTan);
      if (maKLFilter && maKLMatched !== maKLFilter) return;

      const giaGoc = parseFloat(row[19]) || 0;        // Cột T - ĐG_AD (giá gốc tìm được từ bảng báo giá)
      const dieuChinh = parseFloat(row[17]) || 0;      // Cột R - Giảm giá / điều chỉnh
      const donGiaApDung = parseFloat(row[23]) || 0;   // Cột X - Đơn giá_TC (giá cuối cùng áp dụng)
      const thanhTien = parseFloat(row[25]) || 0;      // Cột Z
      const trangThaiGia = String(row[24] || "").trim(); // Cột Y

      tongTien += thanhTien;

      result.push({
        maChungTu: String(row[21] || "").trim(),
        ngayCan1: Utilities.formatDate(ngayCan1, "GMT+7", "dd/MM/yyyy"),
        soXe: xe,
        khachHang: khachHang,
        daiLy: daiLy,
        nguonGoc: nguonGoc,
        maDonGia: maDonGia,
        maKL: maKLMatched,
        klHang: klHang,
        giaGoc: giaGoc,
        dieuChinh: dieuChinh,
        donGiaApDung: donGiaApDung,
        dienGiai: noiDungMap[maDonGia] || (maDonGia ? "(Không tìm thấy trong danh mục Mã Báo Giá)" : ""),
        trangThaiGia: trangThaiGia,
        thanhTien: thanhTien
      });
    });

    result.sort((a, b) => (a.maChungTu > b.maChungTu ? 1 : -1));
    return { status: "success", data: result, summary: { soLuong: result.length, tongTien: tongTien } };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

function exportBaoCaoDonGiaExcel(filters) {
  try {
    const rep = getBaoCaoDonGia(filters);
    if (rep.status !== "success") return rep;
    if (rep.data.length === 0) return { status: "error", message: "Không có dữ liệu phù hợp bộ lọc để xuất." };
    const headers = ["Mã Chứng Từ", "Ngày Cân", "Số Xe", "Khách Hàng", "Đại Lý", "Nguồn Gốc", "Mã Đơn Giá", "Mã KL", "KL Hàng (kg)", "Giá Gốc (ĐG_AD)", "Điều Chỉnh", "Đơn Giá Áp Dụng", "Diễn Giải Đơn Giá", "Trạng Thái Giá", "Thành Tiền"];
    const rows = rep.data.map(r => [r.maChungTu, r.ngayCan1, r.soXe, r.khachHang, r.daiLy, r.nguonGoc, r.maDonGia, r.maKL, r.klHang, r.giaGoc, r.dieuChinh, r.donGiaApDung, r.dienGiai, r.trangThaiGia, r.thanhTien]);
    const tempSS = createTempSheetForExport_("BaoCao_DonGia_" + Utilities.formatDate(new Date(), "GMT+7", "ddMM_HHmm"), headers, rows, [9, 10, 11, 12, 15]);
    logAudit_('EXPORT_EXCEL', 'OK', 'Xuất báo cáo tổng hợp cân theo báo giá, ' + rep.data.length + ' dòng.');
    return { status: "success", url: getExportUrl_(tempSS, "xlsx") };
  } catch (e) { logAudit_('EXPORT_EXCEL', 'ERROR', e.toString()); return { status: "error", message: e.toString() }; }
}

function exportBaoCaoDonGiaPDF(filters) {
  try {
    const rep = getBaoCaoDonGia(filters);
    if (rep.status !== "success") return rep;
    if (rep.data.length === 0) return { status: "error", message: "Không có dữ liệu phù hợp bộ lọc để xuất." };
    const headers = ["Mã Chứng Từ", "Ngày Cân", "Số Xe", "Khách Hàng", "Đại Lý", "Nguồn Gốc", "Mã ĐG", "Mã KL", "Giá Gốc", "Điều Chỉnh", "Đơn Giá AD", "Diễn Giải", "Trạng Thái", "Thành Tiền"];
    const rows = rep.data.map(r => [r.maChungTu, r.ngayCan1, r.soXe, r.khachHang, r.daiLy, r.nguonGoc, r.maDonGia, r.maKL, r.giaGoc, r.dieuChinh, r.donGiaApDung, r.dienGiai, r.trangThaiGia, r.thanhTien]);
    const tempSS = createTempSheetForExport_("BaoCao_DonGia_PDF_" + Utilities.formatDate(new Date(), "GMT+7", "ddMM_HHmm"), headers, rows, [8, 9, 10, 13]);
    logAudit_('EXPORT_PDF', 'OK', 'Xuất PDF báo cáo tổng hợp cân theo báo giá, ' + rep.data.length + ' dòng.');
    return { status: "success", url: getExportUrl_(tempSS, "pdf", false) };
  } catch (e) { logAudit_('EXPORT_PDF', 'ERROR', e.toString()); return { status: "error", message: e.toString() }; }
}

// filters = {fromDate, toDate, xe, khachHang, trangThai}
function getBaoCaoMisa(filters) {
  try {
    filters = filters || {};
    const ss = SpreadsheetApp.openById(CONFIG.MISA_DST_ID);
    const sheet = ss.getSheetByName(CONFIG.MISA_DST_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: "success", data: [], summary: { soLuong: 0, tongKL: 0, tongTien: 0 } };

    const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues(); // A..L
    const dntt = buildDNTTStatusMap_();

    const start = filters.fromDate ? new Date(filters.fromDate + "T00:00:00+07:00") : null;
    const end = filters.toDate ? new Date(filters.toDate + "T23:59:59+07:00") : null;
    const xeFilter = String(filters.xe || "").trim();
    const khFilter = String(filters.khachHang || "").trim();
    const trangThaiFilter = String(filters.trangThai || "").trim();

    let result = []; let tongKL = 0; let tongTien = 0;

    data.forEach(row => {
      const ngay = parseDate(row[0]); // Cột A
      if (!ngay) return;
      if (start && ngay < start) return;
      if (end && ngay > end) return;

      const xe = String(row[3] || "").trim(); // Cột D
      if (xeFilter && xe !== xeFilter) return;

      const tenNCC = String(row[11] || "").trim(); // Cột L
      if (khFilter && tenNCC !== khFilter) return;

      const maChungTu = String(row[2] || "").trim(); // Cột C (chính là Mã Chứng Từ dùng chung với PhieuCan_DN)
      const daLap = dntt.hasOwnProperty(maChungTu) ? dntt[maChungTu] : false;
      if (trangThaiFilter === "da" && !daLap) return;
      if (trangThaiFilter === "chua" && daLap) return;

      const khoiLuong = parseFloat(row[4]) || 0; // Cột E (Tấn)
      const donGia = parseFloat(row[5]) || 0; // Cột F
      const thanhTien = parseFloat(row[6]) || 0; // Cột G

      tongKL += khoiLuong; tongTien += thanhTien;

      result.push({
        maChungTu: maChungTu,
        ngay: Utilities.formatDate(ngay, "GMT+7", "dd/MM/yyyy"),
        soXe: xe,
        khoiLuong: khoiLuong,
        donGia: donGia,
        thanhTien: thanhTien,
        tenNCC: tenNCC,
        trangThaiThanhToan: daLap ? "Đã lập ĐNTT" : "Chưa lập ĐNTT"
      });
    });

    result.sort((a, b) => (a.maChungTu > b.maChungTu ? 1 : -1));
    return { status: "success", data: result, summary: { soLuong: result.length, tongKL: tongKL, tongTien: tongTien } };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

// ---- Helper dùng chung để tạo file tạm phục vụ xuất Excel / PDF ----
function createTempSheetForExport_(title, headers, rows, numberFormatCols) {
  const tempSS = SpreadsheetApp.create(title);
  const tempSheet = tempSS.getSheets()[0];
  tempSheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold").setBackground("#1B4332").setFontColor("#FFFFFF").setHorizontalAlignment("center");
  if (rows.length > 0) {
    tempSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    (numberFormatCols || []).forEach(c => tempSheet.getRange(2, c, rows.length, 1).setNumberFormat("#,##0"));
  }
  tempSheet.autoResizeColumns(1, headers.length);
  tempSheet.setFrozenRows(1);
  const tempFile = DriveApp.getFileById(tempSS.getId());
  DriveApp.getFolderById(CONFIG.FOLDER_DONE).addFile(tempFile);
  DriveApp.getRootFolder().removeFile(tempFile);
  return tempSS;
}

function getExportUrl_(tempSS, format, opt_portrait) {
  const gid = tempSS.getSheets()[0].getSheetId();
  const base = "https://docs.google.com/spreadsheets/d/" + tempSS.getId() + "/export";
  if (format === "pdf") {
    const portrait = opt_portrait === true;
    return base + "?format=pdf&gid=" + gid + "&size=A4&portrait=" + portrait +
      "&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenumbers=true" +
      "&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4";
  }
  return base + "?format=xlsx";
}

function exportBaoCaoTongHopExcel(filters) {
  try {
    const rep = getBaoCaoTongHop(filters);
    if (rep.status !== "success") return rep;
    if (rep.data.length === 0) return { status: "error", message: "Không có dữ liệu phù hợp bộ lọc để xuất." };
    const headers = ["Mã Chứng Từ", "Số Phiếu", "Ngày Cân 1", "Giờ Cân 1", "Ngày Cân 2", "Giờ Cân 2", "Số Xe", "Biển Số 2", "Khách Hàng", "KL Cân 1 (kg)", "KL Cân 2 (kg)", "KL Hàng (kg)", "Đơn Giá", "Thành Tiền", "Trạng Thái Giá", "Trạng Thái Thanh Toán"];
    const rows = rep.data.map(r => [r.maChungTu, r.soPhieu, r.ngayCan1, r.gioCan1, r.ngayCan2, r.gioCan2, r.soXe, r.soXe2, r.khachHang, r.klCan1, r.klCan2, r.klHang, r.donGia, r.thanhTien, r.trangThaiGia, r.trangThaiThanhToan]);
    const tempSS = createTempSheetForExport_("BaoCao_TongHopCan_" + Utilities.formatDate(new Date(), "GMT+7", "ddMM_HHmm"), headers, rows, [10, 11, 12, 13, 14]);
    logAudit_('EXPORT_EXCEL', 'OK', 'Xuất báo cáo tổng hợp cân, ' + rep.data.length + ' dòng.');
    return { status: "success", url: getExportUrl_(tempSS, "xlsx") };
  } catch (e) { logAudit_('EXPORT_EXCEL', 'ERROR', e.toString()); return { status: "error", message: e.toString() }; }
}

function exportBaoCaoTongHopPDF(filters) {
  try {
    const rep = getBaoCaoTongHop(filters);
    if (rep.status !== "success") return rep;
    if (rep.data.length === 0) return { status: "error", message: "Không có dữ liệu phù hợp bộ lọc để xuất." };
    const headers = ["Mã Chứng Từ", "Số Phiếu", "Ngày Cân 1", "Giờ Cân 1", "Ngày Cân 2", "Giờ Cân 2", "Số Xe", "Biển Số 2", "Khách Hàng", "KL Cân1(kg)", "KL Cân2(kg)", "KL Hàng(kg)", "Đơn Giá", "Thành Tiền", "Trạng Thái TT"];
    const rows = rep.data.map(r => [r.maChungTu, r.soPhieu, r.ngayCan1, r.gioCan1, r.ngayCan2, r.gioCan2, r.soXe, r.soXe2, r.khachHang, r.klCan1, r.klCan2, r.klHang, r.donGia, r.thanhTien, r.trangThaiThanhToan]);
    const tempSS = createTempSheetForExport_("BaoCao_TongHopCan_PDF_" + Utilities.formatDate(new Date(), "GMT+7", "ddMM_HHmm"), headers, rows, [10, 11, 12, 13, 14]);
    logAudit_('EXPORT_PDF', 'OK', 'Xuất PDF báo cáo tổng hợp cân, ' + rep.data.length + ' dòng.');
    return { status: "success", url: getExportUrl_(tempSS, "pdf", false) }; // landscape cho bảng nhiều cột
  } catch (e) { logAudit_('EXPORT_PDF', 'ERROR', e.toString()); return { status: "error", message: e.toString() }; }
}

function exportBaoCaoMisaExcel(filters) {
  try {
    const rep = getBaoCaoMisa(filters);
    if (rep.status !== "success") return rep;
    if (rep.data.length === 0) return { status: "error", message: "Không có dữ liệu phù hợp bộ lọc để xuất." };
    const headers = ["Mã Chứng Từ", "Ngày", "Số Xe", "Khối Lượng (Tấn)", "Đơn Giá", "Thành Tiền", "Tên NCC", "Trạng Thái Thanh Toán"];
    const rows = rep.data.map(r => [r.maChungTu, r.ngay, r.soXe, r.khoiLuong, r.donGia, r.thanhTien, r.tenNCC, r.trangThaiThanhToan]);
    const tempSS = createTempSheetForExport_("BaoCao_Misa_" + Utilities.formatDate(new Date(), "GMT+7", "ddMM_HHmm"), headers, rows, [4, 5, 6]);
    logAudit_('EXPORT_EXCEL', 'OK', 'Xuất báo cáo Misa, ' + rep.data.length + ' dòng.');
    return { status: "success", url: getExportUrl_(tempSS, "xlsx") };
  } catch (e) { logAudit_('EXPORT_EXCEL', 'ERROR', e.toString()); return { status: "error", message: e.toString() }; }
}

function exportBaoCaoMisaPDF(filters) {
  try {
    const rep = getBaoCaoMisa(filters);
    if (rep.status !== "success") return rep;
    if (rep.data.length === 0) return { status: "error", message: "Không có dữ liệu phù hợp bộ lọc để xuất." };
    const headers = ["Mã Chứng Từ", "Ngày", "Số Xe", "KL(Tấn)", "Đơn Giá", "Thành Tiền", "Tên NCC", "Trạng Thái TT"];
    const rows = rep.data.map(r => [r.maChungTu, r.ngay, r.soXe, r.khoiLuong, r.donGia, r.thanhTien, r.tenNCC, r.trangThaiThanhToan]);
    const tempSS = createTempSheetForExport_("BaoCao_Misa_PDF_" + Utilities.formatDate(new Date(), "GMT+7", "ddMM_HHmm"), headers, rows, [4, 5, 6]);
    logAudit_('EXPORT_PDF', 'OK', 'Xuất PDF báo cáo Misa, ' + rep.data.length + ' dòng.');
    return { status: "success", url: getExportUrl_(tempSS, "pdf", true) };
  } catch (e) { logAudit_('EXPORT_PDF', 'ERROR', e.toString()); return { status: "error", message: e.toString() }; }
}

// In một phiếu cân riêng lẻ ra PDF (khổ A5), tra theo Mã Chứng Từ
// Chuyển số thành chữ tiếng Việt (dùng cho dòng "Số tiền bằng chữ" trên phiếu nhập kho)
function soThanhChu_(so) {
  so = Math.round(Math.abs(so || 0));
  if (so === 0) return "Không đồng";
  const chuSo = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"];
  const donVi = ["", "nghìn", "triệu", "tỷ"];

  function docBaSo(n, coTramDauKhong) {
    const tram = Math.floor(n / 100);
    const chuc = Math.floor((n % 100) / 10);
    const donvi = n % 10;
    let s = "";
    if (tram > 0 || coTramDauKhong) s += chuSo[tram] + " trăm ";
    if (chuc === 0) { if (donvi > 0 && (tram > 0 || coTramDauKhong)) s += "lẻ "; }
    else if (chuc === 1) s += "mười ";
    else s += chuSo[chuc] + " mươi ";
    if (donvi === 1 && chuc >= 2) s += "mốt";
    else if (donvi === 5 && chuc >= 1) s += "lăm";
    else if (donvi > 0) s += chuSo[donvi];
    return s.trim();
  }

  const nhom = [];
  let n = so;
  while (n > 0) { nhom.push(n % 1000); n = Math.floor(n / 1000); }

  let ketQua = "";
  for (let i = nhom.length - 1; i >= 0; i--) {
    if (nhom[i] === 0) continue;
    const coTramDauKhong = i < nhom.length - 1; // các nhóm sau nhóm đầu tiên luôn đọc đủ hàng trăm
    ketQua += docBaSo(nhom[i], coTramDauKhong) + " " + donVi[i] + " ";
  }
  ketQua = ketQua.replace(/\s+/g, " ").trim();
  return ketQua.charAt(0).toUpperCase() + ketQua.slice(1) + " đồng";
}

// In "PHIẾU NHẬP KHO" theo đúng bố cục chứng từ nhập kho thực tế (khác với báo
// cáo/bảng dữ liệu thô): có tiêu đề công ty, bảng hàng hóa, số tiền bằng chữ,
// và 4 cột chữ ký (Người giao hàng / Thủ kho / Kế toán / Người lập phiếu).
function exportPhieuCanPDF(maChungTu) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.DATA_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: "error", message: "Không có dữ liệu." };
    const data = sheet.getRange(2, 1, lastRow - 1, 27).getValues();
    const found = data.find(row => String(row[21] || "").trim() === String(maChungTu || "").trim());
    if (!found) return { status: "error", message: "Không tìm thấy phiếu cân " + maChungTu };

    const tempSS = SpreadsheetApp.create("PhieuNhapKho_" + String(maChungTu).replace(/\//g, "_"));
    const sh = tempSS.getSheets()[0];
    const NUM_COLS = 6;
    sh.setColumnWidths(1, NUM_COLS, 100);
    sh.setColumnWidth(2, 150);

    const ngayCan1 = found[1] instanceof Date ? found[1] : null;
    const ngayCan2 = found[3] instanceof Date ? found[3] : null;
    const gioCan1 = found[2] instanceof Date ? Utilities.formatDate(found[2], "GMT+7", "HH:mm:ss") : "";
    const gioCan2 = found[4] instanceof Date ? Utilities.formatDate(found[4], "GMT+7", "HH:mm:ss") : "";
    const soXe2 = String(found[6] || "").trim();
    const idDntt = String(found[26] || "").trim();
    const klHang = parseFloat(found[9]) || 0;
    const donGia = parseFloat(found[23]) || 0;
    const thanhTien = parseFloat(found[25]) || 0;
    const ngayLap = ngayCan2 || ngayCan1 || new Date();

    let r = 1;
    sh.getRange(r, 1, 1, NUM_COLS).merge().setValue(COMPANY_NAME).setFontWeight("bold").setFontSize(12); r++;
    sh.getRange(r, 1, 1, NUM_COLS).merge().setValue("Địa chỉ: ................................................................").setFontSize(9).setFontColor("#5B6259"); r += 2;

    sh.getRange(r, 1, 1, NUM_COLS).merge().setValue("PHIẾU NHẬP KHO").setFontWeight("bold").setFontSize(16).setHorizontalAlignment("center"); r++;
    sh.getRange(r, 1, 1, NUM_COLS).merge()
      .setValue("Ngày " + Utilities.formatDate(ngayLap, "GMT+7", "dd") + " tháng " + Utilities.formatDate(ngayLap, "GMT+7", "MM") + " năm " + Utilities.formatDate(ngayLap, "GMT+7", "yyyy"))
      .setFontStyle("italic").setHorizontalAlignment("center"); r += 2;

    const thongTin = [
      ["Số phiếu cân", found[0], "Mã chứng từ", maChungTu],
      ["Khách hàng giao hàng", found[11], "Số xe", found[5] + (soXe2 ? " / " + soXe2 : "")],
      ["Thời gian cân lần 1", (ngayCan1 ? Utilities.formatDate(ngayCan1, "GMT+7", "dd/MM/yyyy") : "") + " " + gioCan1, "Thời gian cân lần 2", (ngayCan2 ? Utilities.formatDate(ngayCan2, "GMT+7", "dd/MM/yyyy") : "") + " " + gioCan2],
      ["Nhập tại kho", "Kho nguyên liệu " + COMPANY_NAME, "Trạng thái ĐNTT", idDntt || "Chưa lập ĐNTT"]
    ];
    thongTin.forEach(row => {
      sh.getRange(r, 1).setValue(row[0]).setFontWeight("bold");
      sh.getRange(r, 2, 1, 2).merge().setValue(row[1]);
      sh.getRange(r, 4).setValue(row[2]).setFontWeight("bold");
      sh.getRange(r, 5, 1, 2).merge().setValue(row[3]);
      r++;
    });
    r++;

    const headerRow = r;
    sh.getRange(r, 1, 1, NUM_COLS).setValues([["STT", "Tên hàng hóa", "ĐVT", "Khối lượng", "Đơn giá", "Thành tiền"]])
      .setFontWeight("bold").setBackground("#1B4332").setFontColor("#FFFFFF").setHorizontalAlignment("center")
      .setBorder(true, true, true, true, true, true);
    r++;
    const goodsRow = r;
    sh.getRange(r, 1, 1, NUM_COLS).setValues([[1, "Gỗ nguyên liệu (Nguồn gốc: " + (found[10] || "") + ")", "Kg", klHang, donGia, thanhTien]])
      .setBorder(true, true, true, true, true, true);
    sh.getRange(r, 4, 1, 3).setNumberFormat("#,##0");
    r++;
    sh.getRange(r, 1, 1, 3).merge().setValue("CỘNG").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
    sh.getRange(r, 4).setValue(klHang).setFontWeight("bold").setNumberFormat("#,##0").setBorder(true, true, true, true, true, true);
    sh.getRange(r, 5).setBorder(true, true, true, true, true, true);
    sh.getRange(r, 6).setValue(thanhTien).setFontWeight("bold").setNumberFormat("#,##0").setBorder(true, true, true, true, true, true);
    r += 2;

    sh.getRange(r, 1, 1, NUM_COLS).merge().setValue("Số tiền bằng chữ: " + soThanhChu_(thanhTien)).setFontStyle("italic");
    r += 2;

    const chuKy = ["Người giao hàng", "Thủ kho", "Kế toán", "Người lập phiếu"];
    // Chia 4 mục ký tên trên NUM_COLS=6 cột: (1-2)=Người giao hàng, (3)=Thủ kho, (4)=Kế toán, (5-6)=Người lập phiếu
    sh.getRange(r, 1, 1, 2).merge().setValue(chuKy[0]).setFontWeight("bold").setHorizontalAlignment("center");
    sh.getRange(r, 3, 1, 1).merge().setValue(chuKy[1]).setFontWeight("bold").setHorizontalAlignment("center");
    sh.getRange(r, 4, 1, 1).merge().setValue(chuKy[2]).setFontWeight("bold").setHorizontalAlignment("center");
    sh.getRange(r, 5, 1, 2).merge().setValue(chuKy[3]).setFontWeight("bold").setHorizontalAlignment("center");
    r++;
    sh.getRange(r, 1, 1, 2).merge().setValue("(Ký, họ tên)").setFontStyle("italic").setFontSize(9).setHorizontalAlignment("center");
    sh.getRange(r, 3, 1, 1).merge().setValue("(Ký, họ tên)").setFontStyle("italic").setFontSize(9).setHorizontalAlignment("center");
    sh.getRange(r, 4, 1, 1).merge().setValue("(Ký, họ tên)").setFontStyle("italic").setFontSize(9).setHorizontalAlignment("center");
    sh.getRange(r, 5, 1, 2).merge().setValue("(Ký, họ tên)").setFontStyle("italic").setFontSize(9).setHorizontalAlignment("center");
    r += 5; // chừa khoảng trống để ký tay

    const tempFile = DriveApp.getFileById(tempSS.getId());
    DriveApp.getFolderById(CONFIG.FOLDER_DONE).addFile(tempFile);
    DriveApp.getRootFolder().removeFile(tempFile);

    const gid = sh.getSheetId();
    const url = "https://docs.google.com/spreadsheets/d/" + tempSS.getId() + "/export?format=pdf&gid=" + gid +
      "&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenumbers=false" +
      "&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5";
    return { status: "success", url: url };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  if (typeof v === "number") return new Date(Math.round((v - 25569) * 86400 * 1000));
  if (typeof v === "string") {
    // Nhánh này chỉ còn phục vụ dữ liệu CŨ đã tồn tại trước khi áp dụng FIX #1
    // (được ghi dưới dạng chuỗi). Dữ liệu mới ghi từ nay là Date object thật
    // nên sẽ luôn rơi vào nhánh "v instanceof Date" ở trên, không đi qua đây.
    let p = v.includes("-") ? v.split("-") : v.split("/");
    if (p.length === 3) return p[0].length === 4 ? new Date(p[0], p[1]-1, p[2]) : new Date(p[2], p[1]-1, p[0]);
  }
  return null;
}

// FIX #4 (NGHIÊM TRỌNG): google.script.run TỰ ĐỘNG chuyển Date object thành
// chuỗi ISO 8601 ("yyyy-MM-ddTHH:mm:ss.sssZ") mỗi khi truyền qua lại giữa
// client <-> server. Cụ thể: step1_PreviewDraft trả về rawRowData (chứa Date
// object thật) cho trình duyệt; khi người dùng bấm "Xác nhận", trình duyệt gửi
// NGUYÊN rawRowData đó về step1_ConfirmImport — nhưng lúc này Date đã bị
// Apps Script tự chuyển thành chuỗi ISO. Bản toDateObj cũ chỉ biết phân tích
// chuỗi kiểu "DD/MM/YYYY..." (dành cho dữ liệu text thật từ file nguồn), nên
// khi gặp chuỗi ISO sẽ đọc SAI HOÀN TOÀN vị trí ngày/tháng/năm (đã kiểm chứng:
// có thể lệch tới hàng chục năm). Fix: nhận diện chuỗi ISO trước và dùng
// new Date() phân giải chuẩn, chỉ dùng regex thủ công cho các chuỗi không phải ISO.
function toDateObj(val) {
  if (val instanceof Date) return val;
  if (typeof val === 'string' && val.trim() !== "") {
    const trimmed = val.trim();
    if (/^\d{4}-\d{2}-\d{2}T/.test(trimmed)) {
      const d = new Date(trimmed);
      return isNaN(d.getTime()) ? null : d;
    }
    const parts = trimmed.match(/(\d+)/g);
    if (parts && parts.length >= 3) return new Date(parts[2], parts[1]-1, parts[0], parts[3]||0, parts[4]||0, parts[5]||0);
  }
  return null;
}

// FIX #3: Dữ liệu thật cho thấy "Ngày cân X" luôn là NGÀY THUẦN (giờ=00:00:00)
// và "Giờ cân X" luôn là GIỜ THUẦN (Sheets lưu Time dựa trên mốc gốc 30/12/1899).
// Trước đây step1_ConfirmImport/addManualPhieuCan ghi CHUNG 1 datetime đầy đủ vào
// cả 2 cột rồi chỉ đổi định dạng hiển thị — vẫn hiển thị đúng, nhưng giá trị GỐC
// lưu trong ô bị lệch chuẩn (cột "Ngày" vẫn mang theo giờ:phút:giây bên trong),
// có thể gây sai lệch nếu có công thức/pivot/QUERY khác dựa vào đúng kiểu dữ liệu.
// 2 helper dưới đây tách rõ ràng để khớp đúng quy ước dữ liệu thật.
function toDateOnly_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function toTimeOnly_(d) { return new Date(1899, 11, 30, d.getHours(), d.getMinutes(), d.getSeconds()); }

// Ghi log vào sheet "Audit" đã có sẵn trong hệ thống (Timestamp, Action, Status, Message)
// để nhất quán với quy ước audit hiện có của hệ thống PhieuCan_DN.
function logAudit_(action, status, message) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.AUDIT_SHEET);
    if (!sheet) return; // Nếu sheet Audit không tồn tại thì bỏ qua, không làm hỏng luồng chính
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, 4).setValues([[new Date(), action, status, message]]);
  } catch (e) { /* Không để lỗi ghi log làm hỏng thao tác chính */ }
}
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

/*********************************************************
 * PHẦN 5: QUẢN LÝ BÁO GIÁ (MENU 3)
 * - Cùng spreadsheet với CONFIG.URL_BAO_GIA / CONFIG.SHEET_BAO_GIA đã dùng
 *   trong runCalculatePrice_core để tra giá cân hàng — giữ NGUYÊN cấu trúc
 *   sheet gốc: Baogia_DN (log nhập), QL_BaoGia (đầu phiếu báo giá),
 *   Baogia_DN_FINAL (chỉ dòng còn hiệu lực), Baogia_DN_SAVE (toàn bộ lịch sử),
 *   Ma_BaoGia (danh mục mã báo giá), Ma_KL (danh mục mã khối lượng).
 * - Toàn bộ logic tính "Còn/Chưa/Hết hiệu lực" (BG_coreLogicProcessor_) được
 *   giữ NGUYÊN VẸN như code gốc để không thay đổi cách tính giá đang chạy.
 * - Cấu hình BAOGIA_CONFIG nằm trong file Config.gs.
 *********************************************************/

function BG_ss_() { return SpreadsheetApp.openById(BAOGIA_CONFIG.SPREADSHEET_ID); }

/* ---------- 5.1 Danh mục mã báo giá (Ma_BaoGia) ---------- */
function BG_getMaBaoGiaList() {
  try {
    const sheet = BG_ss_().getSheetByName(BAOGIA_CONFIG.MA_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: "success", data: [] };
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const result = data
      .map(r => ({ stt: r[0], maBaoGia: r[1], daiLy: r[2], nguonGoc: r[3], hinhAnh: r[4], noiDung: r[5], maDLNG: r[6] }))
      .filter(r => String(r.maBaoGia || "").trim() !== "");
    return { status: "success", data: result };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

// fields: {daiLyMa, daiLyTen, nguonGocMa, nguonGocTen, hinhAnh:'Y'|'N'}
function BG_addMaBaoGia(fields) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(CONFIG.LOCK_TIMEOUT_MS); } catch (e) {
    return { status: "error", message: "Hệ thống đang bận, vui lòng thử lại." };
  }
  try {
    const daiLyMa = String(fields.daiLyMa || "").trim().toUpperCase();
    const daiLyTen = String(fields.daiLyTen || "").trim();
    const nguonGocMa = String(fields.nguonGocMa || "").trim().toUpperCase();
    const nguonGocTen = String(fields.nguonGocTen || "").trim();
    const hinhAnh = String(fields.hinhAnh || "Y").trim().toUpperCase() === "N" ? "N" : "Y";
    if (!daiLyMa || !daiLyTen) return { status: "error", message: "Vui lòng nhập Mã và Tên đầy đủ Đại lý." };
    if (!nguonGocMa || !nguonGocTen) return { status: "error", message: "Vui lòng nhập Mã và Tên đầy đủ Nguồn gốc." };

    const maBaoGia = daiLyMa + "_" + nguonGocMa + "_" + hinhAnh;
    const maDLNG = daiLyMa + "_" + nguonGocMa;
    const noiDung = "Đại lý " + daiLyTen + " Nguồn gốc " + nguonGocTen + ", " + (hinhAnh === "Y" ? "có" : "không") + " hình ảnh";

    const sheet = BG_ss_().getSheetByName(BAOGIA_CONFIG.MA_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const existing = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // Cột B - Mã Báo giá
      for (let row of existing) {
        if (String(row[0] || "").trim().toUpperCase() === maBaoGia) {
          return { status: "error", message: "Mã báo giá " + maBaoGia + " đã tồn tại trong danh mục." };
        }
      }
    }

    const stt = BG_nextMaBaoGiaCode_(sheet, lastRow);
    sheet.getRange(lastRow + 1, 1, 1, 7).setValues([[stt, maBaoGia, daiLyMa, nguonGocMa, hinhAnh, noiDung, maDLNG]]);
    return { status: "success", message: "Đã thêm mã báo giá " + maBaoGia + " vào danh mục.", maBaoGia: maBaoGia };
  } catch (e) { return { status: "error", message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function BG_nextMaBaoGiaCode_(sheet, lastRow) {
  const yy = String(new Date().getFullYear()).slice(-2);
  const prefix = "BG" + yy + "-";
  let maxSeq = 0;
  if (lastRow > 1) {
    const existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    existing.forEach(row => {
      const v = String(row[0] || "").trim();
      if (v.indexOf(prefix) === 0) {
        const seq = parseInt(v.substring(prefix.length), 10);
        if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
      }
    });
  }
  return prefix + String(maxSeq + 1).padStart(3, "0");
}

function BG_deleteMaBaoGia(maBaoGia) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(CONFIG.LOCK_TIMEOUT_MS); } catch (e) {
    return { status: "error", message: "Hệ thống đang bận, vui lòng thử lại." };
  }
  try {
    const sheet = BG_ss_().getSheetByName(BAOGIA_CONFIG.MA_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: "error", message: "Danh mục trống." };
    const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // Cột B
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0] || "").trim() === String(maBaoGia || "").trim()) {
        sheet.deleteRow(i + 2);
        return { status: "success", message: "Đã xóa mã báo giá " + maBaoGia + " khỏi danh mục." };
      }
    }
    return { status: "error", message: "Không tìm thấy mã báo giá " + maBaoGia };
  } catch (e) { return { status: "error", message: e.toString() }; }
  finally { lock.releaseLock(); }
}

/* ---------- 5.2 Danh mục mã khối lượng (Ma_KL) ---------- */
function BG_getMaKLList() {
  try {
    const sheet = BG_ss_().getSheetByName(BAOGIA_CONFIG.MAKL_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: "success", data: [] };
    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const result = data
      .map(r => ({
        timestamp: (r[0] instanceof Date) ? Utilities.formatDate(r[0], "GMT+7", "dd/MM/yyyy HH:mm") : "",
        maKL: r[1], klMinKg: parseFloat(r[2]) || 0, klMaxKg: parseFloat(r[3]) || 0
      }))
      .filter(r => String(r.maKL || "").trim() !== "");
    return { status: "success", data: result };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

// fields: {klMinTan, klMaxTan} - nhập theo đơn vị TẤN cho thân thiện
function BG_addMaKL(fields) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(CONFIG.LOCK_TIMEOUT_MS); } catch (e) {
    return { status: "error", message: "Hệ thống đang bận, vui lòng thử lại." };
  }
  try {
    const klMinTan = parseFloat(fields.klMinTan);
    const klMaxTan = parseFloat(fields.klMaxTan);
    if (isNaN(klMinTan) || isNaN(klMaxTan) || klMaxTan <= klMinTan || klMinTan < 0) {
      return { status: "error", message: "Khoảng khối lượng không hợp lệ (Max phải lớn hơn Min, Min ≥ 0)." };
    }
    // Mã khối lượng ghép trực tiếp theo đơn vị TẤN (vd "0_45"), khớp đúng cách
    // BG_coreLogicProcessor_ tách chuỗi mã (split "_") để tính KL_MIN/KL_MAX khi
    // xác định hiệu lực báo giá — không được đổi quy ước này.
    const maKL = String(klMinTan) + "_" + String(klMaxTan);
    const sheet = BG_ss_().getSheetByName(BAOGIA_CONFIG.MAKL_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const existing = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      for (let row of existing) {
        if (String(row[0] || "").trim() === maKL) {
          return { status: "error", message: "Mã khối lượng " + maKL + " đã tồn tại." };
        }
      }
    }
    // Lưu KL_Min/KL_Max theo đúng quy ước hiện có trong sheet Ma_KL: đơn vị Kg
    sheet.getRange(lastRow + 1, 1, 1, 4).setValues([[new Date(), maKL, klMinTan * 1000, klMaxTan * 1000]]);
    return { status: "success", message: "Đã thêm mã khối lượng " + maKL, maKL: maKL };
  } catch (e) { return { status: "error", message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function BG_deleteMaKL(maKL) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(CONFIG.LOCK_TIMEOUT_MS); } catch (e) {
    return { status: "error", message: "Hệ thống đang bận, vui lòng thử lại." };
  }
  try {
    const sheet = BG_ss_().getSheetByName(BAOGIA_CONFIG.MAKL_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: "error", message: "Danh mục trống." };
    const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0] || "").trim() === String(maKL || "").trim()) {
        sheet.deleteRow(i + 2);
        return { status: "success", message: "Đã xóa mã khối lượng " + maKL };
      }
    }
    return { status: "error", message: "Không tìm thấy mã khối lượng " + maKL };
  } catch (e) { return { status: "error", message: e.toString() }; }
  finally { lock.releaseLock(); }
}

/* ---------- 5.3 Danh sách phiếu báo giá (QL_BaoGia) — phục vụ Sao chép ---------- */
function BG_getQuoteList() {
  try {
    const sheet = BG_ss_().getSheetByName(BAOGIA_CONFIG.QL_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: "success", data: [] };
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const result = data
      .map(r => ({
        ngayBaoGia: (r[0] instanceof Date) ? Utilities.formatDate(r[0], "GMT+7", "dd/MM/yyyy") : "",
        soBaoGia: r[1],
        hieuLuc: (r[3] instanceof Date) ? Utilities.formatDate(r[3], "GMT+7", "dd/MM/yyyy HH:mm") : "",
        idTam: r[4] || ""
      }))
      .filter(r => String(r.soBaoGia || "").trim() !== "");
    result.sort((a, b) => (a.soBaoGia < b.soBaoGia ? 1 : -1));
    return { status: "success", data: result };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

// Lấy các "nhóm giá" của 1 phiếu báo giá cũ để làm dữ liệu mồi cho chức năng SAO CHÉP
function BG_getQuoteDetail(soBaoGia) {
  try {
    const sheet = BG_ss_().getSheetByName(BAOGIA_CONFIG.SRC_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: "success", data: [] };
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    const groups = data
      .filter(r => String(r[7] || "").trim() === String(soBaoGia || "").trim())
      .map(r => ({
        maList: String(r[2] || "").split(",").map(s => s.trim()).filter(Boolean),
        klCode: String(r[3] || "").trim(),
        gia: parseFloat(r[4]) || 0
      }));
    return { status: "success", data: groups };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

/* ---------- 5.3B Sửa / Xóa báo giá (chỉ cho phép khi CHƯA có phiếu cân áp dụng VÀ CHƯA hết hiệu lực) ---------- */

// Đọc PhieuCan_DN 1 lần, gom theo Mã ĐG -> danh sách timestamp Ngày cân 1, dùng
// chung cho cả kiểm tra 1 dòng (BG_checkRowEditable_) lẫn tính hàng loạt cho cả
// bảng (BG_annotateApplied_) - tránh phải đọc lại PhieuCan_DN nhiều lần.
function BG_getPhieuCanByMaDG_() {
  const byMa = {};
  const ssHak = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheetPC = ssHak.getSheetByName(CONFIG.DATA_SHEET);
  const lastRowPC = sheetPC.getLastRow();
  if (lastRowPC > 1) {
    // Cột B(2)..Q(17): idx0=Ngày cân 1, idx15=Mã ĐG
    const dataPC = sheetPC.getRange(2, 2, lastRowPC - 1, 16).getValues();
    dataPC.forEach(pc => {
      const ngayCan1 = pc[0];
      const ma = String(pc[15] || "").trim();
      if (!ma || !(ngayCan1 instanceof Date) || isNaN(ngayCan1.getTime())) return;
      if (!byMa[ma]) byMa[ma] = [];
      byMa[ma].push(ngayCan1.getTime());
    });
  }
  return byMa;
}

// Kiểm tra 1 nhóm giá (idBgct, gồm các mã trong maList) có được phép Sửa/Xóa không.
// KHÔNG cho phép nếu: (a) bất kỳ mã nào trong nhóm đã "Hết hiệu lực" (bảo toàn lịch sử),
// HOẶC (b) bất kỳ mã nào trong nhóm đã có phiếu cân dùng đúng mã đó trong đúng
// khoảng thời gian hiệu lực tương ứng (bảo toàn tính đúng đắn của số liệu đã tính tiền).
// Đây là hàm KIỂM TRA CUỐI CÙNG ngay trước khi ghi (defense in depth) - luôn đọc
// dữ liệu MỚI NHẤT tại thời điểm gọi, không phụ thuộc dữ liệu đã tải sẵn ở client.
function BG_checkRowEditable_(idBgct, maList) {
  idBgct = String(idBgct || "").trim(); // FIX: chuẩn hóa để so khớp đúng dù caller có lỡ không trim
  const result = BG_coreLogicProcessor_(new Set());
  // Dùng finalRows (có Date object thật) thay vì displayData (chuỗi đã format) để so khớp chính xác
  const relevantRows = result.finalRows.filter(row => String(row[0] || "").trim() === idBgct && maList.indexOf(row[3]) !== -1);

  for (const row of relevantRows) {
    if (row[13] === "Hết hiệu lực") {
      return { editable: false, reason: "Mã \"" + row[3] + "\" đã HẾT HIỆU LỰC — không thể sửa/xóa để bảo toàn lịch sử báo giá." };
    }
  }

  if (relevantRows.length > 0) {
    const byMa = BG_getPhieuCanByMaDG_();
    for (const row of relevantRows) {
      const ma = row[3];
      const tuTS = row[1].getTime();
      const denTS = row[2].getTime();
      const list = byMa[ma] || [];
      const daApDung = list.some(ts => ts >= tuTS && ts <= denTS);
      if (daApDung) {
        return { editable: false, reason: "Mã \"" + ma + "\" ĐÃ CÓ PHIẾU CÂN ÁP DỤNG trong khoảng hiệu lực này — không thể sửa/xóa để không làm sai lệch số liệu đã tính tiền." };
      }
    }
  }

  return { editable: true, reason: "" };
}

// Tính hàng loạt "editable/reason/daApDung" cho TOÀN BỘ displayData cùng lúc
// (chỉ đọc PhieuCan_DN 1 lần), dùng để hiển thị NGAY trên bảng Hiệu lực báo giá
// (mờ/khóa nút Sửa-Xóa cho các dòng không đủ điều kiện) thay vì phải bấm thử mới biết.
function BG_annotateApplied_(finalRows, displayData) {
  const byMa = BG_getPhieuCanByMaDG_();
  return displayData.map((d, idx) => {
    const finalRow = finalRows[idx]; // finalRows và displayData luôn cùng thứ tự (push song song trong BG_coreLogicProcessor_)
    const tuTS = finalRow[1].getTime();
    const denTS = finalRow[2].getTime();
    const list = byMa[d.ma] || [];
    const daApDung = list.some(ts => ts >= tuTS && ts <= denTS);
    const hetHieuLuc = d.trangThai === "Hết hiệu lực";
    let reason = "";
    if (hetHieuLuc) reason = "Đã hết hiệu lực, không thể sửa/xóa.";
    else if (daApDung) reason = "Đã có phiếu cân áp dụng, không thể sửa/xóa.";
    const out = {};
    for (const k in d) out[k] = d[k];
    out.editable = !hetHieuLuc && !daApDung;
    out.daApDung = daApDung;
    out.reason = reason;
    return out;
  });
}

// Lấy chi tiết 1 dòng báo giá theo ID_BGCT (mã băm), kèm trạng thái có được sửa/xóa hay không
function BG_getBaogiaRowByHash(idBgct) {
  try {
    const sheet = BG_ss_().getSheetByName(BAOGIA_CONFIG.SRC_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: "error", message: "Không có dữ liệu báo giá." };
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    let found = null;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][6] || "").trim() === String(idBgct || "").trim()) { found = data[i]; break; }
    }
    if (!found) return { status: "error", message: "Không tìm thấy báo giá " + idBgct };

    const maList = String(found[2] || "").split(",").map(s => s.trim()).filter(Boolean);
    const klCode = String(found[3] || "").trim();
    const gia = parseFloat(found[4]) || 0;
    const hieuLuc = found[1];
    const editCheck = BG_checkRowEditable_(idBgct, maList);

    return {
      status: "success",
      idBgct: idBgct,
      maList: maList,
      klCode: klCode,
      gia: gia,
      hieuLuc: (hieuLuc instanceof Date) ? Utilities.formatDate(hieuLuc, "GMT+7", "yyyy-MM-dd'T'HH:mm") : "",
      soBaoGia: String(found[7] || "").trim(),
      editable: editCheck.editable,
      reason: editCheck.reason
    };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

// payload = { idBgct, maList:[...], klCode, gia, hieuLuc:'yyyy-MM-ddTHH:mm' }
function BG_updateBaogiaRow(payload) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(CONFIG.LOCK_TIMEOUT_MS); } catch (e) {
    return { status: "error", message: "Hệ thống đang bận, vui lòng thử lại." };
  }
  try {
    const idBgct = String(payload.idBgct || "").trim();
    const maList = (payload.maList || []).map(s => String(s).trim()).filter(Boolean);
    const klCode = String(payload.klCode || "").trim();
    const gia = parseFloat(payload.gia);
    if (!idBgct) return { status: "error", message: "Thiếu mã định danh dòng báo giá." };
    if (maList.length === 0) return { status: "error", message: "Vui lòng chọn ít nhất 1 mã báo giá." };
    if (!klCode) return { status: "error", message: "Vui lòng chọn mã khối lượng." };
    if (isNaN(gia) || gia <= 0) return { status: "error", message: "Đơn giá phải lớn hơn 0." };
    if (!payload.hieuLuc) return { status: "error", message: "Vui lòng chọn thời điểm hiệu lực." };
    const hieuLucDate = new Date(payload.hieuLuc + ":00+07:00");
    if (isNaN(hieuLucDate.getTime())) return { status: "error", message: "Thời điểm hiệu lực không hợp lệ." };

    const sheet = BG_ss_().getSheetByName(BAOGIA_CONFIG.SRC_SHEET);
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    let rowIndex = -1; let oldMaList = [];
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][6] || "").trim() === idBgct) { rowIndex = i; oldMaList = String(data[i][2] || "").split(",").map(s => s.trim()).filter(Boolean); break; }
    }
    if (rowIndex === -1) return { status: "error", message: "Không tìm thấy báo giá " + idBgct + " (có thể đã bị người khác xóa)." };

    // Kiểm tra lại NGAY TRƯỚC KHI GHI (phòng trường hợp có phiếu cân mới phát sinh
    // giữa lúc mở form sửa và lúc bấm Lưu) — kiểm tra cả danh sách mã CŨ lẫn MỚI
    const checkOld = BG_checkRowEditable_(idBgct, oldMaList);
    if (!checkOld.editable) return { status: "error", message: "Không thể sửa: " + checkOld.reason };

    const sheetRow = rowIndex + 2;
    sheet.getRange(sheetRow, 2).setValue(hieuLucDate);       // Cột B - Thời điểm hiệu lực
    sheet.getRange(sheetRow, 3).setValue(maList.join(" , ")); // Cột C - Mã đơn giá
    sheet.getRange(sheetRow, 4).setValue(klCode);             // Cột D - Khối lượng_Tấn
    sheet.getRange(sheetRow, 5).setValue(gia);                // Cột E - Đơn giá

    return { status: "success", message: "Đã cập nhật báo giá " + idBgct + "." };
  } catch (e) { return { status: "error", message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function BG_deleteBaogiaRow(idBgct) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(CONFIG.LOCK_TIMEOUT_MS); } catch (e) {
    return { status: "error", message: "Hệ thống đang bận, vui lòng thử lại." };
  }
  try {
    idBgct = String(idBgct || "").trim();
    const sheet = BG_ss_().getSheetByName(BAOGIA_CONFIG.SRC_SHEET);
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    let rowIndex = -1; let maList = [];
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][6] || "").trim() === idBgct) { rowIndex = i; maList = String(data[i][2] || "").split(",").map(s => s.trim()).filter(Boolean); break; }
    }
    if (rowIndex === -1) return { status: "error", message: "Không tìm thấy báo giá " + idBgct + " (có thể đã bị xóa trước đó)." };

    const editCheck = BG_checkRowEditable_(idBgct, maList);
    if (!editCheck.editable) return { status: "error", message: "Không thể xóa: " + editCheck.reason };

    sheet.deleteRow(rowIndex + 2);
    return { status: "success", message: "Đã xóa báo giá " + idBgct + "." };
  } catch (e) { return { status: "error", message: e.toString() }; }
  finally { lock.releaseLock(); }
}

/* ---------- 5.4 Tạo phiếu báo giá mới (Nhập mới / Sao chép / Tách nhóm) ---------- */
// Ghi chú: "Tách báo giá" xử lý hoàn toàn ở phía giao diện (client tách 1 nhóm
// nhiều mã thành nhiều nhóm 1 mã trước khi gửi) - không cần hàm backend riêng.
// payload = { ngayBaoGia:'yyyy-MM-dd', hieuLuc:'yyyy-MM-ddTHH:mm', idTam:'', groups:[{maList:[...], klCode, gia}] }
function BG_createQuote(payload) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(CONFIG.LOCK_TIMEOUT_MS); } catch (e) {
    return { status: "error", message: "Hệ thống đang bận xử lý một yêu cầu khác, vui lòng thử lại sau ít giây." };
  }
  try {
    if (!payload || !payload.ngayBaoGia) return { status: "error", message: "Vui lòng chọn ngày báo giá." };
    if (!payload.hieuLuc) return { status: "error", message: "Vui lòng chọn thời điểm hiệu lực." };
    const groups = (payload.groups || []).filter(g => g.maList && g.maList.length > 0 && g.klCode && parseFloat(g.gia) > 0);
    if (groups.length === 0) return { status: "error", message: "Vui lòng thêm ít nhất 1 nhóm giá hợp lệ (có mã báo giá, mã khối lượng và đơn giá > 0)." };

    const ngayBaoGiaDate = new Date(payload.ngayBaoGia + "T00:00:00+07:00");
    const hieuLucDate = new Date(payload.hieuLuc + ":00+07:00");
    if (isNaN(ngayBaoGiaDate.getTime()) || isNaN(hieuLucDate.getTime())) {
      return { status: "error", message: "Ngày báo giá hoặc thời điểm hiệu lực không hợp lệ." };
    }

    const ss = BG_ss_();
    const qlSheet = ss.getSheetByName(BAOGIA_CONFIG.QL_SHEET);
    const srcSheet = ss.getSheetByName(BAOGIA_CONFIG.SRC_SHEET);

    // Sinh Số báo giá dạng YYYYMMDD-NNN, NNN tăng dần riêng theo từng ngày
    const datePrefix = Utilities.formatDate(ngayBaoGiaDate, "GMT+7", "yyyyMMdd");
    const qlLastRow = qlSheet.getLastRow();
    let maxSeq = 0;
    if (qlLastRow > 1) {
      const existingSo = qlSheet.getRange(2, 2, qlLastRow - 1, 1).getValues();
      existingSo.forEach(row => {
        const v = String(row[0] || "").trim();
        if (v.indexOf(datePrefix + "-") === 0) {
          const seq = parseInt(v.substring(datePrefix.length + 1), 10);
          if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
        }
      });
    }
    const soBaoGiaMoi = datePrefix + "-" + String(maxSeq + 1).padStart(3, "0");
    const now = new Date();
    let userEmail = "";
    try { userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || ""; } catch (e) { /* bỏ qua nếu không lấy được email */ }

    qlSheet.getRange(qlLastRow + 1, 1, 1, 5).setValues([[ngayBaoGiaDate, soBaoGiaMoi, now, hieuLucDate, payload.idTam || ""]]);

    const rowsToAppend = groups.map(g => {
      const hash = Utilities.getUuid().replace(/-/g, "").substring(0, 8);
      return [now, hieuLucDate, g.maList.join(" , "), g.klCode, parseFloat(g.gia), userEmail, hash, soBaoGiaMoi];
    });
    srcSheet.getRange(srcSheet.getLastRow() + 1, 1, rowsToAppend.length, 8).setValues(rowsToAppend);

    return { status: "success", message: "Đã lưu báo giá " + soBaoGiaMoi + " với " + groups.length + " nhóm giá.", soBaoGia: soBaoGiaMoi };
  } catch (e) { return { status: "error", message: e.toString() }; }
  finally { lock.releaseLock(); }
}

/* ---------- 5.5 Xử lý hiệu lực & Xem dữ liệu (port nguyên vẹn từ hệ thống Báo giá gốc) ---------- */
function BG_updateHieuLuc() {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(CONFIG.LOCK_TIMEOUT_MS); } catch (e) {
    return { status: "error", message: "Hệ thống đang bận, vui lòng thử lại." };
  }
  try {
    const ss = BG_ss_();
    const dst = ss.getSheetByName(BAOGIA_CONFIG.DST_SHEET);
    const dstData = dst.getDataRange().getValues();
    const blockedIDs = new Set();
    if (dstData.length > 1) {
      for (let i = 1; i < dstData.length; i++) {
        if (dstData[i][13] === "Hết hiệu lực" && dstData[i][0]) blockedIDs.add(dstData[i][0].toString().trim());
      }
    }
    const result = BG_coreLogicProcessor_(blockedIDs);
    const activeRows = result.finalRows.filter(row => row[13] === "Còn hiệu lực");
    if (dst.getLastRow() > 1) dst.getRange(2, 1, dst.getLastRow() - 1, 14).clearContent();
    if (activeRows.length > 0) dst.getRange(2, 1, activeRows.length, 14).setValues(activeRows);
    // Gắn sẵn editable/reason cho TOÀN BỘ trước khi lọc, để chỉ số giữa finalRows và
    // displayData luôn khớp nhau (BG_annotateApplied_ dựa vào cùng vị trí index).
    const annotated = BG_annotateApplied_(result.finalRows, result.displayData);
    return { status: "success", data: annotated.filter(r => r.trangThai === "Còn hiệu lực"), viewType: "FINAL" };
  } catch (e) { return { status: "error", message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function BG_showAllData() {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(CONFIG.LOCK_TIMEOUT_MS); } catch (e) {
    return { status: "error", message: "Hệ thống đang bận, vui lòng thử lại." };
  }
  try {
    const ss = BG_ss_();
    const save = ss.getSheetByName(BAOGIA_CONFIG.SAVE_SHEET);
    const result = BG_coreLogicProcessor_(new Set());
    if (save.getLastRow() > 1) save.getRange(2, 1, save.getLastRow() - 1, 14).clearContent();
    if (result.finalRows.length > 0) save.getRange(2, 1, result.finalRows.length, 14).setValues(result.finalRows);
    const annotated = BG_annotateApplied_(result.finalRows, result.displayData);
    return { status: "success", data: annotated, viewType: "SAVE" };
  } catch (e) { return { status: "error", message: e.toString() }; }
  finally { lock.releaseLock(); }
}

// Giữ NGUYÊN VẸN logic gốc (chỉ đổi tên hàm + trỏ về BAOGIA_CONFIG) để không
// làm thay đổi cách xác định "Còn/Chưa/Hết hiệu lực" đang vận hành thực tế.
function BG_coreLogicProcessor_(blockedIDs) {
  const ss = BG_ss_();
  const srcData = ss.getSheetByName(BAOGIA_CONFIG.SRC_SHEET).getDataRange().getValues();
  const maData = ss.getSheetByName(BAOGIA_CONFIG.MA_SHEET).getDataRange().getValues();
  const now = new Date();
  const currentTS = now.getTime();
  const maLookup = new Map();

  for (let i = 1; i < maData.length; i++) {
    if (maData[i][1]) maLookup.set(maData[i][1].toString().trim(), maData[i][5]);
  }

  const rawRecords = [];
  for (let i = 1; i < srcData.length; i++) {
    const id_bg = srcData[i][0]?.toString().trim() || "";
    const id_goc = srcData[i][6]?.toString().trim() || "";
    if (blockedIDs.has(id_bg) || blockedIDs.has(id_goc)) continue;

    if (id_goc && srcData[i][1] instanceof Date && srcData[i][2]) {
      srcData[i][2].toString().split(",").forEach(m => {
        let kl = (srcData[i][3] || "0_999999").toString().split("_");
        rawRecords.push({
          id: id_goc, tuNgay: srcData[i][1], ma: m.trim(),
          pE: kl[0] || "0", pF: kl[1] || "999999", gia: srcData[i][4]
        });
      });
    }
  }

  const groups = {};
  rawRecords.forEach(r => {
    let groupKey = r.ma + "_MIN" + r.pE + "_MAX" + r.pF;
    if (!groups[groupKey]) groups[groupKey] = [];
    groups[groupKey].push(r);
  });

  const finalRows = [];
  const displayData = [];

  for (let key in groups) {
    const sorted = groups[key].sort((a, b) => a.tuNgay.getTime() - b.tuNgay.getTime());
    for (let j = 0; j < sorted.length; j++) {
      let tN = sorted[j].tuNgay.getTime();
      let dN = (j < sorted.length - 1) ? Math.max(tN, sorted[j + 1].tuNgay.getTime() - 1000) : new Date("2050-12-31").getTime();
      let status = (currentTS >= tN && currentTS <= dN) ? "Còn hiệu lực" : (currentTS < tN ? "Chưa đến hạn" : "Hết hiệu lực");

      const row = [
        sorted[j].id, sorted[j].tuNgay, new Date(dN), sorted[j].ma, sorted[j].pE, sorted[j].pF,
        sorted[j].gia, "", now, "", "", sorted[j].id + "_" + sorted[j].ma,
        maLookup.get(sorted[j].ma) || "", status
      ];

      finalRows.push(row);
      displayData.push({
        id: row[0], ma: row[3], gia: row[6], dienGiai: row[12],
        klMin: row[4], klMax: row[5],
        tuNgay: Utilities.formatDate(row[1], "GMT+7", "dd/MM/yyyy"),
        denNgay: Utilities.formatDate(row[2], "GMT+7", "dd/MM/yyyy"),
        ngayTS: tN, trangThai: status
      });
    }
  }
  return { finalRows, displayData };
}

function BG_exportFileSmart(dateStr, currentView, filteredIds) {
  try {
    const ss = BG_ss_();
    const dataSheet = ss.getSheetByName(currentView === "FINAL" ? BAOGIA_CONFIG.DST_SHEET : BAOGIA_CONFIG.SAVE_SHEET);
    const values = dataSheet.getDataRange().getValues().slice(1);
    const filterTS = new Date(dateStr + "T23:59:59+07:00").getTime();

    const filteredRaw = values.filter(r => {
      const rDate = new Date(r[1]).getTime();
      return rDate <= filterTS && filteredIds.includes(r[0].toString());
    });

    if (filteredRaw.length === 0) return { status: "error", message: "Không có dữ liệu phù hợp!" };

    const newSS = SpreadsheetApp.create("Bao_Gia_HAK_" + dateStr);
    const sheet = newSS.getSheets()[0];

    sheet.getRange("A1").setValue(COMPANY_NAME).setFontWeight("bold");
    sheet.getRange("A3:G3").merge().setValue("BẢNG BÁO GIÁ").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("A4:G4").merge().setValue("Ngày xuất: " + dateStr).setHorizontalAlignment("center").setFontStyle("italic");

    const headers = ["STT", "MÃ", "DIỄN GIẢI", "ĐƠN GIÁ", "HIỆU LỰC TỪ", "HIỆU LỰC ĐẾN", "STATUS"];
    sheet.getRange(6, 1, 1, 7).setValues([headers]).setFontWeight("bold").setBackground("#B7B7B7").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);

    const exportData = filteredRaw.map((r, index) => [index + 1, r[3], r[12], Number(r[6]) || 0, r[1], r[2], r[13]]);
    sheet.getRange(7, 1, exportData.length, 7).setValues(exportData).setBorder(true, true, true, true, true, true);

    sheet.getRange(7, 4, exportData.length, 1).setNumberFormat("#,##0");
    sheet.getRange(7, 5, exportData.length, 2).setNumberFormat("dd/MM/yyyy");
    sheet.getRange(7, 3, exportData.length, 1).setWrap(true);
    sheet.setColumnWidth(2, 100); sheet.setColumnWidth(3, 300); sheet.setColumnWidth(4, 110);

    SpreadsheetApp.flush();
    const folder = DriveApp.getFolderById(BAOGIA_CONFIG.BACKUP_FOLDER_ID);
    folder.addFile(DriveApp.getFileById(newSS.getId()));
    DriveApp.getRootFolder().removeFile(DriveApp.getFileById(newSS.getId()));

    return { status: "success", url: "https://docs.google.com/spreadsheets/d/" + newSS.getId() + "/export?format=xlsx" };
  } catch (e) { return { status: "error", message: e.toString() }; }
}
