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
            // dùng các chuỗi này, nên KHÔNG bị ảnh hưởng bởi lỗi định dạng ở đây.
            // FIX #6: đổi dd/MM/yyyy (kiểu Việt Nam) thay vì MM/dd/yyyy (kiểu Mỹ) — bản
            // cũ khiến ngày 05/01/2026 (5 tháng 1) hiển thị thành "01/05/2026", làm
            // người dùng tưởng nhầm là ngày 1 tháng 5 (bị đảo ngày/tháng), dù dữ liệu
            // gốc ghi vào PhieuCan_DN vẫn luôn đúng.
            ngayCan1: dateC ? Utilities.formatDate(dateC, "GMT+7", "dd/MM/yyyy") : "Lỗi định dạng ngày",
            gioCan1: dateC ? Utilities.formatDate(dateC, "GMT+7", "HH:mm:ss") : "",
            ngayCan2: dateD ? Utilities.formatDate(dateD, "GMT+7", "dd/MM/yyyy") : "Lỗi định dạng ngày",
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
      // FIX #6 (tiếp): ép định dạng Text (@) cho cột Số Chứng Từ..Giờ Cân 2 (B..H)
      // TRƯỚC KHI ghi giá trị — bắt buộc làm trước setValues, nếu không Google
      // Sheets có thể tự động "đoán" và chuyển chuỗi ngày dd/MM/yyyy thành kiểu
      // Date thật theo Locale mặc định của chính file Draft này (thường lấy theo
      // Locale tài khoản Google, có thể là US), khiến ngày lại bị đảo lần nữa dù
      // chuỗi gốc (từ FIX #6 ở trên) đã đúng.
      draftSheet.getRange(2, 2, previewRows.length, 7).setNumberFormat("@"); // B..H (Số Chứng Từ..Giờ Cân 2)
      draftSheet.getRange(2, 1, draftValues.length, draftHeaders.length).setValues(draftValues);
      draftSheet.getRange(2, 9, draftValues.length, 3).setNumberFormat("#,##0");
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
    // Nhánh này chỉ còn phục vụ dữ liệu CŨ đã tồn tại trước khi áp dụng FIX #1/#3
    // (được ghi dưới dạng chuỗi). Dữ liệu mới ghi từ nay là Date object thật
    // nên sẽ luôn rơi vào nhánh "v instanceof Date" ở trên, không đi qua đây.
    // FIX #10: đồng bộ với toDateObj - mặc định DD/MM/YYYY (đã chứng minh đúng
    // bằng số học), tự sửa nếu phần "tháng" > 12 mà phần "ngày" ≤ 12.
    let p = v.includes("-") ? v.split("-") : v.split("/");
    if (p.length === 3) {
      if (p[0].length === 4) return new Date(p[0], p[1] - 1, p[2]); // yyyy-MM-dd, không mơ hồ
      let ngay = parseInt(p[0], 10);
      let thang = parseInt(p[1], 10);
      if (thang > 12 && ngay <= 12) { const tmp = ngay; ngay = thang; thang = tmp; }
      return new Date(p[2], thang - 1, ngay);
    }
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
    // FIX #10 (THAY THẾ FIX #5 - FIX #5 ĐÃ SAI HƯỚNG): Đã có bằng chứng số học rõ
    // ràng: chuỗi "25/07/2026" (ngày thật 25/7) bị FIX #5 (giả định MM/DD/YYYY)
    // đọc nhầm "25" thành THÁNG -> tháng tự tràn thành tháng 1/2028, ra đúng lỗi
    // "07/01/2028" người dùng gặp phải. Vì "25" không thể là tháng, xác nhận file
    // gốc dùng DD/MM/YYYY (kiểu Việt Nam) chứ không phải MM/DD/YYYY.
    // Để tránh phụ thuộc cứng vào 1 giả định cố định (đã sai 1 lần), nay dùng cơ
    // chế TỰ SỬA: mặc định đọc theo DD/MM/YYYY (parts[0]=ngày, parts[1]=tháng),
    // nhưng nếu phần "tháng" > 12 (chắc chắn vô lý) trong khi phần "ngày" ≤ 12,
    // thì tự hoán đổi lại thành MM/DD/YYYY cho đúng dữ liệu. Cách này không bao
    // giờ tạo ra tháng > 12 (không còn hiện tượng tràn năm/tháng như lỗi vừa gặp).
    const parts = trimmed.match(/(\d+)/g);
    if (parts && parts.length >= 3) {
      let ngay = parseInt(parts[0], 10);
      let thang = parseInt(parts[1], 10);
      if (thang > 12 && ngay <= 12) { const tmp = ngay; ngay = thang; thang = tmp; }
      return new Date(parts[2], thang - 1, ngay, parts[3] || 0, parts[4] || 0, parts[5] || 0);
    }
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

/* ---------- File mẫu tải về cho 2 loại Import (không đụng dữ liệu thật) ---------- */

// Mẫu Phiếu cân NHẬP (PhieuCan_DN) - đúng theo tiêu đề đã XÁC NHẬN từ file thật
// "BÁO CÁO XE ĐÃ CÂN" của phần mềm trạm cân (kiểu cân = Nhập).
function taoFileMauPhieuCan() {
  try {
    const headers = ["STT", "Số phiếu", "Ngày giờ cân 1", "Ngày giờ cân 2", "Biển số 1", "Cân lần 1", "Cân lần 2", "KL Hàng (KG)", "quy cách", "Khách hàng", "Người cân 1", "Đại lý", "Nguồn gốc", "ĐL"];
    const vd = [1, 7107, "25/07/2026 02:06:28", "25/07/2026 02:59:58", "92A-54957", 27020, 9030, 17990, "Tọa độ 15.762805, 108.150135, Xã Duy Xuyên, TP Đà Nẵng", "NGUYỄN VĂN HOÀNG", "Phạm Đình Giao", "NGÔ DUY THỨC", "ĐL"];
    const ghiChu = "Lưu ý: Cột 'Ngày giờ cân 1/2' PHẢI theo định dạng DD/MM/YYYY HH:mm:ss (ví dụ: 25/07/2026 02:06:28). Xóa dòng ví dụ này và điền dữ liệu thật của bạn từ dòng 3 trở xuống, giữ nguyên dòng tiêu đề (dòng 2).";
    return taoFileMau_("Mau_Import_PhieuCan_Nhap", headers, vd, ghiChu);
  } catch (e) { return { status: "error", message: e.toString() }; }
}

// Mẫu phiếu cân XUẤT HÀNG (NL_PC_XH) - GHÉP từ tiêu đề THẬT của NL_PC_XH cho các
// cột lõi (chắc chắn đúng), cộng thêm các cột đặc thù xuất hàng theo suy đoán
// tốt nhất (Đơn vị vận chuyển/Tên tài xế) vì CHƯA có file mẫu thật kiểu "Xuất"
// để đối chiếu — nếu tên cột thực tế trên trạm cân khác đi, chỉ cần đổi lại
// đúng tên tiêu đề trong file bạn tải lên cho khớp là hệ thống vẫn đọc được
// (import tìm cột theo TÊN, không theo vị trí cố định).
function taoFileMauXuatHang() {
  try {
    const headers = ["STT", "Số phiếu", "Ngày giờ cân 1", "Ngày giờ cân 2", "Biển số 1", "Cân lần 1", "Cân lần 2", "KL Hàng (KG)", "Đơn vị vận chuyển", "Tên tài xế", "NGƯỜI CÂN", "Kho xuất", "Kho nhập"];
    const vd = [1, "01", "22/01/2025 07:36:52", "22/01/2025 08:02:32", "92C-08727", 14160, 36080, 21920, "Công Ty TNHH Dịch Vụ Và Vận Tải Hùng Hoàng Hoa", "Mr Tâm", "Trần Thị Phương", "Kho Tiên Sa", "Kho Xuất Bán"];
    const ghiChu = "⚠️ Mẫu này là suy đoán tốt nhất dựa trên cấu trúc sheet NL_PC_XH thật, CHƯA có file mẫu thật từ trạm cân kiểu 'Xuất' để đối chiếu. Cột 'Ngày giờ cân 1/2' PHẢI theo định dạng DD/MM/YYYY HH:mm:ss. Cột 'Kho xuất'/'Kho nhập' là TÙY CHỌN - nếu trạm cân không xuất ra 2 cột này, có thể để trống trong file và chọn giá trị áp dụng chung cho cả lô ngay trên giao diện Import. Hệ thống tìm cột theo ĐÚNG TÊN TIÊU ĐỀ (không theo vị trí), nên nếu tên cột thật trên trạm cân khác đi, chỉ cần sửa lại đúng tên cho khớp là vẫn import được.";
    return taoFileMau_("Mau_Import_XuatHang", headers, vd, ghiChu);
  } catch (e) { return { status: "error", message: e.toString() }; }
}

// Helper dùng chung: tạo 1 Google Sheet mẫu (dòng 1 = ghi chú, dòng 2 = tiêu đề,
// dòng 3 = ví dụ), lưu vào FOLDER_DONE, trả về link tải xuống dạng .xlsx.
function taoFileMau_(tenFile, headers, dongViDu, ghiChu) {
  const tempSS = SpreadsheetApp.create(tenFile + "_" + Utilities.formatDate(new Date(), "GMT+7", "ddMM_HHmm"));
  const sh = tempSS.getSheets()[0];
  sh.getRange(1, 1, 1, headers.length).merge().setValue(ghiChu).setWrap(true).setFontStyle("italic").setFontColor("#B3261E");
  sh.getRange(2, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#1B4332").setFontColor("#FFFFFF");
  sh.getRange(3, 1, 1, dongViDu.length).setValues([dongViDu]);
  sh.getRange(3, 1, 1, dongViDu.length).setNumberFormat("@"); // ép Text để không bị Sheets tự đoán lại định dạng ngày
  sh.autoResizeColumns(1, headers.length);
  sh.setRowHeight(1, 50);

  const tempFile = DriveApp.getFileById(tempSS.getId());
  DriveApp.getFolderById(CONFIG.FOLDER_DONE).addFile(tempFile);
  DriveApp.getRootFolder().removeFile(tempFile);

  return { status: "success", url: "https://docs.google.com/spreadsheets/d/" + tempSS.getId() + "/export?format=xlsx" };
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

/* ---------- 5.3C Xóa TOÀN BỘ 1 phiếu báo giá (QL_BaoGia + mọi nhóm giá liên quan) ---------- */

// Kiểm tra 1 PHIẾU BÁO GIÁ (soBaoGia) có được xóa TOÀN BỘ hay không: chỉ cho phép
// khi TẤT CẢ nhóm giá (dòng Baogia_DN) thuộc phiếu này đều editable (chưa hết
// hiệu lực và chưa có phiếu cân áp dụng) - dùng lại đúng quy tắc như xóa từng dòng.
function BG_checkQuoteDeletable_(soBaoGia) {
  soBaoGia = String(soBaoGia || "").trim();
  const sheet = BG_ss_().getSheetByName(BAOGIA_CONFIG.SRC_SHEET);
  const lastRow = sheet.getLastRow();
  const rows = [];
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    data.forEach(r => {
      if (String(r[7] || "").trim() === soBaoGia) {
        rows.push({ idBgct: String(r[6] || "").trim(), maList: String(r[2] || "").split(",").map(s => s.trim()).filter(Boolean) });
      }
    });
  }
  if (rows.length === 0) return { deletable: false, reason: "Không tìm thấy nhóm giá nào thuộc báo giá này (có thể đã bị xóa).", rows: rows };

  for (const row of rows) {
    const check = BG_checkRowEditable_(row.idBgct, row.maList);
    if (!check.editable) {
      return { deletable: false, reason: "Nhóm mã \"" + row.maList.join(", ") + "\": " + check.reason, rows: rows };
    }
  }
  return { deletable: true, reason: "", rows: rows };
}

// Trả danh sách TOÀN BỘ phiếu báo giá (đầu phiếu QL_BaoGia) kèm số nhóm giá và
// trạng thái có được xóa cả phiếu hay không - phục vụ bảng "Danh sách báo giá đã lập".
function BG_getQuoteListWithStatus() {
  try {
    const sheet = BG_ss_().getSheetByName(BAOGIA_CONFIG.QL_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: "success", data: [] };
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

    const srcSheet = BG_ss_().getSheetByName(BAOGIA_CONFIG.SRC_SHEET);
    const lastRowSrc = srcSheet.getLastRow();
    const srcData = lastRowSrc > 1 ? srcSheet.getRange(2, 1, lastRowSrc - 1, 8).getValues() : [];

    // Gom nhóm giá theo Số báo giá 1 lần, tránh quét lại Baogia_DN cho từng phiếu
    const bySoBaoGia = {};
    srcData.forEach(r => {
      const so = String(r[7] || "").trim();
      if (!so) return;
      if (!bySoBaoGia[so]) bySoBaoGia[so] = [];
      bySoBaoGia[so].push({ idBgct: String(r[6] || "").trim(), maList: String(r[2] || "").split(",").map(s => s.trim()).filter(Boolean) });
    });

    // Tính sẵn trạng thái hiệu lực + đã áp dụng cho TẤT CẢ (id, mã) chỉ 1 lần
    const byMa = BG_getPhieuCanByMaDG_();
    const coreResult = BG_coreLogicProcessor_(new Set());
    const statusMap = {}; // key = idBgct + "|" + ma -> {trangThai, tuTS, denTS}
    coreResult.finalRows.forEach(row => {
      statusMap[row[0] + "|" + row[3]] = { trangThai: row[13], tuTS: row[1].getTime(), denTS: row[2].getTime() };
    });

    const result = data.map(r => {
      const soBaoGia = String(r[1] || "").trim();
      const groups = bySoBaoGia[soBaoGia] || [];
      let deletable = groups.length > 0;
      let reason = groups.length === 0 ? "Không tìm thấy nhóm giá thuộc phiếu này." : "";
      outer:
      for (const g of groups) {
        for (const ma of g.maList) {
          const st = statusMap[g.idBgct + "|" + ma];
          if (!st) continue;
          if (st.trangThai === "Hết hiệu lực") { deletable = false; reason = "Mã \"" + ma + "\" đã hết hiệu lực."; break outer; }
          const list = byMa[ma] || [];
          const daApDung = list.some(ts => ts >= st.tuTS && ts <= st.denTS);
          if (daApDung) { deletable = false; reason = "Mã \"" + ma + "\" đã có phiếu cân áp dụng."; break outer; }
        }
      }
      return {
        soBaoGia: soBaoGia,
        ngayBaoGia: (r[0] instanceof Date) ? Utilities.formatDate(r[0], "GMT+7", "dd/MM/yyyy") : "",
        hieuLuc: (r[3] instanceof Date) ? Utilities.formatDate(r[3], "GMT+7", "dd/MM/yyyy HH:mm") : "",
        idTam: r[4] || "",
        soNhom: groups.length,
        deletable: deletable,
        reason: reason
      };
    }).filter(r => r.soBaoGia !== "");

    result.sort((a, b) => (a.soBaoGia < b.soBaoGia ? 1 : -1));
    return { status: "success", data: result };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

// Xóa TOÀN BỘ 1 phiếu báo giá: xóa hết các dòng nhóm giá thuộc phiếu trong
// Baogia_DN, VÀ xóa luôn dòng đầu phiếu trong QL_BaoGia. Chỉ cho phép khi
// TẤT CẢ nhóm giá thuộc phiếu đều editable (BG_checkQuoteDeletable_) - kiểm
// tra lại ngay trước khi xóa (defense in depth), không tin dữ liệu đã tải sẵn ở client.
function BG_deleteQuote(soBaoGia) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(CONFIG.LOCK_TIMEOUT_MS); } catch (e) {
    return { status: "error", message: "Hệ thống đang bận, vui lòng thử lại." };
  }
  try {
    soBaoGia = String(soBaoGia || "").trim();
    if (!soBaoGia) return { status: "error", message: "Thiếu số báo giá cần xóa." };

    const check = BG_checkQuoteDeletable_(soBaoGia);
    if (!check.deletable) return { status: "error", message: "Không thể xóa báo giá " + soBaoGia + ": " + check.reason };

    const srcSheet = BG_ss_().getSheetByName(BAOGIA_CONFIG.SRC_SHEET);
    const lastRowSrc = srcSheet.getLastRow();
    let soDongXoa = 0;
    if (lastRowSrc > 1) {
      const data = srcSheet.getRange(2, 1, lastRowSrc - 1, 8).getValues();
      // Xóa từ DƯỚI LÊN để không bị lệch chỉ số hàng khi deleteRow nhiều lần liên tiếp
      for (let i = data.length - 1; i >= 0; i--) {
        if (String(data[i][7] || "").trim() === soBaoGia) {
          srcSheet.deleteRow(i + 2);
          soDongXoa++;
        }
      }
    }

    const qlSheet = BG_ss_().getSheetByName(BAOGIA_CONFIG.QL_SHEET);
    const lastRowQL = qlSheet.getLastRow();
    if (lastRowQL > 1) {
      const dataQL = qlSheet.getRange(2, 1, lastRowQL - 1, 5).getValues();
      for (let i = dataQL.length - 1; i >= 0; i--) {
        if (String(dataQL[i][1] || "").trim() === soBaoGia) {
          qlSheet.deleteRow(i + 2);
        }
      }
    }

    return { status: "success", message: "Đã xóa toàn bộ báo giá " + soBaoGia + " (" + soDongXoa + " nhóm giá)." };
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

/*********************************************************
 * PHẦN 6: QUẢN LÝ TỒN KHO DĂM GỖ (MENU 4)
 * - Spreadsheet RIÊNG: KHODAM_CONFIG.SPREADSHEET_ID (xem Config.gs) - sheet
 *   ĐANG CHẠY THẬT, không dùng chung với PhieuCan_DN hay Báo giá.
 * - Nguồn gốc: sáp nhập từ hệ thống "Hệ Thống Quản Lý Kho Dăm - Hòa Nhơn"
 *   (project Apps Script riêng trước đây), giữ NGUYÊN VẸN toàn bộ logic
 *   nghiệp vụ gốc (Nhập/Xuất kho, Kỳ Vét Bãi, Độ khô, Tồn kho, Đối soát
 *   hoàn thành đơn hàng xuất bán) để không thay đổi hành vi đã được kiểm
 *   chứng bằng dữ liệu thật. Các thay đổi so với bản gốc:
 *   1) SpreadsheetApp.getActiveSpreadsheet() -> openById(KHODAM_CONFIG.SPREADSHEET_ID)
 *      (bắt buộc vì web app hợp nhất không còn "bound" trực tiếp vào sheet này nữa)
 *   2) Tên sheet hard-code -> tham chiếu KHODAM_CONFIG.SHEET_xxx (Config.gs)
 *   3) Timeout khóa 10s -> dùng chung CONFIG.LOCK_TIMEOUT_MS (30s)
 *   4) Thêm ghi log vào sheet Audit qua logAudit_() cho các hành động ghi dữ liệu
 *   5) Bỏ doGet() trùng lặp (đã có doGet() ở đầu Code.gs phục vụ chung cho cả web app)
 *   Hàm sanitize() dùng để tránh CSV/Formula Injection khi ghi text tự do
 *   vào sheet (đã có sẵn từ bản gốc, giữ nguyên).
 *********************************************************/

// ==========================================
// HÀM BẢO MẬT: Chống chèn mã độc (Injection)
// ==========================================
function sanitize(val) {
  if (typeof val !== 'string') return val;
  var str = val.trim();
  if (/^[=\+\-@]/.test(str)) {
    return "'" + str; 
  }
  return str;
}

function kiemTraVaTaoTieuDeSheets() {
  var cache = CacheService.getScriptCache();
  if (cache.get('sheets_ready_v5') === '1') return;

  var ss = SpreadsheetApp.openById(KHODAM_CONFIG.SPREADSHEET_ID);
  var sheetGD = ss.getSheetByName(KHODAM_CONFIG.SHEET_GIAODICH) || ss.insertSheet(KHODAM_CONFIG.SHEET_GIAODICH);
  if (sheetGD.getLastRow() === 0) {
    sheetGD.appendRow(["Mã phiếu", "Thời gian", "Loại", "Hình thức N/X", "Đợt vét bãi", "Kho xuất", "Kho nhập", "Khối lượng ướt (MT)", "Độ khô", "Tỷ lệ tiêu hao", "Khối lượng khô (BDMT)", "Trạng thái", "Mã phiếu gốc", "Diễn giải độ khô"]);
    sheetGD.getRange(1, 1, 1, 14).setFontWeight("bold").setBackground("#006b5a").setFontColor("#ffffff");
  } else if (sheetGD.getLastColumn() < 14) {
    sheetGD.getRange(1, 14).setValue("Diễn giải độ khô").setFontWeight("bold").setBackground("#006b5a").setFontColor("#ffffff");
  }

  var sheetCfg = ss.getSheetByName(KHODAM_CONFIG.SHEET_CAUHINH) || ss.insertSheet(KHODAM_CONFIG.SHEET_CAUHINH);
  if (sheetCfg.getLastRow() === 0) {
    sheetCfg.appendRow(["Loại cấu hình", "Kỳ Vét Bãi", "Khoảng Thời Gian", "Mức Tiêu Hao"]);
    sheetCfg.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#0056b3").setFontColor("#ffffff");
  }

  var sheetDK = ss.getSheetByName(KHODAM_CONFIG.SHEET_NHAPDOKHO) || ss.insertSheet(KHODAM_CONFIG.SHEET_NHAPDOKHO);
  if (sheetDK.getLastRow() === 0) {
    sheetDK.appendRow(["Ngày nhập", "Hình thức", "Độ Khô", "Độ Ấm", "Trạng thái"]);
    sheetDK.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#d9822b").setFontColor("#ffffff");
  } else if (sheetDK.getLastColumn() < 5) {
    sheetDK.getRange(1, 5).setValue("Trạng thái").setFontWeight("bold").setBackground("#d9822b").setFontColor("#ffffff");
  }

  var sheetDM = ss.getSheetByName(KHODAM_CONFIG.SHEET_DANHMUCKHO) || ss.insertSheet(KHODAM_CONFIG.SHEET_DANHMUCKHO);
  if (sheetDM.getLastRow() === 0) {
    sheetDM.appendRow(["Mã kho", "Tên Nhà Máy", "Tên Kho Hàng", "Ngày Khởi Tạo", "Trạng thái"]);
    sheetDM.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#20c997").setFontColor("#ffffff");
  }

  cache.put('sheets_ready_v5', '1', 21600);
}

function taiDanhSachKyVetBaiCache() {
  var ss = SpreadsheetApp.openById(KHODAM_CONFIG.SPREADSHEET_ID);
  var sheetCfg = ss.getSheetByName(KHODAM_CONFIG.SHEET_CAUHINH);
  var arrKy = [];
  if (!sheetCfg || sheetCfg.getLastRow() <= 1) return arrKy;
  var data = sheetCfg.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var loaiCfg = String(data[i][0]).trim().toLowerCase();
    if (loaiCfg === "thông số kho" || loaiCfg === "thong so kho") {
      var tenKy = String(data[i][1]).trim();
      var rawTimeStr = String(data[i][2]);
      var parts = rawTimeStr.split(" - ");
      var tieuHaoVal = 0.0315;
      var m = String(data[i][3]).match(/[\d.]+/);
      if (m) tieuHaoVal = parseFloat(m[0]) / 100;

      if (parts.length === 2) {
        var tuStr = parts[0].replace("Từ:", "").replace("Từ", "").trim();
        var denStr = parts[1].replace("Đến:", "").replace("Đến", "").trim();
        var tuTime = new Date(tuStr).setHours(0,0,0,0);
        var denTime = new Date(denStr).setHours(23,59,59,999);

        if (!isNaN(tuTime) && !isNaN(denTime)) {
          arrKy.push({ tenKy: tenKy, tu: tuTime, den: denTime, tieuHao: tieuHaoVal, isLocked: false });
        }
      }
    }
  }

  arrKy.sort(function(a, b) { return a.tu - b.tu; });
  for (var j = 0; j < arrKy.length; j++) arrKy[j].isLocked = (j < arrKy.length - 1);
  return arrKy;
}

function kiemTraKhoaKyVetBaiPure(ngayTime, arrKyCache) {
  var t = new Date(ngayTime).setHours(0,0,0,0);
  for (var i = 0; i < arrKyCache.length; i++) {
    if (t >= arrKyCache[i].tu && t <= arrKyCache[i].den) return arrKyCache[i].isLocked;
  }
  return false;
}

function layThongSoVaTieuHaoCache(ngay, arrKyCache) {
  var t = new Date(ngay).setHours(0,0,0,0);
  var kq = { dotVetBai: "Mặc định", tieuHao: 0.0315 };

  var matchedKys = [];
  for (var i = 0; i < arrKyCache.length; i++) {
    if (t >= arrKyCache[i].tu && t <= arrKyCache[i].den) matchedKys.push(arrKyCache[i]);
  }

  if (matchedKys.length > 0) {
    matchedKys.sort(function(a, b) { return b.tu - a.tu; });
    kq.dotVetBai = matchedKys[0].tenKy;
    kq.tieuHao = matchedKys[0].tieuHao;
    return kq;
  }

  if (arrKyCache.length > 0) {
    var latestKy = arrKyCache[arrKyCache.length - 1];
    kq.dotVetBai = latestKy.tenKy;
    kq.tieuHao = latestKy.tieuHao;
  }
  return kq;
}

function chuanHoaNgay(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  if (typeof val === 'number') {
    var d = new Date((val - (25567 + 2)) * 86400 * 1000);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  var s = String(val).trim();
  var parts = s.split(/[\/\-]/);
  if (parts.length === 3) {
    if (parts[0].length === 4) return parts[0] + "-" + parts[1].padStart(2, '0') + "-" + parts[2].padStart(2, '0');
    else {
      var p1 = parseInt(parts[0], 10), p2 = parseInt(parts[1], 10), p3 = parts[2].length === 2 ? "20" + parts[2] : parts[2];
      if (p1 > 12) return p3 + "-" + String(p2).padStart(2, '0') + "-" + String(p1).padStart(2, '0');
      else return p3 + "-" + String(p1).padStart(2, '0') + "-" + String(p2).padStart(2, '0');
    }
  }
  return s;
}

// ==========================================
// HẰNG SỐ & TIỆN ÍCH DÙNG CHUNG: mặc định khoảng thời gian 3 tháng gần nhất + phân trang 20/lần
// ==========================================
var KICH_THUOC_TRANG = 20;

/** Mốc thời gian (ms, 00:00:00) của đúng ngày này 3 tháng trước — dùng làm mặc định "Từ ngày" khi tải danh sách. */
function layThoiDiem3ThangTruoc() {
  var d = new Date();
  d.setMonth(d.getMonth() - 3);
  d.setHours(0, 0, 0, 0);
  return d.getTime();
}

/** Chuỗi "yyyy-MM-dd" của đúng ngày này 3 tháng trước — dùng làm giá trị mặc định hiển thị trên input date. */
function layNgay3ThangTruocStr() {
  var d = new Date();
  d.setMonth(d.getMonth() - 3);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function tinhDoKhoTrungBinhTheoKho(khoName, ngay) {
  var ss = SpreadsheetApp.openById(KHODAM_CONFIG.SPREADSHEET_ID);
  var sheetGD = ss.getSheetByName(KHODAM_CONFIG.SHEET_GIAODICH);
  if (!sheetGD || sheetGD.getLastRow() <= 1) return 0.45;
  var data = sheetGD.getDataRange().getValues();
  var dtCheck = new Date(ngay);
  var m = dtCheck.getMonth(), y = dtCheck.getFullYear();
  var tTuoi = 0, tKho = 0, count = 0;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][11]).trim() === "Đã hủy") continue;
    var dt = new Date(data[i][1]);
    var kN = String(data[i][6]).trim();
    if (data[i][2] === "NHẬP" && kN === khoName && dt.getMonth() === m && dt.getFullYear() === y) {
      var uot = parseFloat(data[i][7]) || 0;
      var kho = parseFloat(data[i][10]) || 0;
      if (uot > 0 && kho > 0) {
        tTuoi += uot; tKho += kho; count++;
      }
    }
  }
  return (tTuoi > 0 && tKho > 0) ? (tKho / tTuoi) : 0.45;
}

function xuLyXacDinhDoKho(loaiPhieu, hinhThuc, khoLienQuan, ngay, doKhoInput, optionChon) {
  // FIX #8: trước đây điều kiện "val > 0" khiến độ khô=0 bị coi giống hệt "để
  // trống", nên luôn bị ghi đè bằng độ khô tham chiếu/trung bình kho ở phía dưới.
  // Điều này sai với phiếu điều chỉnh Tươi (MT) thuần túy — vốn CHỦ ĐỘNG cần độ
  // khô = 0% để BDMT = MT × 0 = 0, không ảnh hưởng số Khô đã cân bằng ở Bước 1.
  // Nay phân biệt rõ: chỉ coi là "để trống" khi input thực sự rỗng/null/undefined;
  // nếu có giá trị hợp lệ (kể cả đúng 0) thì tôn trọng đúng giá trị đó.
  if (doKhoInput !== "" && doKhoInput !== null && doKhoInput !== undefined) {
    var valNhap = parseFloat(doKhoInput);
    if (!isNaN(valNhap) && valNhap >= 0) {
      if (valNhap > 1) valNhap = valNhap / 100;
      return {
        doKho: valNhap,
        dienGiai: valNhap === 0 ? "Điều chỉnh Tươi thuần túy (Độ khô = 0%)" : "Tự điền độ khô (" + (valNhap * 100).toFixed(2) + "%)"
      };
    }
  }

  var ss = SpreadsheetApp.openById(KHODAM_CONFIG.SPREADSHEET_ID);
  var targetDateStr = chuanHoaNgay(ngay);

  if (loaiPhieu === "NHẬP" && hinhThuc === "TP") {
    var sheetDK = ss.getSheetByName(KHODAM_CONFIG.SHEET_NHAPDOKHO);
    var foundVal = null;
    if (sheetDK && sheetDK.getLastRow() > 1) {
      var dataDK = sheetDK.getDataRange().getValues();
      for (var i = dataDK.length - 1; i >= 1; i--) {
        if (String(dataDK[i][4]).trim() === "Đã hủy") continue; 
        if (dataDK[i][0] && chuanHoaNgay(dataDK[i][0]) === targetDateStr) {
          var v = parseFloat(dataDK[i][2]);
          if (!isNaN(v) && v > 0) {
            foundVal = v > 1 ? v / 100 : v;
            break;
          }
        }
      }
    }

    if (foundVal !== null) {
      return { doKho: foundVal, dienGiai: "Tham chiếu Nhapdokho ngày " + targetDateStr };
    } else {
      if (optionChon === "1") {
        if (sheetDK && sheetDK.getLastRow() > 1) {
          var dataDK = sheetDK.getDataRange().getValues();
          var targetTime = new Date(ngay).getTime();
          var closestVal = 0.45, closestDate = "", minDiff = Infinity;
          for (var i = 1; i < dataDK.length; i++) {
            if (String(dataDK[i][4]).trim() === "Đã hủy") continue; 
            if (dataDK[i][0]) {
              var dTime = new Date(dataDK[i][0]).getTime();
              var diff = targetTime - dTime;
              if (diff > 0 && diff < minDiff) {
                minDiff = diff;
                var v = parseFloat(dataDK[i][2]);
                closestVal = v > 1 ? v / 100 : v;
                closestDate = chuanHoaNgay(dataDK[i][0]);
              }
            }
          }
          return { doKho: closestVal, dienGiai: "Lấy độ khô ngày trước đó (" + closestDate + ")" };
        }
      } else if (optionChon === "2") {
        var tb = tinhDoKhoTrungBinhTheoKho(khoLienQuan, ngay);
        return { doKho: tb, dienGiai: "Độ khô trung bình tháng (" + (tb * 100).toFixed(2) + "%)" };
      }
      return { doKho: 0.45, dienGiai: "Mặc định 45% (Chưa có dữ liệu độ khô ngày " + targetDateStr + ")" };
    }
  }

  if (loaiPhieu === "XUẤT" || hinhThuc.indexOf("MUON") === 0 || hinhThuc === "DC" || hinhThuc === "Khác" || hinhThuc === "XB" || hinhThuc === "TC") {
    var tbKhoXuat = tinhDoKhoTrungBinhTheoKho(khoLienQuan, ngay);
    return { doKho: tbKhoXuat, dienGiai: "Độ khô trung bình kho xuất [" + khoLienQuan + "] (" + (tbKhoXuat * 100).toFixed(2) + "%)" };
  }

  return { doKho: 0.45, dienGiai: "Mặc định 45%" };
}

function xuLyNhapSanPhamSanXuat(ngayNhap, optionChon) {
  try {
    var idFileCan = "1vqMVxccBA7zlAMHrGsVBydGFwZJ6QuDZW10zJ74V29g";
    var ssCan = SpreadsheetApp.openById(idFileCan);
    var sheetCan = ssCan.getSheetByName("PhieuCan_DN") || ssCan.getSheets()[0];

    if (!sheetCan) return { status: "error", message: "Không tìm thấy sheet PhieuCan_DN trong file trạm cân!" };

    var dataCan = sheetCan.getDataRange().getValues();
    if (dataCan.length <= 1) return { status: "success", tongKhoiLuongUot: 0, doKho: 0.45, dotVetBai: "Mặc định", tongGoKeo: 0, dienGiaiDoKho: "Mặc định 45%" };

    var targetDateStr = chuanHoaNgay(ngayNhap);
    var arrKyCache = taiDanhSachKyVetBaiCache();
    var thongSo = layThongSoVaTieuHaoCache(ngayNhap, arrKyCache);

    var tongKhoiLuongGoKeo = 0, colDateIdx = 1, colWeightIdx = 9;

    for (var i = 1; i < dataCan.length; i++) {
      var row = dataCan[i];
      if (row[colDateIdx] !== undefined && row[colDateIdx] !== "") {
        if (chuanHoaNgay(row[colDateIdx]) === targetDateStr) {
          tongKhoiLuongGoKeo += parseFloat(row[colWeightIdx]) || 0;
        }
      }
    }

    var khoiLuongDamTuoi = (tongKhoiLuongGoKeo - (tongKhoiLuongGoKeo * thongSo.tieuHao)) / 1000;
    var ss = SpreadsheetApp.openById(KHODAM_CONFIG.SPREADSHEET_ID);
    var sheetDK = ss.getSheetByName(KHODAM_CONFIG.SHEET_NHAPDOKHO);
    var hasExactDateDK = false;
    
    if (sheetDK && sheetDK.getLastRow() > 1) {
      var dataDK = sheetDK.getDataRange().getValues();
      for (var i = 1; i < dataDK.length; i++) {
        if (String(dataDK[i][4]).trim() === "Đã hủy") continue;
        if (dataDK[i][0] && chuanHoaNgay(dataDK[i][0]) === targetDateStr) {
          hasExactDateDK = true; break;
        }
      }
    }

    if (!hasExactDateDK && !optionChon) {
      return {
        status: "need_dokho_choice",
        message: "⚠️ Chưa có độ khô ngày " + targetDateStr + " trong sheet Nhapdokho!",
        ngayNhap: ngayNhap, dotVetBai: thongSo.dotVetBai, tieuHao: thongSo.tieuHao,
        tongGoKeo: tongKhoiLuongGoKeo, tongKhoiLuongUot: khoiLuongDamTuoi
      };
    }

    var refDK = xuLyXacDinhDoKho("NHẬP", "TP", "", ngayNhap, 0, optionChon);
    return {
      status: "success", ngayNhap: ngayNhap, dotVetBai: thongSo.dotVetBai, tieuHao: thongSo.tieuHao,
      tongGoKeo: tongKhoiLuongGoKeo, tongKhoiLuongUot: khoiLuongDamTuoi, doKho: refDK.doKho,
      dienGiaiDoKho: refDK.dienGiai, bdmt: (khoiLuongDamTuoi * refDK.doKho)
    };
  } catch (err) {
    return { status: "error", message: "Lỗi: " + err.toString() };
  }
}

function layDanhSachTenKho() {
  var cache = CacheService.getScriptCache();
  var cachedKho = cache.get('danh_sach_ten_kho');
  if (cachedKho) { try { return JSON.parse(cachedKho); } catch(e) {} }

  kiemTraVaTaoTieuDeSheets();
  var ss = SpreadsheetApp.openById(KHODAM_CONFIG.SPREADSHEET_ID);
  var sheetDM = ss.getSheetByName(KHODAM_CONFIG.SHEET_DANHMUCKHO);
  var khoList = [];
  if (sheetDM && sheetDM.getLastRow() > 1) {
    var data = sheetDM.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][4]).trim() === "Đã hủy") continue;
      if (data[i][2] && khoList.indexOf(data[i][2]) === -1) khoList.push(data[i][2]);
    }
  }
  cache.put('danh_sach_ten_kho', JSON.stringify(khoList), 300);
  return khoList;
}

function xoaCacheKho() { CacheService.getScriptCache().remove('danh_sach_ten_kho'); }

function getDataForGiaoDichForm(loai, tab, tuNgayStr, denNgayStr, trang) {
  var khoList = layDanhSachTenKho();
  var ketQuaPhieu = layDanhSachGiaoDichTheoBoLoc(loai, tab, tuNgayStr, denNgayStr, trang);
  return {
    khoList: khoList,
    danhSach: ketQuaPhieu.list,
    trang: ketQuaPhieu.trang,
    tongSoTrang: ketQuaPhieu.tongSoTrang,
    tongSoPhieu: ketQuaPhieu.tongSoPhieu,
    tuNgay: ketQuaPhieu.tuNgay,
    denNgay: ketQuaPhieu.denNgay
  };
}

/**
 * Lấy danh sách phiếu theo bộ lọc, có PHÂN TRANG (20 phiếu/trang) để tránh tải quá nhiều dữ liệu 1 lần.
 * Mặc định khoảng thời gian: 3 tháng gần nhất tính từ hiện tại (nếu không truyền tuNgay/denNgay).
 * @param {number} [trang] - Số trang cần lấy (bắt đầu từ 1). Mặc định: 1.
 */
function layDanhSachGiaoDichTheoBoLoc(loaiPhieu, tabPhieu, tuNgay, denNgay, trang) {
  kiemTraVaTaoTieuDeSheets();
  var ss = SpreadsheetApp.openById(KHODAM_CONFIG.SPREADSHEET_ID);
  var sheetGD = ss.getSheetByName(KHODAM_CONFIG.SHEET_GIAODICH);

  var tuNgayStr = tuNgay || layNgay3ThangTruocStr();
  var denNgayStr = denNgay || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var soTrang = parseInt(trang, 10);
  if (isNaN(soTrang) || soTrang < 1) soTrang = 1;

  if (!sheetGD || sheetGD.getLastRow() <= 1) {
    return { list: [], trang: 1, tongSoTrang: 1, tongSoPhieu: 0, tuNgay: tuNgayStr, denNgay: denNgayStr };
  }

  var data = sheetGD.getDataRange().getValues();
  var tuDate = tuNgay ? new Date(tuNgay).setHours(0, 0, 0, 0) : layThoiDiem3ThangTruoc();
  var denDate = new Date(denNgayStr).setHours(23, 59, 59, 999);
  var arrKyCache = taiDanhSachKyVetBaiCache();

  // Bước 1: lọc ra chỉ số dòng phù hợp bộ lọc (không dựng object ngay để đỡ tốn công cho các phiếu không thuộc trang cần lấy)
  var matchedRowIdx = [];
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][11]).trim() === "Đã hủy") continue;

    var loaiVal = data[i][2];
    var hinhThucVal = data[i][3];
    var matchTab = false;

    if (loaiPhieu === "NHẬP") {
      if (tabPhieu === "Dăm sản xuất" && hinhThucVal === "TP") matchTab = true;
      if (tabPhieu === "Nhập khác" && hinhThucVal !== "TP") matchTab = true;
    } else {
      if (tabPhieu === "Xuất bán" && hinhThucVal === "TT") matchTab = true;
      if (tabPhieu === "Xuất khác" && hinhThucVal !== "TT" && hinhThucVal !== "TC") matchTab = true;
      if (tabPhieu === "Xuất trung chuyển" && hinhThucVal === "TC") matchTab = true;
    }

    if (loaiVal === loaiPhieu && matchTab && data[i][1]) {
      var dTimeOnly = new Date(new Date(data[i][1])).setHours(0, 0, 0, 0);
      if (tuDate && dTimeOnly < tuDate) continue;
      if (denDate && dTimeOnly > denDate) continue;
      matchedRowIdx.push(i);
    }
  }

  var tongSoPhieu = matchedRowIdx.length;
  var tongSoTrang = Math.max(1, Math.ceil(tongSoPhieu / KICH_THUOC_TRANG));
  if (soTrang > tongSoTrang) soTrang = tongSoTrang;
  var batDau = (soTrang - 1) * KICH_THUOC_TRANG;
  var trangRowIdx = matchedRowIdx.slice(batDau, batDau + KICH_THUOC_TRANG);

  var list = trangRowIdx.map(function(i) {
    var dObj = new Date(data[i][1]); // không mutate: giữ nguyên giờ phút để hiển thị đúng
    return {
      maPhieu: data[i][0],
      ngay: Utilities.formatDate(dObj, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"),
      hinhThuc: data[i][3],
      dotVetBai: data[i][4],
      khoXuat: data[i][5],
      khoNhap: data[i][6],
      mt: parseFloat(data[i][7]) || 0,
      doKho: parseFloat(data[i][8]) || 0,
      bdmt: parseFloat(data[i][10]) || 0,
      nguonDK: data[i][13] || "Mặc định/Thủ công",
      isLocked: kiemTraKhoaKyVetBaiPure(dObj, arrKyCache)
    };
  });

  return {
    list: list,
    trang: soTrang,
    tongSoTrang: tongSoTrang,
    tongSoPhieu: tongSoPhieu,
    tuNgay: tuNgayStr,
    denNgay: denNgayStr
  };
}

function taoMaPhieuMoi(prefix) {
  var tsp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  var randSuffix = Math.floor(Math.random() * 9000) + 1000;
  return prefix + tsp + "_" + randSuffix;
}

function xuLySuaXoaGiaoDich(dataEdit) {
  var ss = SpreadsheetApp.openById(KHODAM_CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(KHODAM_CONFIG.SHEET_GIAODICH);
  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;

  if (dataEdit.maPhieu) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(dataEdit.maPhieu).trim()) { rowIdx = i + 1; break; }
    }
  }

  var ngayCheck = dataEdit.ngay || (rowIdx > -1 ? data[rowIdx-1][1] : new Date());
  var arrKyCache = taiDanhSachKyVetBaiCache();
  if (kiemTraKhoaKyVetBaiPure(ngayCheck, arrKyCache)) return "❌ Khóa bảo mật: Phiếu thuộc kỳ cũ!";

  if (dataEdit.hanhDong === "XOA") {
    if (rowIdx > -1) { 
      sheet.getRange(rowIdx, 12).setValue("Đã hủy"); 
      return "🗑️ Đã xóa phiếu (cập nhật trạng thái 'Đã hủy')."; 
    }
    return "❌ Không tìm thấy.";
  } else {
    var ngay = new Date(dataEdit.ngay);
    var thongSo = layThongSoVaTieuHaoCache(ngay, arrKyCache);
    var mt = parseFloat(dataEdit.khoiLuongMT) || 0;
    
    var dienGiaiDK = sanitize(dataEdit.dienGiaiDoKho || "");
    var hinhThucVal = sanitize(dataEdit.hinhThuc);

    if (dataEdit.loai === "NHẬP" && hinhThucVal !== "TP" && hinhThucVal !== "DC") {
      var dCheck = parseFloat(dataEdit.doKho);
      if (isNaN(dCheck) || dCheck <= 0) {
        return "❌ Lỗi: Nhập hình thức [" + hinhThucVal + "] bắt buộc phải điền độ khô, không được để trống!";
      }
    }

    if (hinhThucVal === "TP" && mt === 0) {
      var kqScan = xuLyNhapSanPhamSanXuat(dataEdit.ngay, dataEdit.optionChon);
      if (kqScan.status === "error") return "❌ " + kqScan.message;
      mt = kqScan.tongKhoiLuongUot;
      if (!dienGiaiDK) dienGiaiDK = kqScan.dienGiaiDoKho;
    }

    var khoXuatValRaw = "Không có", khoNhapValRaw = "Không có";
    if (dataEdit.loai === "NHẬP") {
      khoNhapValRaw = dataEdit.khoNhap || "Không có";
      if (hinhThucVal !== "TP" && hinhThucVal !== "Nhập khác") khoXuatValRaw = dataEdit.khoXuat || "Không có";
    } else {
      khoXuatValRaw = dataEdit.khoXuat || "Không có";
      if (hinhThucVal === "TC") khoNhapValRaw = dataEdit.khoNhap || "Không có";
    }

    var khoXuatVal = sanitize(khoXuatValRaw);
    var khoNhapVal = sanitize(khoNhapValRaw);
    var khoLienQuan = (dataEdit.loai === "NHẬP") ? khoNhapVal : khoXuatVal;

    var refDK = xuLyXacDinhDoKho(dataEdit.loai, hinhThucVal, khoLienQuan, ngay, dataEdit.doKho, dataEdit.optionChon);
    var doKho = refDK.doKho;
    if (!dienGiaiDK) dienGiaiDK = refDK.dienGiai;

    var tyLeTieuHao = thongSo.tieuHao || 0.0315;
    var dotVetBai = thongSo.dotVetBai;
    var bdmt = mt * doKho;
    var ngayFormatted = Utilities.formatDate(ngay, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

    if (rowIdx > -1 && dataEdit.hanhDong === "SUA") {
      sheet.getRange(rowIdx, 2, 1, 11).setValues([[
        ngayFormatted, dataEdit.loai, hinhThucVal, dotVetBai,
        khoXuatVal, khoNhapVal, mt, doKho, tyLeTieuHao, bdmt, "Hợp lệ"
      ]]);
      sheet.getRange(rowIdx, 14).setValue(dienGiaiDK);
      return "✏️ Đã cập nhật phiếu thành công.";
    } else {
      if (hinhThucVal === "TP" && !dataEdit.boQuaTrung) {
        var ngayKiemTra = Utilities.formatDate(ngay, Session.getScriptTimeZone(), "yyyy-MM-dd");
        var maTrungLap = [];
        for (var r = 1; r < data.length; r++) {
          if (String(data[r][11]).trim() === "Đã hủy") continue;
          if (String(data[r][2]).trim() === "NHẬP" && String(data[r][3]).trim() === "TP" && String(data[r][6]).trim() === khoNhapVal) {
            var dRow = new Date(data[r][1]);
            if (!isNaN(dRow.getTime()) && Utilities.formatDate(dRow, Session.getScriptTimeZone(), "yyyy-MM-dd") === ngayKiemTra) {
              maTrungLap.push(String(data[r][0]));
            }
          }
        }
        if (maTrungLap.length > 0) {
          return "⚠️TRUNG_LAP_TP::Đã có " + maTrungLap.length + " phiếu Nhập Dăm sản xuất (TP) ngày " + ngayKiemTra +
                 " tại kho [" + khoNhapVal + "] (Mã: " + maTrungLap.join(", ") + "). Bạn có chắc muốn tạo thêm?";
        }
      }

      var prefix = dataEdit.loai === "NHẬP" ? "NK_" : "XK_";
      var maPhieuChinh = taoMaPhieuMoi(prefix);

      sheet.appendRow([
        maPhieuChinh, ngayFormatted, dataEdit.loai, hinhThucVal, dotVetBai,
        khoXuatVal, khoNhapVal, mt, doKho, tyLeTieuHao, bdmt, "Hợp lệ", "", dienGiaiDK
      ]);

      if (dataEdit.loai === "XUẤT" && hinhThucVal === "TC" && khoNhapVal !== "Không có") {
        return "✅ Xuất trung chuyển sang [" + khoNhapVal + "] thành công!";
      }
      return "✅ Đã tạo mới phiếu " + dataEdit.loai + " thành công.";
    }
  }
}

// ==========================================
// HOÀN THÀNH ĐƠN HÀNG XUẤT BÁN (Đối soát Khô/Tươi)
// ==========================================

/**
 * Tổng hợp số liệu Nhập/Xuất liên quan đến 1 kho xuất bán trong khoảng thời gian.
 * Tách riêng để dùng chung cho cả 2 bước đối soát (Khô BDMT & Tươi MT).
 *
 * LƯU Ý QUAN TRỌNG: Mọi phiếu Trung chuyển (hình thức "TC") trong hệ thống — kể cả phiếu điều
 * chỉnh 1 chiều do chính chức năng này tạo ra — đều được lưu với Loại = "XUẤT" (vì được lập từ
 * màn hình "Xuất kho > Xuất trung chuyển"). Vì vậy KHÔNG được lọc theo Loại="NHẬP" để tìm phiếu
 * nhập trung chuyển; phải xét độc lập cột Kho Xuất / Kho Nhập của các phiếu hình thức TC, bất kể Loại.
 */
function tongHopSoLieuKhoXuatBan(khoXuatBan, tuTime, denTime) {
  var ss = SpreadsheetApp.openById(KHODAM_CONFIG.SPREADSHEET_ID);
  var sheetGD = ss.getSheetByName(KHODAM_CONFIG.SHEET_GIAODICH);
  var kq = {
    nhapTC_BDMT: 0, nhapTC_MT: 0,     // Nhập trung chuyển VÀO kho này (gồm cả phiếu điều chỉnh TC 1 chiều)
    xuatBan_BDMT: 0, xuatBan_MT: 0,   // Xuất bán (TT) TỪ kho này
    xuatKhac_BDMT: 0, xuatKhac_MT: 0  // Xuất trung chuyển RA khỏi kho + các xuất khác (ĐC, mượn...) TỪ kho này
  };
  if (!sheetGD || sheetGD.getLastRow() <= 1) return kq;

  var data = sheetGD.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][11]).trim() === "Đã hủy") continue;
    var dt = new Date(data[i][1]).getTime();
    if (dt < tuTime || dt > denTime) continue;

    var loai = data[i][2], hinhThuc = String(data[i][3]).trim();
    var kX = String(data[i][5]).trim(), kN = String(data[i][6]).trim();
    var mt = parseFloat(data[i][7]) || 0, bdmt = parseFloat(data[i][10]) || 0;

    if (hinhThuc === "TC") {
      // Xét độc lập 2 chiều theo đúng cột Kho Xuất / Kho Nhập, KHÔNG phụ thuộc "Loại"
      if (kN === khoXuatBan) { kq.nhapTC_BDMT += bdmt; kq.nhapTC_MT += mt; }
      if (kX === khoXuatBan) { kq.xuatKhac_BDMT += bdmt; kq.xuatKhac_MT += mt; }
    } else if (loai === "XUẤT" && kX === khoXuatBan) {
      if (hinhThuc === "TT") { kq.xuatBan_BDMT += bdmt; kq.xuatBan_MT += mt; }
      else { kq.xuatKhac_BDMT += bdmt; kq.xuatKhac_MT += mt; }
    }
  }
  return kq;
}

/** Tìm kho đã "Xuất trung chuyển" (TC) nhiều BDMT nhất VÀO khoXuatBan trong khoảng thời gian — dùng làm gợi ý mặc định cho Nút 1. */
function timKhoNguonTCLonNhat(khoXuatBan, tuTime, denTime) {
  var ss = SpreadsheetApp.openById(KHODAM_CONFIG.SPREADSHEET_ID);
  var sheetGD = ss.getSheetByName(KHODAM_CONFIG.SHEET_GIAODICH);
  if (!sheetGD || sheetGD.getLastRow() <= 1) return null;
  var data = sheetGD.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][11]).trim() === "Đã hủy") continue;
    var dt = new Date(data[i][1]).getTime();
    if (dt < tuTime || dt > denTime) continue;
    if (String(data[i][3]).trim() !== "TC") continue;

    var kX = String(data[i][5]).trim();
    var kN = String(data[i][6]).trim();
    if (kN !== khoXuatBan) continue;
    if (!kX || kX === "Không có") continue;

    var bdmt = parseFloat(data[i][10]) || 0;
    map[kX] = (map[kX] || 0) + bdmt;
  }
  var best = null, bestVal = -Infinity;
  for (var k in map) { if (map[k] > bestVal) { bestVal = map[k]; best = k; } }
  return best;
}

/** Bước 1 (nút 1): Đối soát chênh lệch KHÔ (BDMT) giữa Nhập trung chuyển và Xuất bán. */
function taoPhieuDieuChinhKho(khoXuatBan, tuTime, denTime, denNgayGoc) {
  var sl = tongHopSoLieuKhoXuatBan(khoXuatBan, tuTime, denTime);
  var bdmtKhaDung = sl.nhapTC_BDMT - sl.xuatKhac_BDMT;
  // Guard chia cho 0: nếu không có phiếu nhập trung chuyển nào trong kỳ thì lấy độ khô mặc định 45%
  var doKhoTB = sl.nhapTC_MT > 0 ? (sl.nhapTC_BDMT / sl.nhapTC_MT) : 0.45;
  var chenhLechBDMT = sl.xuatBan_BDMT - bdmtKhaDung;
  var mt = doKhoTB > 0 ? Math.abs(chenhLechBDMT) / doKhoTB : 0;
  // Gợi ý mặc định: kho đã trung chuyển vào kho xuất bán nhiều nhất trong kỳ (người dùng vẫn có thể đổi kho khác)
  var khoNguonTC = timKhoNguonTCLonNhat(khoXuatBan, tuTime, denTime);

  return "DATA_STEP1::" + JSON.stringify({
    chenhLechKho: chenhLechBDMT,
    doKho: doKhoTB,
    mt: mt,
    denNgay: denNgayGoc,
    khoXuatBan: khoXuatBan,
    khoNguonTC: khoNguonTC || null
  });
}

/**
 * Bước 2 (nút 2): Đối soát chênh lệch TƯƠI (MT) do hao hụt ẩm, sau khi Khô đã khớp.
 * BẮT BUỘC Bước 1 (Khô/BDMT) phải cân bằng trước — nếu người dùng vừa lưu phiếu điều
 * chỉnh Khô ở Bước 1 thì tongHopSoLieuKhoXuatBan() sẽ tự động lấy dữ liệu MỚI NHẤT từ
 * sheet (bao gồm cả phiếu Bước 1 vừa tạo) vì hàm luôn đọc lại dữ liệu tại thời điểm gọi.
 */
function taoPhieuDieuChinhTuoi(khoXuatBan, tuTime, denTime, tuNgayGoc, denNgayGoc) {
  var sl = tongHopSoLieuKhoXuatBan(khoXuatBan, tuTime, denTime);
  var bdmtKhaDung = sl.nhapTC_BDMT - sl.xuatKhac_BDMT;
  var mtKhaDung = sl.nhapTC_MT - sl.xuatKhac_MT;
  var chenhLechMT = mtKhaDung - sl.xuatBan_MT;
  var chenhLechBDMT = bdmtKhaDung - sl.xuatBan_BDMT;

  // GUARD: Bước 1 (Khô/BDMT) phải cân bằng trước khi được phép tính Bước 2 (Tươi/MT).
  // Nếu còn chênh lệch BDMT nghĩa là người dùng chưa lưu phiếu điều chỉnh Khô ở Bước 1.
  // FIX #7: ngưỡng "coi như cân bằng" đổi từ 0.001 -> 0.01 để KHỚP với 2 chữ số thập
  // phân đang hiển thị (.toFixed(2)). Ngưỡng cũ 0.001 quá nhỏ so với độ chính xác hiển
  // thị: một chênh lệch thực tế như -0.0015 vẫn > 0.001 nên bị chặn, nhưng khi hiển thị
  // ra màn hình lại làm tròn thành "-0.00" (trông như đã bằng 0) khiến người dùng bối
  // rối không hiểu vì sao hệ thống vẫn báo còn chênh lệch. 0.01 BDMT/MT (~10kg) là sai
  // số dấu phẩy động/làm tròn không đáng kể so với quy mô tấn của kho dăm, không phải
  // chênh lệch nghiệp vụ thật.
  if (Math.abs(chenhLechBDMT) > 0.01) {
    return "❌ Chưa thể thực hiện Bước 2 (Tươi): Lượng Khô (BDMT) tại kho [" + khoXuatBan +
      "] vẫn còn chênh lệch " + chenhLechBDMT.toFixed(2) +
      " BDMT. Vui lòng bấm nút 1️⃣ để tạo VÀ LƯU phiếu điều chỉnh Khô trước, sau đó mới quay lại bấm nút 2️⃣!";
  }

  var tuStr = tuNgayGoc.split(" ")[0];
  var denStr = denNgayGoc.split(" ")[0];

  var msg = "Hiện tại kho Xuất bán (" + khoXuatBan + ") từ ngày " + tuStr + " đến ngày " + denStr +
    " có tổng lượng nhập là " + mtKhaDung.toFixed(2) + " MT, " + bdmtKhaDung.toFixed(2) + " BDMT, " +
    "tổng lượng xuất bán thực tế là " + sl.xuatBan_MT.toFixed(2) + " MT, " + sl.xuatBan_BDMT.toFixed(2) + " BDMT. " +
    "Chênh lệch " + Math.abs(chenhLechMT).toFixed(2) + " MT, " + Math.abs(chenhLechBDMT).toFixed(2) + " BDMT. " +
    (Math.abs(chenhLechMT) <= 0.01 ? "Chênh lệch bằng 0, có thể hoàn thành." : "Khác 0, cần tạo phiếu điều chỉnh tươi!");

  return "DATA_STEP2::" + JSON.stringify({
    chenhLechMT: chenhLechMT,
    denNgay: denNgayGoc,
    khoXuatBan: khoXuatBan,
    msg: msg
  });
}

/** Hàm điều phối chính — gọi từ processFormData('Hoanthanhdonhang', ...) */
function xuLyHoanThanhDonHangXuatBan(dataParam) {
  var khoXuatBan = String((dataParam && dataParam.khoXuatBan) || "").trim();
  if (!khoXuatBan || khoXuatBan === "Không có") return "❌ Vui lòng chọn Kho Xuất Bán hợp lệ!";
  if (!dataParam.tuNgay || !dataParam.denNgay) return "❌ Vui lòng chọn đầy đủ khoảng thời gian!";

  var tuTime = new Date(dataParam.tuNgay).getTime();
  var denTime = new Date(dataParam.denNgay).getTime();
  if (isNaN(tuTime) || isNaN(denTime)) return "❌ Định dạng ngày không hợp lệ!";
  if (tuTime > denTime) return "❌ 'Từ ngày' phải trước 'Đến ngày'!";

  if (dataParam.cheDo === "BDMT") return taoPhieuDieuChinhKho(khoXuatBan, tuTime, denTime, dataParam.denNgay);
  if (dataParam.cheDo === "MT") return taoPhieuDieuChinhTuoi(khoXuatBan, tuTime, denTime, dataParam.tuNgay, dataParam.denNgay);
  return "❌ Chế độ không hợp lệ.";
}

/**
 * FIX #9: Tạo VÀ LƯU TRỰC TIẾP phiếu điều chỉnh từ panel "Hoàn thành đơn hàng",
 * không cần điều hướng qua form Nhập/Xuất kho thủ công rồi tự quay lại tab Xuất
 * bán. Tái sử dụng nguyên vẹn xuLySuaXoaGiaoDich() để giữ đúng toàn bộ logic
 * tính kỳ vét bãi/tiêu hao/khóa kỳ/chống trùng đã có, tránh viết trùng lặp.
 *
 * payload.buoc = 'BDMT': tạo phiếu Xuất trung chuyển (TC).
 *   { buoc:'BDMT', khoXuatBan, ngay:'yyyy-MM-dd', khoXuat, khoNhap, mt, doKho }
 * payload.buoc = 'MT': tạo phiếu Xuất khác (dư tươi) hoặc Nhập khác (thiếu tươi),
 *   hình thức "DC", độ khô CỐ ĐỊNH = 0 (điều chỉnh Tươi thuần túy, không đụng Khô).
 *   { buoc:'MT', khoXuatBan, ngay:'yyyy-MM-dd', chenhLechMT }
 */
function KD_taoPhieuDieuChinhTuDong(payload) {
  payload = payload || {};
  var khoXuatBan = String(payload.khoXuatBan || "").trim();
  if (!khoXuatBan) return "❌ Thiếu thông tin Kho Xuất Bán.";

  if (payload.buoc === "BDMT") {
    var khoXuat = String(payload.khoXuat || "").trim();
    var khoNhap = String(payload.khoNhap || "").trim();
    if (!khoXuat || !khoNhap) return "❌ Vui lòng xác định đủ Kho Xuất và Kho Nhập cho phiếu điều chỉnh Khô.";
    if (khoXuat === khoNhap) return "❌ Kho Xuất và Kho Nhập không được trùng nhau.";

    var dataEditBDMT = {
      loai: "XUẤT", tab: "Xuất trung chuyển", maPhieu: "",
      ngay: payload.ngay, hinhThuc: "TC",
      khoXuat: khoXuat, khoNhap: khoNhap,
      khoiLuongMT: payload.mt, doKho: payload.doKho,
      dienGiaiDoKho: "Tự động điều chỉnh Khô (BDMT) - Hoàn thành đơn hàng xuất bán [" + khoXuatBan + "]",
      hanhDong: "THEM", boQuaTrung: true
    };
    return xuLySuaXoaGiaoDich(dataEditBDMT);
  }

  if (payload.buoc === "MT") {
    var chenhLechMT = parseFloat(payload.chenhLechMT) || 0;
    if (Math.abs(chenhLechMT) <= 0.01) return "❌ Không có chênh lệch Tươi (MT) đáng kể để tạo phiếu.";
    var loaiTarget = chenhLechMT > 0 ? "XUẤT" : "NHẬP";
    var tabTarget = chenhLechMT > 0 ? "Xuất khác" : "Nhập khác";

    var dataEditMT = {
      loai: loaiTarget, tab: tabTarget, maPhieu: "",
      ngay: payload.ngay, hinhThuc: "DC",
      khoXuat: loaiTarget === "XUẤT" ? khoXuatBan : "Không có",
      khoNhap: loaiTarget === "NHẬP" ? khoXuatBan : "Không có",
      khoiLuongMT: Math.abs(chenhLechMT), doKho: 0,
      dienGiaiDoKho: "Tự động điều chỉnh Tươi (MT) do hao hụt ẩm - Hoàn thành đơn hàng xuất bán [" + khoXuatBan + "]",
      hanhDong: "THEM"
    };
    return xuLySuaXoaGiaoDich(dataEditMT);
  }

  return "❌ Bước điều chỉnh không hợp lệ.";
}

function layDanhSachDanhMucKho() {
  kiemTraVaTaoTieuDeSheets();
  var ss = SpreadsheetApp.openById(KHODAM_CONFIG.SPREADSHEET_ID);
  var sheetDM = ss.getSheetByName(KHODAM_CONFIG.SHEET_DANHMUCKHO);
  if (!sheetDM || sheetDM.getLastRow() <= 1) return [];
  var data = sheetDM.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][4]).trim() === "Đã hủy") continue;
    list.push({ maKho: data[i][0], tenNhaMay: data[i][1], tenKho: data[i][2], ngayKhoiTao: Utilities.formatDate(new Date(data[i][3]), Session.getScriptTimeZone(), "yyyy-MM-dd") });
  }
  return list;
}

function xuLyDanhMucKho(dataDM) {
  kiemTraVaTaoTieuDeSheets();
  var ss = SpreadsheetApp.openById(KHODAM_CONFIG.SPREADSHEET_ID);
  var sheetDM = ss.getSheetByName(KHODAM_CONFIG.SHEET_DANHMUCKHO);
  if (dataDM.hanhDong === "XOA") {
    var data = sheetDM.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(dataDM.maKho).trim()) { 
        sheetDM.getRange(i + 1, 5).setValue("Đã hủy"); 
        xoaCacheKho(); return "🗑️ Đã xóa kho."; 
      }
    }
  } else if (dataDM.hanhDong === "THEM") {
    sheetDM.appendRow(["KHO_" + new Date().getTime(), sanitize(dataDM.tenNhaMay), sanitize(dataDM.tenKho), new Date(dataDM.ngayKhoiTao), "Hoạt động"]);
    xoaCacheKho();
    return "✅ Đã thêm kho.";
  }
  return "❌ Lỗi.";
}

function layDanhSachKyVetBai() {
  kiemTraVaTaoTieuDeSheets();
  var ss = SpreadsheetApp.openById(KHODAM_CONFIG.SPREADSHEET_ID);
  var sheetCfg = ss.getSheetByName(KHODAM_CONFIG.SHEET_CAUHINH);
  if (!sheetCfg || sheetCfg.getLastRow() <= 1) return [];
  var data = sheetCfg.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var loaiCfg = String(data[i][0]).trim().toLowerCase();
    if (loaiCfg === "thông số kho" || loaiCfg === "thong so kho") {
      list.push({ rowIndex: i + 1, kyVetBai: data[i][1], thoiGian: data[i][2], tieuHao: data[i][3] });
    }
  }
  for (var j = 0; j < list.length; j++) { list[j].isLocked = (j < list.length - 1); }
  return list.reverse();
}

function xuLyKyVetBai(dataEdit) {
  var ss = SpreadsheetApp.openById(KHODAM_CONFIG.SPREADSHEET_ID);
  var sheetCfg = ss.getSheetByName(KHODAM_CONFIG.SHEET_CAUHINH);
  if (dataEdit.hanhDong === "XOA") {
    sheetCfg.deleteRow(parseInt(dataEdit.rowIndex)); return "🗑️ Đã xóa kỳ.";
  } else if (dataEdit.hanhDong === "THEM") {
    if (!dataEdit.tuNgay) return "❌ Lỗi: Bạn chưa chọn Ngày và Giờ bắt đầu cho kỳ mới!";
    var tuMoi = new Date(dataEdit.tuNgay);
    var denMoi = new Date(tuMoi.getTime() + (730 * 24 * 60 * 60 * 1000));
    var data = sheetCfg.getDataRange().getValues();
    var lastKyRowIdx = -1, lastKyOldStr = "";

    for (var i = 1; i < data.length; i++) {
      var loaiCfg = String(data[i][0]).trim().toLowerCase();
      if (loaiCfg === "thông số kho" || loaiCfg === "thong so kho") {
         lastKyRowIdx = i + 1; lastKyOldStr = String(data[i][2]);
      }
    }

    if (lastKyRowIdx > -1) {
        var parts = lastKyOldStr.split(" - ");
        if (parts.length > 0) {
            sheetCfg.getRange(lastKyRowIdx, 3).setValue(parts[0].trim() + " - Đến: " + Utilities.formatDate(tuMoi, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"));
        }
    }

    var strTuMoi = "Từ: " + Utilities.formatDate(tuMoi, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    var strDenMoi = "Đến: " + Utilities.formatDate(denMoi, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    sheetCfg.appendRow(["Thông Số Kho", sanitize(dataEdit.kyVetBai), strTuMoi + " - " + strDenMoi, "Tiêu hao: " + dataEdit.tieuHao + "%"]);
    return "✅ Đã tạo kỳ mới thành công!";
  }
  return "❌ Lỗi.";
}

/** Lấy danh sách độ khô theo bộ lọc, có phân trang (20/trang). Mặc định 3 tháng gần nhất. */
function layDanhSachDoKhoTheoBoLoc(tuNgay, denNgay, trang) {
  kiemTraVaTaoTieuDeSheets();
  var ss = SpreadsheetApp.openById(KHODAM_CONFIG.SPREADSHEET_ID);
  var sheetDK = ss.getSheetByName(KHODAM_CONFIG.SHEET_NHAPDOKHO);

  var tuNgayStr = tuNgay || layNgay3ThangTruocStr();
  var denNgayStr = denNgay || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var soTrang = parseInt(trang, 10);
  if (isNaN(soTrang) || soTrang < 1) soTrang = 1;

  if (!sheetDK || sheetDK.getLastRow() <= 1) {
    return { list: [], trang: 1, tongSoTrang: 1, tongSoPhieu: 0, tuNgay: tuNgayStr, denNgay: denNgayStr };
  }

  var data = sheetDK.getDataRange().getValues();
  var tuDate = tuNgay ? new Date(tuNgay).setHours(0, 0, 0, 0) : layThoiDiem3ThangTruoc();
  var denDate = new Date(denNgayStr).setHours(23, 59, 59, 999);
  var arrKyCache = taiDanhSachKyVetBaiCache();

  var matchedRowIdx = [];
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][4]).trim() === "Đã hủy") continue;
    if (!data[i][0]) continue;
    var dTimeOnly = new Date(new Date(data[i][0])).setHours(0, 0, 0, 0);
    if (tuDate && dTimeOnly < tuDate) continue;
    if (denDate && dTimeOnly > denDate) continue;
    matchedRowIdx.push(i);
  }

  var tongSoPhieu = matchedRowIdx.length;
  var tongSoTrang = Math.max(1, Math.ceil(tongSoPhieu / KICH_THUOC_TRANG));
  if (soTrang > tongSoTrang) soTrang = tongSoTrang;
  var batDau = (soTrang - 1) * KICH_THUOC_TRANG;
  var trangRowIdx = matchedRowIdx.slice(batDau, batDau + KICH_THUOC_TRANG);

  var list = trangRowIdx.map(function(i) {
    var dObj = new Date(data[i][0]);
    return {
      ngay: Utilities.formatDate(dObj, Session.getScriptTimeZone(), "yyyy-MM-dd"),
      hinhThuc: data[i][1] || "NKSX",
      doKho: parseFloat(data[i][2]) || 0,
      doAm: parseFloat(data[i][3]) || 0,
      isLocked: kiemTraKhoaKyVetBaiPure(dObj, arrKyCache)
    };
  });

  return {
    list: list,
    trang: soTrang,
    tongSoTrang: tongSoTrang,
    tongSoPhieu: tongSoPhieu,
    tuNgay: tuNgayStr,
    denNgay: denNgayStr
  };
}

function xuLySuaXoaDoKho(dataEdit) {
  var target = String(dataEdit.ngay).trim();
  var arrKyCache = taiDanhSachKyVetBaiCache();
  if (kiemTraKhoaKyVetBaiPure(target, arrKyCache)) return "❌ Khóa bảo mật: Ngày thuộc kỳ cũ!";
  var ss = SpreadsheetApp.openById(KHODAM_CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(KHODAM_CONFIG.SHEET_NHAPDOKHO);
  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd") === target) { rowIdx = i + 1; break; }
  }

  if (dataEdit.hanhDong === "XOA") {
    if (rowIdx > -1) { 
      sheet.getRange(rowIdx, 5).setValue("Đã hủy"); 
      return "🗑️ Đã xóa độ khô (cập nhật trạng thái 'Đã hủy')."; 
    }
  } else {
    var hinhThucSanitized = sanitize(dataEdit.hinhThuc || "NKSX");
    var dk = parseFloat(dataEdit.doKho) || 0, da = parseFloat(dataEdit.doAm) || 0;
    if (dk > 1) dk = dk / 100; if (da > 1) da = da / 100;
    
    if (rowIdx > -1) {
      sheet.getRange(rowIdx, 2).setValue(hinhThucSanitized); 
      sheet.getRange(rowIdx, 3).setValue(dk); 
      sheet.getRange(rowIdx, 4).setValue(da);
      sheet.getRange(rowIdx, 5).setValue("Hợp lệ");
      return "✏️ Đã cập nhật.";
    } else {
      sheet.appendRow([new Date(target), hinhThucSanitized, dk, da, "Hợp lệ"]);
      return "✅ Đã thêm.";
    }
  }
}

// ==========================================
// BÁO CÁO TỒN KHO
// ==========================================

/** Tên hiển thị đầy đủ cho mã hình thức nhập/xuất, dùng để nhóm báo cáo theo hình thức. */
function tenHienThiHinhThuc(code) {
  var map = {
    "TP": "TP - Dăm sản xuất",
    "TT": "TT - Xuất bán",
    "TC": "TC - Trung chuyển",
    "DC": "DC - Điều chuyển nội bộ",
    "XB": "XB - Xuất bán khác",
    "MUON_DH": "MUON_DH - Mượn Đại Hiệp",
    "MUON_QC": "MUON_QC - Mượn HAK_QN",
    "MUON_QS": "MUON_QS - Mượn CNHAK",
    "Khác": "Khác"
  };
  return map[code] || code || "(Không xác định)";
}

function layDanhSachTenNhaMay() {
  var arr = [];
  layDanhSachDanhMucKho().forEach(function(k) {
    if (arr.indexOf(k.tenNhaMay) === -1) arr.push(k.tenNhaMay);
  });
  return arr;
}

/** Chuyển map {hinhThuc: {mt, bdmt}} thành mảng đã sắp xếp giảm dần theo BDMT, có tên hiển thị. */
function _chuyenHinhThucThanhMang(mapHinhThuc) {
  var arr = [];
  for (var k in mapHinhThuc) {
    arr.push({ hinhThuc: k, ten: tenHienThiHinhThuc(k), mt: mapHinhThuc[k].mt, bdmt: mapHinhThuc[k].bdmt });
  }
  arr.sort(function(a, b) { return b.bdmt - a.bdmt; });
  return arr;
}

/**
 * Báo cáo tồn kho đầy đủ: Tồn đầu kỳ / Nhập-Xuất trong kỳ (theo hình thức) / Tồn cuối kỳ.
 * Tính cho cả MT (tươi) và BDMT (khô), kèm % độ khô bình quân của số tồn.
 * Hỗ trợ lọc theo khoảng ngày, theo kho, và theo nhà máy.
 *
 * @param {Object} [params]
 * @param {string} [params.tuNgay]     - Đầu kỳ báo cáo (yyyy-MM-dd). Mặc định: đầu tháng hiện tại.
 * @param {string} [params.denNgay]    - Cuối kỳ báo cáo (yyyy-MM-dd). Mặc định: hôm nay.
 * @param {string} [params.tenKho]     - Chỉ lọc 1 kho cụ thể. "Tất cả" hoặc bỏ trống = không lọc.
 * @param {string} [params.tenNhaMay]  - Chỉ lọc 1 nhà máy cụ thể. "Tất cả" hoặc bỏ trống = không lọc.
 */
function layBaoCaoTonKho(params) {
  kiemTraVaTaoTieuDeSheets();
  params = params || {};

  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var tuNgayStr = params.tuNgay || layNgay3ThangTruocStr();
  var denNgayStr = params.denNgay || todayStr;
  var tenKhoLoc = (params.tenKho && params.tenKho !== "Tất cả") ? params.tenKho : null;
  var tenNhaMayLoc = (params.tenNhaMay && params.tenNhaMay !== "Tất cả") ? params.tenNhaMay : null;

  var tuTime = new Date(tuNgayStr).setHours(0,0,0,0);
  var denTime = new Date(denNgayStr).setHours(23,59,59,999);

  function taoKhoiRong(tenKho, tenNhaMay) {
    return {
      tenKho: tenKho, tenNhaMay: tenNhaMay,
      tonDauMT: 0, tonDauBDMT: 0,
      nhapKyMT: 0, nhapKyBDMT: 0,
      xuatKyMT: 0, xuatKyBDMT: 0,
      nhapTheoHinhThuc: {}, xuatTheoHinhThuc: {}
    };
  }

  // Khởi tạo từ danh mục kho trước (đã lọc theo kho/nhà máy), để kho chưa có giao dịch vẫn hiện tồn = 0
  var map = {};
  layDanhSachDanhMucKho().forEach(function(k) {
    if (tenKhoLoc && k.tenKho !== tenKhoLoc) return;
    if (tenNhaMayLoc && k.tenNhaMay !== tenNhaMayLoc) return;
    map[k.tenKho] = taoKhoiRong(k.tenKho, k.tenNhaMay);
  });

  function layOrTaoKho(tenKho) {
    if (map[tenKho]) return map[tenKho];
    if (tenKhoLoc && tenKho !== tenKhoLoc) return null;
    if (tenNhaMayLoc) return null; // kho ngoài danh mục: không xác định được nhà máy nên loại khi đang lọc theo nhà máy
    map[tenKho] = taoKhoiRong(tenKho, "(Ngoài danh mục)");
    return map[tenKho];
  }

  var ss = SpreadsheetApp.openById(KHODAM_CONFIG.SPREADSHEET_ID);
  var sheetGD = ss.getSheetByName(KHODAM_CONFIG.SHEET_GIAODICH);
  if (sheetGD && sheetGD.getLastRow() > 1) {
    var data = sheetGD.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][11]).trim() === "Đã hủy") continue;
      var dt = new Date(data[i][1]).getTime();
      if (dt > denTime) continue; // phát sinh sau kỳ báo cáo thì bỏ qua hoàn toàn

      var loaiPhieu = String(data[i][2]).trim();
      var hinhThuc = String(data[i][3]).trim();
      var kX = String(data[i][5]).trim();
      var kN = String(data[i][6]).trim();
      var mt = parseFloat(data[i][7]) || 0;
      var bdmt = parseFloat(data[i][10]) || 0;
      var laTonDau = dt < tuTime;
      var laTrongKy = dt >= tuTime && dt <= denTime;

      // QUAN TRỌNG: dùng cột "Loại" (NHẬP/XUẤT) làm căn cứ chính, không chỉ dựa vào cột nào có giá trị.
      // - Loại = NHẬP: chỉ tính là NHẬP cho Kho Nhập (kN). Bỏ qua kX kể cả khi có giá trị (phòng dữ liệu chỉnh tay sai).
      // - Loại = XUẤT: luôn tính là XUẤT cho Kho Xuất (kX). Nếu là Trung chuyển (TC) và có Kho Nhập hợp lệ
      //   thì CỘNG THÊM phần NHẬP cho kho đó (vì trung chuyển vừa là xuất khỏi kho này vừa là nhập vào kho kia).
      if (loaiPhieu === "NHẬP") {
        if (kN && kN !== "Không có") {
          var khoNhanNhap = layOrTaoKho(kN);
          if (khoNhanNhap) {
            if (laTonDau) { khoNhanNhap.tonDauMT += mt; khoNhanNhap.tonDauBDMT += bdmt; }
            if (laTrongKy) {
              khoNhanNhap.nhapKyMT += mt; khoNhanNhap.nhapKyBDMT += bdmt;
              if (!khoNhanNhap.nhapTheoHinhThuc[hinhThuc]) khoNhanNhap.nhapTheoHinhThuc[hinhThuc] = { mt: 0, bdmt: 0 };
              khoNhanNhap.nhapTheoHinhThuc[hinhThuc].mt += mt;
              khoNhanNhap.nhapTheoHinhThuc[hinhThuc].bdmt += bdmt;
            }
          }
        }
      } else if (loaiPhieu === "XUẤT") {
        if (kX && kX !== "Không có") {
          var khoGuiXuat = layOrTaoKho(kX);
          if (khoGuiXuat) {
            if (laTonDau) { khoGuiXuat.tonDauMT -= mt; khoGuiXuat.tonDauBDMT -= bdmt; }
            if (laTrongKy) {
              khoGuiXuat.xuatKyMT += mt; khoGuiXuat.xuatKyBDMT += bdmt;
              if (!khoGuiXuat.xuatTheoHinhThuc[hinhThuc]) khoGuiXuat.xuatTheoHinhThuc[hinhThuc] = { mt: 0, bdmt: 0 };
              khoGuiXuat.xuatTheoHinhThuc[hinhThuc].mt += mt;
              khoGuiXuat.xuatTheoHinhThuc[hinhThuc].bdmt += bdmt;
            }
          }
        }
        if (hinhThuc === "TC" && kN && kN !== "Không có") {
          var khoNhanTC = layOrTaoKho(kN);
          if (khoNhanTC) {
            if (laTonDau) { khoNhanTC.tonDauMT += mt; khoNhanTC.tonDauBDMT += bdmt; }
            if (laTrongKy) {
              khoNhanTC.nhapKyMT += mt; khoNhanTC.nhapKyBDMT += bdmt;
              if (!khoNhanTC.nhapTheoHinhThuc[hinhThuc]) khoNhanTC.nhapTheoHinhThuc[hinhThuc] = { mt: 0, bdmt: 0 };
              khoNhanTC.nhapTheoHinhThuc[hinhThuc].mt += mt;
              khoNhanTC.nhapTheoHinhThuc[hinhThuc].bdmt += bdmt;
            }
          }
        }
      }
    }
  }

  var ketQua = [];
  for (var key in map) {
    var it = map[key];
    var tonCuoiMT = it.tonDauMT + it.nhapKyMT - it.xuatKyMT;
    var tonCuoiBDMT = it.tonDauBDMT + it.nhapKyBDMT - it.xuatKyBDMT;
    ketQua.push({
      tenKho: it.tenKho,
      tenNhaMay: it.tenNhaMay,
      tonDauMT: it.tonDauMT,
      tonDauBDMT: it.tonDauBDMT,
      tonDauDoKho: it.tonDauMT > 0.0001 ? (it.tonDauBDMT / it.tonDauMT) : null,
      nhapKyMT: it.nhapKyMT,
      nhapKyBDMT: it.nhapKyBDMT,
      xuatKyMT: it.xuatKyMT,
      xuatKyBDMT: it.xuatKyBDMT,
      tonCuoiMT: tonCuoiMT,
      tonCuoiBDMT: tonCuoiBDMT,
      tonCuoiDoKho: tonCuoiMT > 0.0001 ? (tonCuoiBDMT / tonCuoiMT) : null,
      nhapTheoHinhThuc: _chuyenHinhThucThanhMang(it.nhapTheoHinhThuc),
      xuatTheoHinhThuc: _chuyenHinhThucThanhMang(it.xuatTheoHinhThuc)
    });
  }
  ketQua.sort(function(a, b) {
    if (a.tenNhaMay !== b.tenNhaMay) return a.tenNhaMay.localeCompare(b.tenNhaMay, 'vi');
    return a.tenKho.localeCompare(b.tenKho, 'vi');
  });

  return {
    tuNgay: tuNgayStr,
    denNgay: denNgayStr,
    tenKhoLoc: tenKhoLoc || "Tất cả",
    tenNhaMayLoc: tenNhaMayLoc || "Tất cả",
    khoList: layDanhSachTenKho(),
    nhaMayList: layDanhSachTenNhaMay(),
    chiTiet: ketQua
  };
}

// ==========================================
// BÁO CÁO THEO KỲ VÉT BÃI
// ==========================================

/**
 * Báo cáo tổng hợp Nhập/Xuất theo từng Kỳ Vét Bãi (cột "Đợt vét bãi" trong DATA_GIAODICH).
 * Mặc định khoảng thời gian: 3 tháng gần nhất. Có thể lọc theo 1 kho cụ thể.
 * @param {Object} [params]
 * @param {string} [params.tuNgay]
 * @param {string} [params.denNgay]
 * @param {string} [params.tenKho] - "Tất cả" hoặc bỏ trống = không lọc theo kho.
 */
function layBaoCaoTheoKyVetBai(params) {
  kiemTraVaTaoTieuDeSheets();
  params = params || {};
  var tuNgayStr = params.tuNgay || layNgay3ThangTruocStr();
  var denNgayStr = params.denNgay || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var tenKhoLoc = (params.tenKho && params.tenKho !== "Tất cả") ? params.tenKho : null;

  var tuTime = new Date(tuNgayStr).setHours(0, 0, 0, 0);
  var denTime = new Date(denNgayStr).setHours(23, 59, 59, 999);

  var map = {};
  function layOrTaoKy(tenKy) {
    if (!map[tenKy]) map[tenKy] = { tenKy: tenKy, nhapMT: 0, nhapBDMT: 0, xuatMT: 0, xuatBDMT: 0, soPhieuNhap: 0, soPhieuXuat: 0 };
    return map[tenKy];
  }

  var ss = SpreadsheetApp.openById(KHODAM_CONFIG.SPREADSHEET_ID);
  var sheetGD = ss.getSheetByName(KHODAM_CONFIG.SHEET_GIAODICH);
  if (sheetGD && sheetGD.getLastRow() > 1) {
    var data = sheetGD.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][11]).trim() === "Đã hủy") continue;
      var dt = new Date(data[i][1]).getTime();
      if (dt < tuTime || dt > denTime) continue;

      var loaiVal = String(data[i][2]).trim();
      var kX = String(data[i][5]).trim();
      var kN = String(data[i][6]).trim();
      var tenKy = String(data[i][4]).trim() || "Mặc định";
      var mt = parseFloat(data[i][7]) || 0;
      var bdmt = parseFloat(data[i][10]) || 0;

      if (loaiVal === "NHẬP") {
        if (tenKhoLoc && kN !== tenKhoLoc) continue;
        var kyN = layOrTaoKy(tenKy);
        kyN.nhapMT += mt; kyN.nhapBDMT += bdmt; kyN.soPhieuNhap++;
      } else if (loaiVal === "XUẤT") {
        if (tenKhoLoc && kX !== tenKhoLoc) continue;
        var kyX = layOrTaoKy(tenKy);
        kyX.xuatMT += mt; kyX.xuatBDMT += bdmt; kyX.soPhieuXuat++;
      }
    }
  }

  // Sắp xếp theo đúng thứ tự thời gian thực tế của kỳ vét bãi; kỳ không xác định được xếp cuối
  var arrKyCache = taiDanhSachKyVetBaiCache();
  var thuTu = {};
  arrKyCache.forEach(function(k, idx) { thuTu[k.tenKy] = idx; });

  var ketQua = [];
  for (var key in map) {
    var it = map[key];
    ketQua.push({
      tenKy: it.tenKy,
      nhapMT: it.nhapMT, nhapBDMT: it.nhapBDMT,
      xuatMT: it.xuatMT, xuatBDMT: it.xuatBDMT,
      chenhLechMT: it.nhapMT - it.xuatMT,
      chenhLechBDMT: it.nhapBDMT - it.xuatBDMT,
      soPhieuNhap: it.soPhieuNhap, soPhieuXuat: it.soPhieuXuat
    });
  }
  ketQua.sort(function(a, b) {
    var ia = (thuTu[a.tenKy] !== undefined) ? thuTu[a.tenKy] : 9999;
    var ib = (thuTu[b.tenKy] !== undefined) ? thuTu[b.tenKy] : 9999;
    return ia - ib;
  });

  return {
    tuNgay: tuNgayStr,
    denNgay: denNgayStr,
    tenKhoLoc: tenKhoLoc || "Tất cả",
    khoList: layDanhSachTenKho(),
    chiTiet: ketQua
  };
}

function processFormData(action, data) {
  kiemTraVaTaoTieuDeSheets();
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    var ketQua;
    if (action === "Danhmuckho") ketQua = xuLyDanhMucKho(data);
    else if (action === "Thongsokho") ketQua = xuLyKyVetBai(data);
    else if (action === "Nhapdokho") ketQua = xuLySuaXoaDoKho(data);
    else if (action === "Nhapkho" || action === "Xuatkho") ketQua = xuLySuaXoaGiaoDich(data);
    else if (action === "Hoanthanhdonhang") ketQua = xuLyHoanThanhDonHangXuatBan(data);
    else if (action === "HoanthanhTuDong") ketQua = KD_taoPhieuDieuChinhTuDong(data);
    else if (action === "Baocaotonkho") ketQua = layBaoCaoTonKho(data);
    else if (action === "BaocaoKyVetBai") ketQua = layBaoCaoTheoKyVetBai(data);
    // Ghi audit cho các hành động THAY ĐỔI dữ liệu (bỏ qua các hành động chỉ ĐỌC báo cáo)
    if (["Danhmuckho","Thongsokho","Nhapdokho","Nhapkho","Xuatkho","Hoanthanhdonhang","HoanthanhTuDong"].indexOf(action) !== -1) {
      var thanhCong = !(typeof ketQua === "string" && ketQua.indexOf("❌") === 0);
      logAudit_("KHODAM_" + action.toUpperCase(), thanhCong ? "OK" : "ERROR", typeof ketQua === "string" ? ketQua : JSON.stringify(data));
    }
    return ketQua;
  } catch (e) {
    logAudit_("KHODAM_" + String(action).toUpperCase(), "ERROR", e.toString());
    return "❌ Lỗi: " + e.toString();
  } finally {
    lock.releaseLock();
  }
}

/*********************************************************
 * PHẦN 7: XUẤT HÀNG - IMPORT PHIẾU CÂN XUẤT (NL_PC_XH) &
 *         NHẬP LIỆU ĐƠN HÀNG XUẤT BÁN (NL_DH_XB) (MENU 1)
 * - Spreadsheet RIÊNG: XUATHANG_CONFIG.SPREADSHEET_ID (xem Config.gs).
 * - Luồng Import giống hệt "Import phiếu cân" (PhieuCan_DN): Xem trước ->
 *   XUATHANG_CONFIG.SHEET_NLPCXH_DRAFT (sheet tạm để đối soát) -> chọn dòng ->
 *   Xác nhận -> ghi vào XUATHANG_CONFIG.SHEET_NLPCXH (chính thức).
 * - LƯU Ý: File import phải theo ĐÚNG cấu trúc 16 cột của XUATHANG_CONFIG.
 *   SHEET_NLPCXH (Số phiếu, Ngày giờ cân 1, Ngày giờ cân 2, Biển số 1, Cân lần
 *   1, Cân lần 2, KL Hàng (KG), Đơn vị vận chuyển, Tên tài xế, Khối lượng
 *   (Tấn), Ngày xuất, Số BKLS, Khối lượng (M3), NGƯỜI CÂN, SỐ TKHQ, TÀU XUẤT) -
 *   không còn đoán mò định dạng "kiểu cân Xuất" của trạm cân nữa. Có file mẫu
 *   tải sẵn qua hàm taiFileMauImport('XUATHANG') / nút "Tải file mẫu" trên
 *   giao diện, để người chuẩn bị dữ liệu biết chính xác cột nào cần điền.
 *   Đọc cột vẫn tìm THEO TÊN TIÊU ĐỀ (không theo vị trí cố định) để chấp
 *   nhận sai khác nhỏ về thứ tự cột, miễn tên tiêu đề đúng như trên.
 *********************************************************/

function XH_ss_() { return SpreadsheetApp.openById(XUATHANG_CONFIG.SPREADSHEET_ID); }

// Đảm bảo sheet NL_DH_XB có đủ cột "Kho xuất" (cột P, mới thêm theo yêu cầu) -
// sheet gốc người dùng cung cấp chỉ có 15 cột (A..O), gọi hàm này 1 lần trước
// khi ghi/đọc để tự bổ sung header cột 16 nếu còn thiếu, không cần chỉnh tay.
function XH_dambaoHeaderDonHang_() {
  const sheet = XH_ss_().getSheetByName(XUATHANG_CONFIG.SHEET_DHXB);
  if (!sheet) return;
  const hienTai = sheet.getRange(1, 16).getValue();
  if (!hienTai) sheet.getRange(1, 16).setValue("Kho xuất").setFontWeight("bold");
}

// Đảm bảo sheet NL_PC_XH có đúng cấu trúc cột MỚI: cột P đổi tên từ "TÀU XUẤT"
// (chưa từng có dữ liệu, đã xác nhận an toàn) thành "Kho xuất", và thêm cột Q
// mới "Kho nhập". Cột O (SỐ TKHQ) giữ NGUYÊN, không đụng tới. Idempotent - gọi
// nhiều lần không sao, chỉ đổi khi tiêu đề còn là "TÀU XUẤT"/rỗng.
function XH_dambaoHeaderPhieuCanXuat_() {
  const sheet = XH_ss_().getSheetByName(XUATHANG_CONFIG.SHEET_NLPCXH);
  if (!sheet) return;
  const headerP = String(sheet.getRange(1, 16).getValue() || "").trim();
  if (headerP === "" || headerP.toUpperCase() === "TÀU XUẤT") {
    sheet.getRange(1, 16).setValue("Kho xuất").setFontWeight("bold");
  }
  const headerQ = String(sheet.getRange(1, 17).getValue() || "").trim();
  if (headerQ === "") {
    sheet.getRange(1, 17).setValue("Kho nhập").setFontWeight("bold");
  }
}

// Tên tiêu đề CÓ THỂ gặp cho từng trường (ưu tiên khớp đúng tên NL_PC_XH trước)
const XH_HEADER_ALIASES_ = {
  soPhieu: ["Số phiếu"],
  ngayCan1: ["Ngày giờ cân 1"],
  ngayCan2: ["Ngày giờ cân 2"],
  bienSo: ["Biển số 1"],
  canLan1: ["Cân lần 1"],
  canLan2: ["Cân lần 2"],
  klHang: ["KL Hàng (KG)", "KL Hàng(KG)", "KL Hàng"],
  donViVanChuyen: ["Đơn vị vận chuyển"],
  tenTaiXe: ["Tên tài xế"],
  khoiLuongTan: ["Khối lượng (Tấn)", "Khối lượng(Tấn)"],
  ngayXuat: ["Ngày xuất"],
  soBKLS: ["Số BKLS"],
  khoiLuongM3: ["Khối lượng (M3)", "Khối lượng(M3)"],
  nguoiCan: ["NGƯỜI CÂN", "Người cân 1", "Người cân"],
  soTKHQ: ["SỐ TKHQ", "Số TKHQ"],
  khoXuat: ["Kho xuất", "KHO XUẤT"],
  khoNhap: ["Kho nhập", "KHO NHẬP"]
};

function XH_timCotTheoTen_(headerRow) {
  const map = {};
  headerRow.forEach((cell, idx) => {
    const ten = String(cell || "").trim();
    if (!ten) return;
    for (const key in XH_HEADER_ALIASES_) {
      if (map[key] !== undefined) continue; // đã tìm thấy rồi thì thôi, ưu tiên cột xuất hiện trước
      if (XH_HEADER_ALIASES_[key].some(alias => alias.toLowerCase() === ten.toLowerCase())) {
        map[key] = idx;
      }
    }
  });
  return map;
}

function XH_step1_PreviewDraft(fileData, khoXuatMacDinh, khoNhapMacDinh) {
  try {
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_INPUT);
    if (fileData && fileData.base64) {
      folder.createFile(Utilities.newBlob(Utilities.base64Decode(fileData.base64), fileData.mimeType, fileData.name));
    }

    const ss = XH_ss_();
    const dataSheet = ss.getSheetByName(XUATHANG_CONFIG.SHEET_NLPCXH);
    const lastRow = dataSheet.getLastRow();
    const duplicateMap = new Map();
    if (lastRow > 0) {
      const existingData = dataSheet.getRange(1, 1, lastRow, 1).getValues();
      existingData.forEach(row => {
        const key = String(row[0] || "").trim();
        if (key) duplicateMap.set(key, true);
      });
    }

    const files = folder.getFiles();
    let previewRows = [];
    const seenInThisBatch = new Set();

    while (files.hasNext()) {
      const file = files.next();
      if (!file.getName().match(/\.xls[x]?$/i)) continue;

      const tempFile = Drive.Files.insert({ title: "TMP_XH_" + file.getName() }, file.getBlob(), { convert: true });
      try {
        const values = SpreadsheetApp.openById(tempFile.id).getSheets()[0].getDataRange().getValues();
        let hIdx = values.findIndex(r => r.some(c => String(c).toLowerCase().includes("số phiếu")));
        if (hIdx === -1) continue;

        const cotMap = XH_timCotTheoTen_(values[hIdx]);
        if (cotMap.soPhieu === undefined || cotMap.ngayCan1 === undefined) continue; // thiếu cột lõi, bỏ qua file này

        const rowsToProcess = values.slice(hIdx + 1);
        for (let r of rowsToProcess) {
          const spRaw = String(r[cotMap.soPhieu] || "").trim();
          if (!spRaw || spRaw.toLowerCase().includes("ngày") || spRaw.toLowerCase().includes("tổng") || spRaw.length > 20) continue;

          const klHangRaw = cotMap.klHang !== undefined ? r[cotMap.klHang] : null;
          let klHang = parseFloat(String(klHangRaw || "").replace(/[^0-9.]/g, ""));
          if (!klHangRaw || isNaN(klHang) || klHang === 0) continue;

          let dateC = toDateObj(r[cotMap.ngayCan1]);
          let dateD = cotMap.ngayCan2 !== undefined ? toDateObj(r[cotMap.ngayCan2]) : null;

          let now = new Date();
          let nam = dateC ? dateC.getFullYear() : now.getFullYear();
          let currentMaChungTu = spRaw + "/" + nam + "/XK";

          let isError = false; let errorMsg = "";
          if (!dateC || isNaN(dateC.getTime())) { isError = true; errorMsg += "Lỗi Ngày Cân 1. "; }
          if (cotMap.ngayCan2 !== undefined && (!dateD || isNaN(dateD.getTime()))) { isError = true; errorMsg += "Lỗi Ngày Cân 2. "; }

          if (seenInThisBatch.has(currentMaChungTu)) {
            isError = true;
            errorMsg += "Trùng Số Chứng Từ ngay trong dữ liệu đang nạp (" + currentMaChungTu + "). ";
          } else {
            seenInThisBatch.add(currentMaChungTu);
          }

          let typeImport = duplicateMap.has(currentMaChungTu) ? "Bỏ qua (Đã tồn tại)" : "Mới";

          let ngayXuat = null;
          if (cotMap.ngayXuat !== undefined) ngayXuat = toDateObj(r[cotMap.ngayXuat]);

          // Kho xuất/Kho nhập: ưu tiên lấy TỪ FILE nếu người chuẩn bị dữ liệu đã
          // điền theo đúng cột (Kho xuất/Kho nhập); nếu file không có 2 cột này,
          // dùng giá trị mặc định áp dụng cho CẢ LÔ do người dùng chọn trước khi tải lên.
          const khoXuatRow = cotMap.khoXuat !== undefined ? String(r[cotMap.khoXuat] || "").trim() : "";
          const khoNhapRow = cotMap.khoNhap !== undefined ? String(r[cotMap.khoNhap] || "").trim() : "";

          previewRows.push({
            isError: isError, errorMsg: errorMsg.trim(), typeImport: typeImport, uniqueKey: currentMaChungTu,
            soPhieu: spRaw,
            ngayCan1: dateC ? Utilities.formatDate(dateC, "GMT+7", "dd/MM/yyyy") : "Lỗi định dạng ngày",
            gioCan1: dateC ? Utilities.formatDate(dateC, "GMT+7", "HH:mm:ss") : "",
            ngayCan2: dateD ? Utilities.formatDate(dateD, "GMT+7", "dd/MM/yyyy") : (cotMap.ngayCan2 !== undefined ? "Lỗi định dạng ngày" : ""),
            gioCan2: dateD ? Utilities.formatDate(dateD, "GMT+7", "HH:mm:ss") : "",
            bienSo: cotMap.bienSo !== undefined ? String(r[cotMap.bienSo] || "").trim() : "",
            canLan1: cotMap.canLan1 !== undefined ? (parseFloat(r[cotMap.canLan1]) || 0) : 0,
            canLan2: cotMap.canLan2 !== undefined ? (parseFloat(r[cotMap.canLan2]) || 0) : 0,
            klHang: klHang,
            donViVanChuyen: cotMap.donViVanChuyen !== undefined ? String(r[cotMap.donViVanChuyen] || "").trim() : "",
            tenTaiXe: cotMap.tenTaiXe !== undefined ? String(r[cotMap.tenTaiXe] || "").trim() : "",
            nguoiCan: cotMap.nguoiCan !== undefined ? String(r[cotMap.nguoiCan] || "").trim() : "",
            // Các cột đặc thù NL_PC_XH - có gì lấy nấy, không có thì để trống (điền bổ sung sau khi đối chiếu đơn hàng)
            khoiLuongTan: cotMap.khoiLuongTan !== undefined ? (parseFloat(r[cotMap.khoiLuongTan]) || (klHang / 1000)) : (klHang / 1000),
            soBKLS: cotMap.soBKLS !== undefined ? String(r[cotMap.soBKLS] || "").trim() : "",
            khoiLuongM3: cotMap.khoiLuongM3 !== undefined ? (parseFloat(r[cotMap.khoiLuongM3]) || "") : "",
            soTKHQ: cotMap.soTKHQ !== undefined ? String(r[cotMap.soTKHQ] || "").trim() : "",
            khoXuat: khoXuatRow || String(khoXuatMacDinh || "").trim(),
            khoNhap: khoNhapRow || String(khoNhapMacDinh || "").trim(),
            rawDateC: dateC ? dateC.toISOString() : "", rawDateD: dateD ? dateD.toISOString() : "",
            rawNgayXuat: ngayXuat ? ngayXuat.toISOString() : ""
          });
        }
      } finally {
        Drive.Files.remove(tempFile.id);
      }
    }

    // Sheet Draft để đối soát (ghi cùng lúc, tách biệt hoàn toàn với PhieuCan_DN)
    let draftUrl = "";
    if (previewRows.length > 0) {
      const draftSheet = ss.getSheetByName(XUATHANG_CONFIG.SHEET_NLPCXH_DRAFT);
      if (draftSheet) {
        if (draftSheet.getLastRow() > 1) draftSheet.getRange(2, 1, draftSheet.getLastRow() - 1, 16).clearContent();
        const draftHeaders = ["Trạng Thái", "Mã Chứng Từ", "Số Phiếu", "Số Xe", "Ngày Cân 1", "Giờ Cân 1", "Ngày Cân 2", "Giờ Cân 2", "Cân Lần 1", "Cân Lần 2", "KL Hàng (Kg)", "Đơn Vị Vận Chuyển", "Tên Tài Xế", "Số TKHQ", "Kho Xuất", "Kho Nhập"];
        draftSheet.getRange(1, 1, 1, draftHeaders.length).setValues([draftHeaders]).setFontWeight("bold").setBackground("#cfe2ff");
        draftSheet.getRange(2, 2, previewRows.length, 7).setNumberFormat("@"); // ép Text cột B..H (Mã CT..Giờ Cân 2), tránh Sheets tự đoán lại theo Locale
        const draftValues = previewRows.map(row => [
          row.isError ? "❌ " + row.errorMsg : "✔️ " + row.typeImport, row.uniqueKey, row.soPhieu, row.bienSo,
          row.ngayCan1, row.gioCan1, row.ngayCan2, row.gioCan2, row.canLan1, row.canLan2, row.klHang, row.donViVanChuyen, row.tenTaiXe,
          row.soTKHQ, row.khoXuat, row.khoNhap
        ]);
        draftSheet.getRange(2, 1, draftValues.length, draftHeaders.length).setValues(draftValues);
        draftSheet.getRange(2, 9, draftValues.length, 3).setNumberFormat("#,##0");
        draftUrl = ss.getUrl() + "#gid=" + draftSheet.getSheetId();
      }
    }

    return { status: "success", data: previewRows, draftUrl: draftUrl };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

function XH_step1_ConfirmImport(confirmedDataList) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
  } catch (e) {
    return { status: "error", message: "Hệ thống đang bận xử lý một yêu cầu khác, vui lòng thử lại sau ít giây." };
  }
  try {
    XH_dambaoHeaderPhieuCanXuat_();
    const ss = XH_ss_();
    const dataSheet = ss.getSheetByName(XUATHANG_CONFIG.SHEET_NLPCXH);
    const lastRow = dataSheet.getLastRow();

    const existingKeys = new Set();
    if (lastRow > 0) {
      dataSheet.getRange(1, 1, lastRow, 1).getValues().forEach(row => {
        const k = String(row[0] || "").trim();
        if (k) existingKeys.add(k);
      });
    }

    let countNew = 0, countSkip = 0; const batchNew = []; const seenThisBatch = new Set();

    for (let item of confirmedDataList) {
      if (item.isError) continue;
      if (existingKeys.has(item.uniqueKey) || seenThisBatch.has(item.uniqueKey)) { countSkip++; continue; }
      seenThisBatch.add(item.uniqueKey);

      const dateC = item.rawDateC ? new Date(item.rawDateC) : null;
      const dateD = item.rawDateD ? new Date(item.rawDateD) : null;
      const ngayXuat = item.rawNgayXuat ? new Date(item.rawNgayXuat) : null;
      const ngayCan1Str = dateC ? Utilities.formatDate(dateC, "GMT+7", "dd/MM/yyyy HH:mm:ss") : "";
      const ngayCan2Str = dateD ? Utilities.formatDate(dateD, "GMT+7", "dd/MM/yyyy HH:mm:ss") : "";
      const klTan = parseFloat(item.khoiLuongTan) || ((parseFloat(item.klHang) || 0) / 1000);

      // Khớp đúng 17 cột của NL_PC_XH (A..Q): cột O = Số TKHQ (không đổi), cột P
      // = Kho xuất (thay cho "TÀU XUẤT" cũ - đã xác nhận chưa từng có dữ liệu
      // nên an toàn để tái sử dụng), cột Q = Kho nhập (cột MỚI thêm).
      batchNew.push([
        item.uniqueKey, ngayCan1Str, ngayCan2Str, item.bienSo || "",
        parseFloat(item.canLan1) || 0, parseFloat(item.canLan2) || 0, parseFloat(item.klHang) || 0,
        item.donViVanChuyen || "", item.tenTaiXe || "", klTan,
        ngayXuat || "", item.soBKLS || "", item.khoiLuongM3 || "", item.nguoiCan || "",
        item.soTKHQ || "", item.khoXuat || "", item.khoNhap || ""
      ]);
      countNew++;
    }

    if (batchNew.length > 0) dataSheet.getRange(dataSheet.getLastRow() + 1, 1, batchNew.length, 17).setValues(batchNew);

    const folder = DriveApp.getFolderById(CONFIG.FOLDER_INPUT); const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next(); if (!file.getName().match(/\.xls[x]?$/i)) continue;
      const meta = Drive.Files.get(file.getId(), { fields: "parents" });
      if (meta.parents && meta.parents.length > 0) {
        Drive.Files.update({}, file.getId(), null, { addParents: CONFIG.FOLDER_DONE, removeParents: meta.parents[0].id });
      }
    }

    const finalMsg = `Mới: ${countNew}, Bỏ qua (đã tồn tại/trùng): ${countSkip}`;
    logAudit_("IMPORT_XUATHANG", "OK", finalMsg);
    return { status: "success", message: finalMsg };
  } catch (e) {
    logAudit_("IMPORT_XUATHANG", "ERROR", e.toString());
    return { status: "error", message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Nhập liệu Đơn Hàng Xuất Bán (NL_DH_XB) ---------- */

// Danh sách Kho xuất - tái sử dụng danh mục kho của module Kho Dăm
function XH_getKhoXuatList() {
  try { return { status: "success", data: layDanhSachTenKho() }; }
  catch (e) { return { status: "error", message: e.toString() }; }
}

// Độ khô nhà máy = độ khô TRUNG BÌNH của "hàng nhập sản xuất" (sheet Nhapdokho
// của module Kho Dăm) trong khoảng [tuNgay, denNgay]. Chỉ tính các dòng "Hợp lệ".
function XH_tinhDoKhoNhaMay(tuNgay, denNgay) {
  try {
    if (!tuNgay || !denNgay) return { status: "error", message: "Vui lòng chọn đủ Từ ngày và Đến ngày." };
    const ss = SpreadsheetApp.openById(KHODAM_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(KHODAM_CONFIG.SHEET_NHAPDOKHO);
    if (!sheet || sheet.getLastRow() <= 1) return { status: "success", doKho: null, soNgay: 0, message: "Chưa có dữ liệu độ khô." };

    const data = sheet.getDataRange().getValues();
    const tuDate = new Date(tuNgay + "T00:00:00");
    const denDate = new Date(denNgay + "T23:59:59");
    let tong = 0, dem = 0;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][4] || "").trim() === "Đã hủy") continue;
      const ngayRow = data[i][0];
      if (!(ngayRow instanceof Date) || isNaN(ngayRow.getTime())) continue;
      if (ngayRow < tuDate || ngayRow > denDate) continue;
      const dk = parseFloat(data[i][2]);
      if (isNaN(dk)) continue;
      tong += dk; dem++;
    }

    if (dem === 0) return { status: "success", doKho: null, soNgay: 0, message: "Không có dữ liệu độ khô trong khoảng ngày đã chọn." };
    return { status: "success", doKho: tong / dem, soNgay: dem, message: "Trung bình " + dem + " ngày có ghi nhận độ khô." };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

// payload: {ngayDonHang, soTKHQ, tau, khachHang, diaChiKH, tenHangHoa, donGiaUSD,
//           klMT, klBDMT, khoXuat, tuNgay, denNgay, doKhoNhaMay, loaiXe}
function XH_saveDonHang(payload) {
  XH_dambaoHeaderDonHang_();
  const lock = LockService.getScriptLock();
  try { lock.waitLock(CONFIG.LOCK_TIMEOUT_MS); } catch (e) {
    return { status: "error", message: "Hệ thống đang bận, vui lòng thử lại." };
  }
  try {
    payload = payload || {};
    if (!payload.ngayDonHang) return { status: "error", message: "Vui lòng chọn Ngày đơn hàng." };
    if (!String(payload.khachHang || "").trim()) return { status: "error", message: "Vui lòng nhập Khách hàng." };
    const klMT = parseFloat(payload.klMT) || 0;
    if (klMT <= 0) return { status: "error", message: "KL_MT phải lớn hơn 0." };
    // FIX: KL_BDMT không còn nhập tay - tự tính = KL_MT × Độ khô (theo đúng công
    // thức nghiệp vụ KL_MT × ĐỘ KHÔ = KL_BDMT). "Độ khô" ở đây là độ khô THỰC TẾ
    // áp dụng cho riêng lô hàng này (payload.doKho, dạng %), KHÁC với "Độ khô
    // nhà máy" (chỉ là số tham chiếu trung bình, không dùng để tính KL_BDMT).
    let doKho = parseFloat(payload.doKho) || 0;
    if (doKho > 1) doKho = doKho / 100; // cho phép nhập 39.45 hoặc 0.3945 đều ra đúng tỷ lệ
    if (doKho <= 0) return { status: "error", message: "Vui lòng nhập Độ khô > 0 để tính KL_BDMT." };
    const klBDMT = klMT * doKho;
    const doKhoTinh = doKho;

    const sheet = XH_ss_().getSheetByName(XUATHANG_CONFIG.SHEET_DHXB);
    const lastRow = sheet.getLastRow();
    const stt = lastRow; // header ở dòng 1, dữ liệu bắt đầu dòng 2 -> STT = lastRow (số dòng dữ liệu hiện có + 1, vì lastRow đang TÍNH CẢ header)

    const loaiXe = String(payload.loaiXe || "Y").toUpperCase() === "N" ? "N" : "Y";

    sheet.appendRow([
      stt, new Date(payload.ngayDonHang), String(payload.soTKHQ || "").trim(), String(payload.tau || "").trim(),
      String(payload.khachHang || "").trim(), String(payload.diaChiKH || "").trim(), String(payload.tenHangHoa || "").trim(),
      parseFloat(payload.donGiaUSD) || 0, klMT, klBDMT, doKhoTinh,
      parseFloat(payload.doKhoNhaMay) || 0, payload.tuNgay || "", payload.denNgay || "", loaiXe,
      String(payload.khoXuat || "").trim()
    ]);

    logAudit_("XUATHANG_DONHANG", "OK", "Thêm đơn hàng xuất bán: " + payload.khachHang + " - " + payload.tau);
    return { status: "success", message: "✅ Đã lưu đơn hàng xuất bán." };
  } catch (e) {
    logAudit_("XUATHANG_DONHANG", "ERROR", e.toString());
    return { status: "error", message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function XH_getDonHangList() {
  try {
    XH_dambaoHeaderDonHang_();
    const sheet = XH_ss_().getSheetByName(XUATHANG_CONFIG.SHEET_DHXB);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: "success", data: [] };
    const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    // FIX: định dạng đúng NL_Từ ngày/NL_Đến ngày dù ô đang lưu dạng Date object
    // thật (Sheets có thể tự auto-detect chuỗi "yyyy-MM-dd" thành Date) hay vẫn
    // là chuỗi text thuần - trước đây 2 cột này được lấy về đúng nhưng KHÔNG
    // được hiển thị ở bảng danh sách, gây cảm giác "biến mất".
    const layNgayHienThi = v => (v instanceof Date && !isNaN(v.getTime())) ? Utilities.formatDate(v, "GMT+7", "dd/MM/yyyy") : String(v || "").trim();
    const result = data.map((r, idx) => ({
      rowIndex: idx + 2, // dòng thật trên sheet, dùng để Sửa/Xóa chính xác
      stt: r[0],
      ngayDonHang: (r[1] instanceof Date) ? Utilities.formatDate(r[1], "GMT+7", "dd/MM/yyyy") : "",
      soTKHQ: r[2], tau: r[3], khachHang: r[4], diaChiKH: r[5], tenHangHoa: r[6],
      donGiaUSD: parseFloat(r[7]) || 0, klMT: parseFloat(r[8]) || 0, klBDMT: parseFloat(r[9]) || 0,
      doKho: parseFloat(r[10]) || 0, doKhoNhaMay: parseFloat(r[11]) || 0,
      tuNgay: layNgayHienThi(r[12]), denNgay: layNgayHienThi(r[13]), loaiXe: r[14], khoXuat: r[15]
    })).filter(r => r.khachHang).reverse();
    return { status: "success", data: result };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

// Lấy chi tiết 1 đơn hàng theo rowIndex (dòng thật trên sheet) để nạp vào form Sửa
function XH_getDonHangByRow(rowIndex) {
  try {
    const sheet = XH_ss_().getSheetByName(XUATHANG_CONFIG.SHEET_DHXB);
    const r = parseInt(rowIndex, 10);
    if (isNaN(r) || r < 2 || r > sheet.getLastRow()) return { status: "error", message: "Không tìm thấy đơn hàng." };
    const row = sheet.getRange(r, 1, 1, 16).getValues()[0];
    const layNgayInput = v => (v instanceof Date && !isNaN(v.getTime())) ? Utilities.formatDate(v, "GMT+7", "yyyy-MM-dd") : String(v || "").trim();
    return {
      status: "success", rowIndex: r,
      ngayDonHang: (row[1] instanceof Date) ? Utilities.formatDate(row[1], "GMT+7", "yyyy-MM-dd") : "",
      soTKHQ: row[2], tau: row[3], khachHang: row[4], diaChiKH: row[5], tenHangHoa: row[6],
      donGiaUSD: parseFloat(row[7]) || 0, klMT: parseFloat(row[8]) || 0,
      doKho: (parseFloat(row[10]) || 0) * 100, // trả về dạng % cho khớp ô nhập
      doKhoNhaMay: parseFloat(row[11]) || 0,
      tuNgay: layNgayInput(row[12]), denNgay: layNgayInput(row[13]),
      loaiXe: row[14] || "Y", khoXuat: row[15] || ""
    };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

// Cập nhật 1 đơn hàng đã có, theo đúng dòng thật (rowIndex) - ghi đè toàn bộ dữ liệu
function XH_updateDonHang(rowIndex, payload) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(CONFIG.LOCK_TIMEOUT_MS); } catch (e) {
    return { status: "error", message: "Hệ thống đang bận, vui lòng thử lại." };
  }
  try {
    payload = payload || {};
    const r = parseInt(rowIndex, 10);
    const sheet = XH_ss_().getSheetByName(XUATHANG_CONFIG.SHEET_DHXB);
    if (isNaN(r) || r < 2 || r > sheet.getLastRow()) return { status: "error", message: "Không tìm thấy đơn hàng để sửa (có thể đã bị xóa)." };
    if (!payload.ngayDonHang) return { status: "error", message: "Vui lòng chọn Ngày đơn hàng." };
    if (!String(payload.khachHang || "").trim()) return { status: "error", message: "Vui lòng nhập Khách hàng." };
    const klMT = parseFloat(payload.klMT) || 0;
    if (klMT <= 0) return { status: "error", message: "KL_MT phải lớn hơn 0." };
    let doKho = parseFloat(payload.doKho) || 0;
    if (doKho > 1) doKho = doKho / 100;
    if (doKho <= 0) return { status: "error", message: "Vui lòng nhập Độ khô > 0 để tính KL_BDMT." };
    const klBDMT = klMT * doKho;
    const loaiXe = String(payload.loaiXe || "Y").toUpperCase() === "N" ? "N" : "Y";
    const sttCu = sheet.getRange(r, 1).getValue();

    sheet.getRange(r, 1, 1, 16).setValues([[
      sttCu, new Date(payload.ngayDonHang), String(payload.soTKHQ || "").trim(), String(payload.tau || "").trim(),
      String(payload.khachHang || "").trim(), String(payload.diaChiKH || "").trim(), String(payload.tenHangHoa || "").trim(),
      parseFloat(payload.donGiaUSD) || 0, klMT, klBDMT, doKho,
      parseFloat(payload.doKhoNhaMay) || 0, payload.tuNgay || "", payload.denNgay || "", loaiXe,
      String(payload.khoXuat || "").trim()
    ]]);

    logAudit_("XUATHANG_DONHANG", "OK", "Sửa đơn hàng xuất bán dòng " + r + ": " + payload.khachHang);
    return { status: "success", message: "✏️ Đã cập nhật đơn hàng." };
  } catch (e) {
    logAudit_("XUATHANG_DONHANG", "ERROR", e.toString());
    return { status: "error", message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function XH_deleteDonHang(rowIndex) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(CONFIG.LOCK_TIMEOUT_MS); } catch (e) {
    return { status: "error", message: "Hệ thống đang bận, vui lòng thử lại." };
  }
  try {
    const r = parseInt(rowIndex, 10);
    const sheet = XH_ss_().getSheetByName(XUATHANG_CONFIG.SHEET_DHXB);
    if (isNaN(r) || r < 2 || r > sheet.getLastRow()) return { status: "error", message: "Không tìm thấy đơn hàng để xóa (có thể đã bị xóa trước đó)." };
    sheet.deleteRow(r);
    logAudit_("XUATHANG_DONHANG", "OK", "Đã xóa đơn hàng xuất bán dòng " + r);
    return { status: "success", message: "🗑️ Đã xóa đơn hàng." };
  } catch (e) {
    logAudit_("XUATHANG_DONHANG", "ERROR", e.toString());
    return { status: "error", message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/*********************************************************
 * PHẦN 8: BÁO CÁO TỔNG HỢP XUẤT KHO (MENU 5)
 * - Song song với "Báo cáo tổng hợp nhập kho": gồm "Báo cáo xuất qua cân"
 *   (đọc NL_PC_XH, lọc + phân trang 20 dòng/trang) và "Báo cáo xuất kết xuất
 *   Misa" (đọc NL_DH_XB - đơn hàng xuất bán, tính thành tiền theo KL_BDMT x
 *   Đơn giá USD, tương tự cách "Báo cáo nhập kết xuất Misa" phục vụ hạch toán).
 *********************************************************/

// Helper dùng chung: lọc + sắp xếp toàn bộ NL_PC_XH theo bộ lọc, KHÔNG phân
// trang - dùng cả cho hiển thị (phân trang ở lớp gọi) lẫn xuất Excel/PDF (toàn bộ).
function XH_locBaoCaoXuatQuaCan_(filters) {
  const sheet = XH_ss_().getSheetByName(XUATHANG_CONFIG.SHEET_NLPCXH);
  const lastRow = sheet.getLastRow();
  const ketQua = { matched: [], dsDonViVanChuyen: [], dsSoXe: [] };
  if (lastRow <= 1) return ketQua;

  const data = sheet.getRange(2, 1, lastRow - 1, 17).getValues();
  const donViVanChuyenFilter = String(filters.donViVanChuyen || "").trim();
  const soXeFilter = String(filters.soXe || "").trim();
  const soTKHQFilter = String(filters.soTKHQ || "").trim();
  const tuNgay = filters.tuNgay ? new Date(filters.tuNgay + "T00:00:00+07:00") : null;
  const denNgay = filters.denNgay ? new Date(filters.denNgay + "T23:59:59+07:00") : null;

  const dsDVVCSet = new Set(); const dsXeSet = new Set();

  data.forEach(row => {
    const donViVanChuyen = String(row[7] || "").trim();
    const soXe = String(row[3] || "").trim();
    if (donViVanChuyen) dsDVVCSet.add(donViVanChuyen);
    if (soXe) dsXeSet.add(soXe);

    // "Ngày giờ cân 1" lưu dạng chuỗi text "dd/MM/yyyy HH:mm:ss" trong NL_PC_XH
    // (không phải kiểu Date thật) - tái sử dụng toDateObj() để phân tích đúng.
    const ngayCanObj = toDateObj(row[1]);
    if (!ngayCanObj || isNaN(ngayCanObj.getTime())) return;
    if (tuNgay && ngayCanObj < tuNgay) return;
    if (denNgay && ngayCanObj > denNgay) return;
    if (donViVanChuyenFilter && donViVanChuyen !== donViVanChuyenFilter) return;
    if (soXeFilter && soXe !== soXeFilter) return;
    const soTKHQ = String(row[14] || "").trim();
    if (soTKHQFilter && soTKHQ !== soTKHQFilter) return;

    ketQua.matched.push({
      soPhieu: row[0],
      ngayGioCan1: Utilities.formatDate(ngayCanObj, "GMT+7", "dd/MM/yyyy HH:mm:ss"),
      bienSo: soXe,
      canLan1: parseFloat(row[4]) || 0,
      canLan2: parseFloat(row[5]) || 0,
      klHang: parseFloat(row[6]) || 0,
      donViVanChuyen: donViVanChuyen,
      tenTaiXe: row[8] || "",
      khoiLuongTan: parseFloat(row[9]) || 0,
      soTKHQ: soTKHQ,
      khoXuat: row[15] || "",
      khoNhap: row[16] || "",
      ngayTS: ngayCanObj.getTime()
    });
  });

  ketQua.matched.sort((a, b) => b.ngayTS - a.ngayTS); // mới nhất lên trước
  ketQua.dsDonViVanChuyen = Array.from(dsDVVCSet).sort();
  ketQua.dsSoXe = Array.from(dsXeSet).sort();
  return ketQua;
}

// filters = {donViVanChuyen, soXe, soTKHQ, tuNgay, denNgay, trang} - trả về ĐÚNG
// 20 dòng/trang theo yêu cầu, kèm danh mục lọc và tổng số liệu TOÀN BỘ (không
// chỉ trang hiện tại).
function XH_getBaoCaoXuatQuaCan(filters) {
  try {
    filters = filters || {};
    const loc = XH_locBaoCaoXuatQuaCan_(filters);
    const KICH_THUOC_TRANG_XH = 20;
    const tongSoDong = loc.matched.length;
    const tongSoTrang = Math.max(1, Math.ceil(tongSoDong / KICH_THUOC_TRANG_XH));
    let trang = parseInt(filters.trang, 10); if (isNaN(trang) || trang < 1) trang = 1;
    if (trang > tongSoTrang) trang = tongSoTrang;
    const batDau = (trang - 1) * KICH_THUOC_TRANG_XH;
    const trangHienTai = loc.matched.slice(batDau, batDau + KICH_THUOC_TRANG_XH);

    let tongKLHang = 0, tongTan = 0;
    loc.matched.forEach(m => { tongKLHang += m.klHang; tongTan += m.khoiLuongTan; });

    return {
      status: "success", data: trangHienTai, trang: trang, tongSoTrang: tongSoTrang, tongSoDong: tongSoDong,
      summary: { soLuong: tongSoDong, tongKLHang: tongKLHang, tongTan: tongTan },
      dsDonViVanChuyen: loc.dsDonViVanChuyen, dsSoXe: loc.dsSoXe
    };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

function XH_exportBaoCaoXuatQuaCanExcel(filters) {
  try {
    const loc = XH_locBaoCaoXuatQuaCan_(filters || {});
    if (loc.matched.length === 0) return { status: "error", message: "Không có dữ liệu phù hợp bộ lọc để xuất." };
    const headers = ["Số Phiếu", "Ngày Giờ Cân 1", "Số Xe", "KL Hàng (Kg)", "Khối Lượng (Tấn)", "Đơn Vị Vận Chuyển", "Tên Tài Xế", "Số TKHQ", "Kho Xuất", "Kho Nhập"];
    const rows = loc.matched.map(r => [r.soPhieu, r.ngayGioCan1, r.bienSo, r.klHang, r.khoiLuongTan, r.donViVanChuyen, r.tenTaiXe, r.soTKHQ, r.khoXuat, r.khoNhap]);
    const tempSS = createTempSheetForExport_("BaoCao_XuatQuaCan_" + Utilities.formatDate(new Date(), "GMT+7", "ddMM_HHmm"), headers, rows, [4, 5]);
    logAudit_('EXPORT_EXCEL', 'OK', 'Xuất báo cáo xuất qua cân, ' + loc.matched.length + ' dòng.');
    return { status: "success", url: getExportUrl_(tempSS, "xlsx") };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

function XH_exportBaoCaoXuatQuaCanPDF(filters) {
  try {
    const loc = XH_locBaoCaoXuatQuaCan_(filters || {});
    if (loc.matched.length === 0) return { status: "error", message: "Không có dữ liệu phù hợp bộ lọc để xuất." };
    const headers = ["Số Phiếu", "Ngày Giờ Cân 1", "Số Xe", "KL Hàng(Kg)", "KL(Tấn)", "Đơn Vị Vận Chuyển", "Tên Tài Xế", "Số TKHQ", "Kho Xuất", "Kho Nhập"];
    const rows = loc.matched.map(r => [r.soPhieu, r.ngayGioCan1, r.bienSo, r.klHang, r.khoiLuongTan, r.donViVanChuyen, r.tenTaiXe, r.soTKHQ, r.khoXuat, r.khoNhap]);
    const tempSS = createTempSheetForExport_("BaoCao_XuatQuaCan_PDF_" + Utilities.formatDate(new Date(), "GMT+7", "ddMM_HHmm"), headers, rows, [4, 5]);
    logAudit_('EXPORT_PDF', 'OK', 'Xuất PDF báo cáo xuất qua cân, ' + loc.matched.length + ' dòng.');
    return { status: "success", url: getExportUrl_(tempSS, "pdf", false) };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

// ---------- Báo cáo xuất kết xuất Misa (từ NL_DH_XB - đơn hàng xuất bán) ----------
// Thành tiền tính theo KL_BDMT × Đơn giá (USD) - khớp đúng cách dữ liệu mẫu gốc
// NL_DH_XB đã thể hiện (đơn giá áp cho tấn khô, không phải tấn tươi).
function XH_locBaoCaoXuatMisa_(filters) {
  const sheet = XH_ss_().getSheetByName(XUATHANG_CONFIG.SHEET_DHXB);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
  const tuNgay = filters.tuNgay ? new Date(filters.tuNgay + "T00:00:00+07:00") : null;
  const denNgay = filters.denNgay ? new Date(filters.denNgay + "T23:59:59+07:00") : null;

  const result = [];
  data.forEach(row => {
    if (!row[4]) return; // bỏ dòng trống (chưa có khách hàng)
    const ngay = row[1];
    if (!(ngay instanceof Date) || isNaN(ngay.getTime())) return;
    if (tuNgay && ngay < tuNgay) return;
    if (denNgay && ngay > denNgay) return;

    const klMT = parseFloat(row[8]) || 0;
    const klBDMT = parseFloat(row[9]) || 0;
    const donGia = parseFloat(row[7]) || 0;
    const thanhTienUSD = klBDMT * donGia;

    result.push({
      ngayDonHang: Utilities.formatDate(ngay, "GMT+7", "dd/MM/yyyy"), ngayTS: ngay.getTime(),
      soTKHQ: row[2], tau: row[3], khachHang: row[4], tenHangHoa: row[6],
      donGiaUSD: donGia, klMT: klMT, klBDMT: klBDMT, thanhTienUSD: thanhTienUSD, khoXuat: row[15] || ""
    });
  });
  result.sort((a, b) => b.ngayTS - a.ngayTS);
  return result;
}

function XH_getBaoCaoXuatMisa(filters) {
  try {
    const list = XH_locBaoCaoXuatMisa_(filters || {});
    let tongKLMT = 0, tongKLBDMT = 0, tongThanhTienUSD = 0;
    list.forEach(r => { tongKLMT += r.klMT; tongKLBDMT += r.klBDMT; tongThanhTienUSD += r.thanhTienUSD; });
    return { status: "success", data: list, summary: { soLuong: list.length, tongKLMT: tongKLMT, tongKLBDMT: tongKLBDMT, tongThanhTienUSD: tongThanhTienUSD } };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

function XH_exportBaoCaoXuatMisaExcel(filters) {
  try {
    const list = XH_locBaoCaoXuatMisa_(filters || {});
    if (list.length === 0) return { status: "error", message: "Không có dữ liệu phù hợp bộ lọc để xuất." };
    const headers = ["Ngày Đơn Hàng", "Số TKHQ", "Tàu", "Khách Hàng", "Tên Hàng Hóa", "Đơn Giá (USD)", "KL_MT", "KL_BDMT", "Thành Tiền (USD)", "Kho Xuất"];
    const rows = list.map(r => [r.ngayDonHang, r.soTKHQ, r.tau, r.khachHang, r.tenHangHoa, r.donGiaUSD, r.klMT, r.klBDMT, r.thanhTienUSD, r.khoXuat]);
    const tempSS = createTempSheetForExport_("BaoCao_XuatMisa_" + Utilities.formatDate(new Date(), "GMT+7", "ddMM_HHmm"), headers, rows, [6, 7, 8, 9]);
    logAudit_('EXPORT_EXCEL', 'OK', 'Xuất báo cáo xuất kết xuất Misa, ' + list.length + ' dòng.');
    return { status: "success", url: getExportUrl_(tempSS, "xlsx") };
  } catch (e) { return { status: "error", message: e.toString() }; }
}
