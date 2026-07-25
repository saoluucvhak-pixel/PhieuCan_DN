/*********************************************************
 * FILE CẤU HÌNH TRUNG TÂM - HỆ THỐNG QUẢN LÝ HAKGROUP
 *
 * Mục đích: gom toàn bộ ID Spreadsheet/Folder, tên Sheet, hằng số dùng chung
 * vào MỘT NƠI DUY NHẤT, để khi cần đổi (ví dụ đổi ID Google Sheet, đổi tên
 * Sheet, đổi thư mục Drive...) chỉ cần sửa ở đây, không phải tìm rải rác
 * khắp các file Code.gs / Index.html.
 *
 * LƯU Ý: Apps Script gộp TẤT CẢ các file .gs trong project vào chung 1 phạm
 * vi toàn cục (global scope), nên các hàm trong Code.gs vẫn gọi được
 * CONFIG.xxx / BAOGIA_CONFIG.xxx bình thường mà không cần "import" gì thêm.
 * Thứ tự các file trong project (Config.gs, Code.gs, Index.html) không ảnh
 * hưởng gì, vì mọi hàm chỉ thực sự chạy SAU KHI toàn bộ project đã được nạp.
 *********************************************************/

/* ---------- CẤU HÌNH HỆ THỐNG CÂN HÀNG & THANH TOÁN (PhieuCan_DN) ---------- */
const CONFIG = {
  // Thư mục Drive chứa file Excel phiếu cân mới tải lên chờ xử lý
  FOLDER_INPUT: "1sybSo9vSdSq_puu1LQR2m1zT59ZopgrL",
  // Thư mục Drive lưu file đã xử lý xong (file gốc đã import + các file export)
  FOLDER_DONE: "1bAp97Lwrpq6N8z4-2oXSaieszSL2roca",

  // Google Sheet chính chứa dữ liệu phiếu cân
  SPREADSHEET_ID: "1vqMVxccBA7zlAMHrGsVBydGFwZJ6QuDZW10zJ74V29g",
  DATA_SHEET: "PhieuCan_DN",

  // Đường dẫn/spreadsheet Báo giá dùng để tra giá khi tính tiền phiếu cân
  // (chính là spreadsheet của BAOGIA_CONFIG bên dưới - xem thêm ghi chú ở đó)
  URL_BAO_GIA: "https://docs.google.com/spreadsheets/d/1SIhfjP5-6ouRPDj265lAMmI5yWs1XcnedjqpzDwaIC0/edit",
  SHEET_BAO_GIA: "Baogia_DN_SAVE",

  // Sheet tham chiếu dùng khi trích xuất dữ liệu hạch toán MISA (copyDataWithFinalLookup)
  SRC_FILE_ID: "1cv11ORWuAF3Sit4f-kA0xrP6-ab4SF-7LEdkCvGi_gI",
  // Spreadsheet đích chứa dữ liệu đã map sẵn theo đúng cấu trúc import MISA
  MISA_DST_ID: "1vkeu2YxME6fsp9ed8DokdtV1jxla5pA-H7heHBt-BRs",
  MISA_DST_SHEET: "Update_MiSa_PC",
  // Sheet Đề Nghị Thanh Toán tạm dùng khi map số hợp đồng cho MISA
  DNTT_FILE_ID: "1oUm87_gbDbnuPc_We0dyZ_e4kHXBHXs95AQAxp5okYo",
  DNTT_SHEET: "DNTT_GK_DN_CT",

  // Timeout chờ khóa LockService dùng chung (ms) - chống ghi đè dữ liệu khi nhiều người dùng cùng lúc
  LOCK_TIMEOUT_MS: 30000,

  // Tên sheet log audit (Timestamp, Action, Status, Message) - đã có sẵn trong hệ thống
  AUDIT_SHEET: "Audit"
};

/* ---------- CẤU HÌNH HỆ THỐNG QUẢN LÝ BÁO GIÁ (Baogia_DN...) ---------- */
// Đây là 1 spreadsheet RIÊNG (khác với CONFIG.SPREADSHEET_ID ở trên), nhưng
// CHÍNH LÀ spreadsheet mà CONFIG.URL_BAO_GIA / CONFIG.SHEET_BAO_GIA trỏ tới -
// vì vậy SAVE_SHEET bên dưới lấy trực tiếp từ CONFIG để tránh khai báo lệch
// tên sheet ở 2 nơi khác nhau gây tra sai dữ liệu giá.
const BAOGIA_CONFIG = {
  SPREADSHEET_ID: "1SIhfjP5-6ouRPDj265lAMmI5yWs1XcnedjqpzDwaIC0",
  // Thư mục Drive lưu các file báo giá export ra (Excel/PDF)
  BACKUP_FOLDER_ID: "1N9qkwh0qBcdYaU4vJI307Eh5pSnL7zRA",

  SRC_SHEET: "Baogia_DN",         // Log các nhóm giá đã nhập (mỗi dòng = 1 nhóm mã dùng chung 1 giá)
  QL_SHEET: "QL_BaoGia",          // Đầu phiếu báo giá (Số báo giá, ngày, hiệu lực, sao chép từ đâu)
  DST_SHEET: "Baogia_DN_FINAL",   // Chỉ chứa các mã đang "Còn hiệu lực" (dùng để tính giá)
  SAVE_SHEET: CONFIG.SHEET_BAO_GIA, // "Baogia_DN_SAVE" - toàn bộ lịch sử (còn/chưa/hết hiệu lực)
  MA_SHEET: "Ma_BaoGia",          // Danh mục mã báo giá (Đại lý_Nguồn gốc_Hình ảnh)
  MAKL_SHEET: "Ma_KL"             // Danh mục mã khối lượng (dải Tấn áp dụng giá)
};

/* ---------- HẰNG SỐ DÙNG CHUNG KHÁC ---------- */
// Tên công ty hiển thị trên các phiếu in (Phiếu nhập kho, bảng báo giá xuất Excel...)
const COMPANY_NAME = "CÔNG TY TNHH HOÀNG ANH KHÔI";
