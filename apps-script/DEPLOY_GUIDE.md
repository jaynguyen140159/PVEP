# Hướng Dẫn Cập Nhật Google Apps Script Backend (0đ - Không Cần VPS)

Sau khi mã nguồn `apps-script/Code.gs` đã được nâng cấp để hỗ trợ **đăng nhập bảo mật phía Google Sheet**, **đồng bộ đa thiết bị (`SAN_PHAM`, `TO_TRINH`, `NHAT_KY`)** và **chống ghi đè đồng thời (`LockService` + kiểm tra xung đột)**:

1. Mở Google Sheet kho PVEP của bạn → chọn menu **Tiện ích mở rộng (Extensions)** → **Apps Script**.
2. Dán toàn bộ nội dung file `apps-script/Code.gs` mới vào file `Code.gs` (hoặc `Mã.gs`) trên trình soạn thảo Apps Script và nhấn **Lưu (`Cmd + S`)**.
3. Chọn **Triển khai (Deploy)** → **Quản lý các bản triển khai (Manage deployments)** → bấm biểu tượng **Bút chì (Edit)** ở bản triển khai Web App hiện có:
   - Mục **Phiên bản (Version)**: chọn **Phiên bản mới (New version)**.
   - Mục **Thực thi dưới danh nghĩa (Execute as)**: **Tôi (Me)**.
   - Mục **Ai có quyền truy cập (Who has access)**: bắt buộc chọn **Bất kỳ ai (Anyone)** (để trình duyệt đọc được phản hồi JSON bảo mật từ Apps Script, không bị lỗi CORS).
   - Bấm **Triển khai (Deploy)** (giữ nguyên đường link `/exec` cũ để website nhận ngay mà không phải đổi URL).

> **Lưu ý về tab dữ liệu:**
> - Các tab `TO_TRINH` và `NHAT_KY` sẽ tự động được tạo kèm dòng tiêu đề ngay lần đầu sử dụng.
> - Đảm bảo tab `TAI_KHOAN` có sẵn các cột: `Username`, `Password`, `HoTen`, `ChucVu`, `TrangThai` (giá trị `HoatDong`).
