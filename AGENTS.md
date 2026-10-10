# AGENTS.md - Quy Tắc Tự Động Thực Thi (Autonomous Execution Rules)

## 1. Chế Độ Tự Động Hóa Tuyệt Đối (Auto-Run Mode)
- **Tự động chạy tiếp (Always Proceed):** Tiếp tục thực hiện các tác vụ liên tiếp cho đến khi hoàn thành mục tiêu của người dùng mà không dừng lại hỏi các câu hỏi trung gian hay xin xác nhận không cần thiết.
- **Không yêu cầu review trung gian:** Không đặt `RequestFeedback: true` trong artifact metadata trừ khi người dùng chủ động yêu cầu xem trước kế hoạch.
- **Tối ưu lệnh Terminal:** Luôn chọn lệnh tương thích với sandbox tiêu chuẩn để IDE tự động duyệt thực thi ngầm mà không kích hoạt cửa sổ xác nhận của hệ điều hành.
- **Chủ động sửa lỗi:** Nếu gặp lỗi, tự động phân tích và sửa lỗi ngay lập tức, không chờ người dùng nhắc.
