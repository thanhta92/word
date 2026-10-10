# CloudDocs - Trình Quản Lý & Soạn Thảo Word Online (PWA)

Ứng dụng soạn thảo văn bản Word (.docx) chuyên nghiệp kết hợp quản lý thư mục và đồng bộ đám mây thời gian thực (Realtime Cloud Sync) đa thiết bị giữa Điện thoại, Máy tính bảng và Laptop.

---

## 🚀 Hướng Dẫn Đưa Lên GitHub & Chạy Online Miễn Phí (GitHub Pages)

Để bạn và bất kỳ ai có thể truy cập bằng link online trên điện thoại và cài đặt về màn hình chính, làm theo 3 bước đơn giản sau:

### Bước 1: Tạo Repository Trên GitHub
1. Đăng nhập vào tài khoản [GitHub](https://github.com/).
2. Nhấn nút **New** (Tạo kho lưu trữ mới).
3. Đặt tên Repository (ví dụ: `word` hoặc `clouddocs`).
4. Chọn chế độ **Public** (Công khai).
5. Nhấn **Create repository**.

### Bước 2: Tải Code Lên GitHub
Bạn có thể dùng Git dòng lệnh hoặc tải trực tiếp trên trình duyệt:

#### Cách A: Dùng Git (Khuyên dùng)
Mở Terminal / PowerShell tại thư mục này và chạy các lệnh:
```bash
git init
git add .
git commit -m "Initial commit CloudDocs PWA"
git branch -M main
git remote add origin https://github.com/<tên_tài_khoản_github>/<tên_repo>.git
git push -u origin main --force
```

#### Cách B: Tải trực tiếp bằng trình duyệt (Không cần cài Git)
1. Trong trang repository vừa tạo trên GitHub, chọn **"uploading an existing file"**.
2. Kéo thả toàn bộ các file và thư mục (`index.html`, `manifest.json`, `sw.js`, `.nojekyll`, thư mục `css/`, `js/`, `icons/`) vào trang GitHub.
3. Nhấn **Commit changes**.

---

### Bước 3: Kích Hoạt GitHub Pages (Chạy Online)
1. Trong repository trên GitHub, vào mục **Settings** (ở thanh menu trên cùng).
2. Ở cột bên trái, chọn **Pages**.
3. Tại mục **Build and deployment**:
   - **Source:** Chọn `Deploy from a branch`.
   - **Branch:** Chọn `main`, thư mục `/ (root)`.
   - Nhấn **Save**.
4. Chờ khoảng 1 - 2 phút, GitHub sẽ hiển thị đường link trang web của bạn:
   👉 **`https://<tên_tài_khoản>.github.io/<tên_repo>/`**

---

## 📱 Cách Cài Đặt Về Màn Hình Chính Điện Thoại Như Một App Riêng

Sau khi mở đường link GitHub Pages trên điện thoại:

### 🍏 Trên iPhone / iPad (iOS Safari)
1. Mở link bằng trình duyệt **Safari**.
2. Chạm vào nút **Chia sẻ** (biểu tượng hình vuông có mũi tên hướng lên `⬆` ở thanh dưới màn hình).
3. Cuộn xuống và chọn **"Thêm vào Màn hình chính" (Add to Home Screen)**.
4. Nhấn **"Thêm" (Add)** ở góc trên bên phải.
5. 👉 Biểu tượng CloudDocs màu xanh Word sẽ xuất hiện ngay trên màn hình chính của bạn. Khi mở ra, app chạy toàn màn hình độc lập, không có thanh địa chỉ web, mượt mà như app tải từ App Store!

### 🤖 Trên Điện Thoại Android (Samsung, Xiaomi, Oppo, Pixel...)
1. Mở link bằng trình duyệt **Chrome**.
2. Bấm vào nút **"Cài App ĐT"** trên thanh tiêu đề ứng dụng, hoặc nhấn vào **menu 3 chấm (⋮)** ở góc trên bên phải Chrome.
3. Chọn **"Cài đặt ứng dụng"** hoặc **"Thêm vào Màn hình chính"**.
4. Chọn **Cài đặt**. Ứng dụng sẽ được cài đặt trực tiếp vào màn hình điện thoại.

---

## ⚡ Các Tính Năng Nổi Bật

- **Quản lý thư mục & tệp:** Tạo thư mục con, đổi tên, đổi màu icon, tạo tệp Word `.docx` mới, tải lên/tải xuống file docx thực tế.
- **Soạn thảo Word phong phú:** Thanh công cụ Ribbon đầy đủ (Home, Insert, Draw, Layout, v.v.), font chữ, in đậm, căn lề, chèn bảng, chèn ảnh, thước đo A4.
- **Đồng bộ thời gian thực (Realtime Cloud Sync):** Nhập văn bản từ điện thoại, laptop nhận ngay lập tức không cần F5 (sử dụng Firebase Realtime Database).
- **Mã QR kết nối nhanh:** Quét mã QR bằng camera điện thoại để kết nối cùng tài khoản trong 1 giây.
- **Offline-First:** Hoạt động ngay cả khi không có mạng Internet nhờ Service Worker cache và LocalStorage.
