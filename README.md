# 🤖 Gemini for Excel (AI Assistant & Formulas)

> Bộ Add-in Microsoft Excel tích hợp sức mạnh của **Google Gemini AI**, được thiết kế theo kiến trúc **Office Web Add-in (Shared Runtime)** với mô hình **BYOK (Bring Your Own Key)**.

---

## ✨ Tính Năng Nổi Bật

1. **Khung trợ lý AI Sidebar (Task Pane)**:
   - **Tự động nhận diện vùng chọn**: Nhận diện ngay bảng/ô bạn đang bôi đen (`Sheet1!A1:D20`).
   - **Phân tích số liệu thông minh**: Tìm kiếm xu hướng, điểm bất thường, tóm tắt bảng báo cáo tài chính/doanh thu.
   - **Tự động viết công thức Excel**: Chuyển mô tả tiếng Việt thành công thức `=SUMIFS(...)`, `=XLOOKUP(...)` chuẩn xác.
   - **Ghi kết quả trực tiếp**: Nút 1-click chèn kết quả phân tích hoặc công thức vào ngay ô tính đang mở.
2. **Bộ công thức AI trong ô tính (`=GEMINI.*`)**:
   - `=GEMINI.ASK("Yêu cầu", [Ô_tham_chiếu])`: Xử lý mọi yêu cầu văn bản hoặc phân tích ngữ cảnh.
   - `=GEMINI.TRANSLATE(A2, "tiếng Nhật")`: Dịch thuật tự động hàng loạt ô.
   - `=GEMINI.EXTRACT(A2, "số điện thoại")`: Trích xuất Email, SĐT, Ngày tháng, Mã đơn... từ văn bản thô.
   - `=GEMINI.FORMULA("Mô tả cách tính", [Vùng_dữ_liệu])`: Tự sinh công thức Excel tương thích.
3. **Bộ đệm thông minh (In-Memory Cache)**:
   - Ngăn Excel gọi lại API nhiều lần khi tự động tính toán lại, giúp tiết kiệm quota và tăng tốc độ bảng tính.
4. **Mô hình BYOK (Mỗi người tự dùng Key của mình)**:
   - Không cần máy chủ trung gian phức tạp hay tốn chi phí duy trì.
   - Mọi người trong nhóm/văn phòng tự lưu API Key vào trình duyệt/Excel cá nhân.

---

## 🚀 Hướng Dẫn Cài Đặt & Dùng Thử Ngay (Local)

### Bước 1: Khởi động máy chủ nội bộ
Chỉ cần nhấp đúp vào tệp:
```
start.bat
```
*(Hoặc chạy lệnh: `python server.py` trong Terminal)*

Máy chủ sẽ tự động kích hoạt HTTPS tại `https://localhost:3000/`.

> **Lưu ý chứng chỉ SSL lần đầu trên máy tính**:
> Do Excel yêu cầu kết nối an toàn HTTPS, nếu mở Task Pane bị trắng hoặc báo lỗi bảo mật, bạn chỉ cần mở trình duyệt Edge hoặc Chrome, truy cập vào `https://localhost:3000/src/taskpane/taskpane.html` một lần và bấm **"Nâng cao" -> "Tiếp tục truy cập (Không an toàn)"** để máy tính chấp nhận chứng chỉ nội bộ.

---

### Bước 2: Nạp Add-in vào Microsoft Excel (Sideload)

#### Cách làm trên Excel Desktop (Windows / Mac):
1. Khởi động phần mềm **Microsoft Excel** và mở một file bất kỳ.
2. Trên thanh menu, vào thẻ **Insert (Chèn)** $\rightarrow$ chọn **Get Add-ins (Tải phần bổ trợ)** (hoặc *My Add-ins*).
3. Trong cửa sổ hiện ra, chọn mục **Manage My Add-ins (Quản lý phần bổ trợ của tôi)** $\rightarrow$ bấm vào mũi tên nhỏ chọn **Upload My Add-in (Tải lên phần bổ trợ của tôi)**.
4. Bấm **Browse (Duyệt)** và chọn đến tệp:
   ```
   g:\Tool\AI\Office AI\gemini for excel\manifest.xml
   ```
5. Bấm **Upload (Tải lên)**.

Ngay lập tức, bạn sẽ thấy nhóm biểu tượng **Gemini AI** xuất hiện trên tab **Home (Trang chủ)** của Excel!

---

## 🔑 Hướng Dẫn Cấu Hình API Key Cá Nhân

1. Bấm vào nút **Gemini Assistant** trên thanh Ribbon để mở Sidebar bên phải.
2. Chuyển sang thẻ **⚙️ Cài đặt**.
3. Dán **Gemini API Key** của bạn vào ô (nếu chưa có, bấm vào đường dẫn [Google AI Studio](https://aistudio.google.com/app/apikey) để tạo khóa miễn phí trong 30 giây).
4. Bấm nút **🧪 Kiểm tra kết nối**:
   - Nếu hiện dòng chữ xanh *"Kết nối thành công!"*, bạn đã sẵn sàng sử dụng!
5. Bấm **💾 Lưu Cài Đặt**.

---

## 🌐 Cách Chia Sẻ Cho Đồng Nghiệp Dùng Chung (Zero Server)

Để cả văn phòng hoặc bạn bè cùng dùng mà bạn không cần phải bật máy tính chạy `start.bat`:

1. **Deploy web tĩnh lên GitHub Pages hoặc Vercel (Miễn phí 100%)**:
   - Tạo một repository trên GitHub và tải toàn bộ các tệp trong thư mục này lên (hoặc chỉ cần thư mục `src`, `assets`).
   - Bật tính năng **GitHub Pages** trong Settings của repository để nhận một đường dẫn HTTPS (Ví dụ: `https://ten-ban.github.io/gemini-excel/`).
2. **Cập nhật URL trong tệp `manifest.xml`**:
   - Dùng Text Editor mở `manifest.xml`, thay thế toàn bộ chuỗi `https://localhost:3000` bằng URL GitHub Pages của bạn (Ví dụ: `https://ten-ban.github.io/gemini-excel`).
3. **Gửi tệp `manifest.xml` cho đồng nghiệp**:
   - Bạn chỉ cần gửi duy nhất tệp `manifest.xml` cho mọi người trong nhóm.
   - Họ làm theo [Bước 2](#bước-2-nạp-add-in-vào-microsoft-excel-sideload) để tải file này vào Excel của họ.
   - Mỗi người tự mở thẻ Cài đặt và dán API Key cá nhân của họ vào là xong!

---

## 📊 Bảng Tra Cứu Công Thức

| Cú pháp hàm | Mục đích | Ví dụ thực tế |
| :--- | :--- | :--- |
| `=GEMINI.ASK(prompt, [context])` | Phân loại, tóm tắt, trả lời câu hỏi | `=GEMINI.ASK("Phân loại sentiment: Tích cực, Tiêu cực hay Trung tính", B2)` |
| `=GEMINI.TRANSLATE(text, lang)` | Dịch nội dung sang ngôn ngữ khác | `=GEMINI.TRANSLATE(A2, "tiếng Anh")` |
| `=GEMINI.EXTRACT(text, target)` | Bóc tách thông tin đặc thù | `=GEMINI.EXTRACT(A2, "email")` |
| `=GEMINI.FORMULA(request, [sample])` | Sinh công thức Excel tự động | `=GEMINI.FORMULA("Tính trung bình cột C nếu cột B > 100", B2:C20)` |

---

## 🛠️ Cấu Trúc Mã Nguồn

```
gemini for excel/
├── manifest.xml                 # Cấu hình Add-in cho Microsoft Office
├── server.py                    # Máy chủ HTTPS cục bộ (tự sinh SSL x509)
├── start.bat                    # Phím tắt khởi chạy 1-click trên Windows
├── package.json                 # Định nghĩa metadata & scripts
├── generate_assets.py           # Script sinh bộ icons kích thước 16, 32, 64, 80, 128px
├── assets/                      # Thư mục chứa biểu tượng Add-in
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html        # Giao diện Sidebar 3 tab (Chat, Hàm, Cài đặt)
│   │   ├── taskpane.css         # Bộ giao diện chuẩn Office Fluent UI
│   │   └── taskpane.js          # Controller điều khiển tương tác bảng tính và Gemini
│   └── functions/
│       ├── functions.json       # Metadata khai báo hàm cho Excel IntelliSense
│       └── functions.js         # Logic thực thi các hàm =GEMINI.* kèm Cache
└── README.md                    # Tài liệu hướng dẫn sử dụng
```
