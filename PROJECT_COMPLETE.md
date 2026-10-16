# 🗺️ VIETNAM MAP AUTO MARKER - HOÀN THÀNH

## ✅ ĐÃ TẠO THÀNH CÔNG

### 📦 File Excel đã sẵn sàng
- **BanDoUngVien.xlsm** - File Excel hoàn chỉnh trên Desktop
  - Sheet "DanhSach": Có sẵn 10 ứng viên mẫu
  - Sheet "BanDo": Đã chèn ảnh bản đồ Việt Nam

### 🖼️ Ảnh bản đồ
- **BanDoVN.png** - Đã trích xuất từ file Word (trong thư mục VietnamMapAutoMarker)

### 💻 Mã nguồn VBA
- **MapGenerator.bas** - Mã VBA hoàn chỉnh với 63 tỉnh thành

## 🚀 CÁCH SỬ DỤNG

### Phương án 1: Tự động (Khuyên dùng)
1. **Double-click file:** `AddVBA.vbs` (trong thư mục VietnamMapAutoMarker)
2. Nếu gặp lỗi trust, làm theo hướng dẫn trong popup
3. Xong! File Excel đã có VBA và nút bấm

### Phương án 2: Thủ công
1. Mở file **BanDoUngVien.xlsm**
2. Nhấn **Alt + F11** (VBA Editor)
3. Menu: **Insert → Module**
4. Mở file **MapGenerator.bas** và copy toàn bộ vào Module
5. Đóng VBA Editor (**Alt + Q**)
6. Vào sheet **BanDo**
7. Click ảnh bản đồ → Name Box (góc trên trái) → gõ: **BanDoVN** → Enter
8. Tab **Developer → Insert → Button**
9. Vẽ nút, chọn macro: **GenerateVietnamMap**
10. Đặt tên nút: **TẠO BẢN ĐỒ**

## 📝 CÁCH DÙNG FILE EXCEL

1. Mở file **BanDoUngVien.xlsm**
2. Vào sheet **DanhSach**
3. Nhập danh sách ứng viên:
   - Cột A: STT (01, 02, 03...)
   - Cột B: Họ tên (IN HOA, KHÔNG DẤU)
   - Cột C: Tỉnh/Thành (IN HOA, KHÔNG DẤU)
   - Cột D: Loại (省, 首都, 市)
4. Vào sheet **BanDo**
5. Bấm nút **TẠO BẢN ĐỒ**
6. Hoàn thành!

## 📚 CÁC FILE TRONG DỰ ÁN

```
Desktop/
├── BanDoUngVien.xlsm          ⭐ File Excel chính
├── HUONG_DAN_THEM_VBA.txt     📄 Hướng dẫn thêm VBA thủ công
│
└── VietnamMapAutoMarker/
    ├── MapGenerator.bas       💻 Mã VBA chính (700+ dòng)
    ├── BanDoVN.png           🖼️ Ảnh bản đồ đã trích xuất
    ├── generate_excel.py     🐍 Script Python tạo Excel
    ├── AddVBA.vbs           📜 Script tự động thêm VBA
    ├── README.md            📖 Tài liệu tiếng Anh
    ├── HUONG_DAN_SU_DUNG.md 📖 Tài liệu tiếng Việt đầy đủ
    └── QUICK_START.txt      ⚡ Hướng dẫn nhanh 5 phút
```

## 🎯 TÍNH NĂNG

✅ **63 tỉnh thành đầy đủ** - Tọa độ chính xác theo bản đồ hành chính VN
✅ **Smart Layout** - Tự động xếp Text Box tránh chồng đè
✅ **Màu sắc phân biệt** - Đỏ cho Hà Nội/TP.HCM, xanh lá cho tỉnh khác
✅ **Mũi tên đỏ** - Kết nối từ Text Box đến vị trí chính xác trên bản đồ
✅ **One-click** - Chỉ cần bấm nút là tự động vẽ toàn bộ
✅ **Auto cleanup** - Tự động xóa đối tượng cũ trước khi vẽ mới

## 🔧 KHẮC PHỤC LỖI

### Lỗi: "Trust access to VBA project"
**Giải pháp:**
1. File → Options → Trust Center
2. Trust Center Settings
3. Macro Settings → ✓ Trust access to the VBA project object model
4. OK → Khởi động lại Excel

### Lỗi: "Không tìm thấy tỉnh: XXX"
**Giải pháp:**
- Tên tỉnh phải IN HOA, KHÔNG DẤU
- Ví dụ đúng: **HA NOI**, **PHU THO**, **HO CHI MINH**
- Ví dụ sai: "Hà Nội", "ha noi", "Phú Thọ"

### Lỗi: "Không tìm thấy BanDoVN"
**Giải pháp:**
1. Vào sheet BanDo
2. Click chọn ảnh bản đồ
3. Name Box (góc trên trái) → gõ: **BanDoVN** → Enter

### Text Box bị chồng lên nhau
**Giải pháp:**
- Kéo ảnh bản đồ lớn hơn
- Hoặc sửa trong VBA: `MIN_VERTICAL_SPACING` từ 35 lên 40-45

## 💡 MẸO SỬ DỤNG

### Tạo nhiều bản đồ
- Copy sheet BanDo → đặt tên khác nhau cho từng đợt phỏng vấn
- Mỗi sheet có thể vẽ độc lập

### Export sang PDF
- File → Export → Create PDF/XPS
- Chọn sheet BanDo → Export

### In ấn
- Khuyên in trên giấy A3 hoặc A2 để rõ nét
- Page Layout → Scale to fit nếu cần

### Backup dữ liệu
- Thường xuyên Save file
- Tạo bản sao trước khi thử nghiệm

## 📊 DỮ LIỆU MẪU CÓ SẴN

File Excel đã có sẵn 10 ứng viên mẫu:
1. NGUYEN DIN TRUNG → HA NOI
2. PHAM CHUNG HIEU → HUNG YEN
3. LE PHUOC QUAN → BINH DINH
4. LE HOANG THACH → HA TINH
5. TRIEU ANH MINH → PHU THO
6. LOC HOANG NHAN → LAM DONG
7. TUONG XUAN LOC → HA NOI
8. DAM QUANG VINH → HA NOI
9. TRIEU KHANH HUNG → PHU THO
10. NGUYEN VAN AN → HO CHI MINH

## 🔄 TẠO LẠI FILE EXCEL

Nếu muốn tạo lại từ đầu:
```bash
python VietnamMapAutoMarker/generate_excel.py
```

## 📞 HỖ TRỢ

Nếu gặp vấn đề:
1. Xem file **HUONG_DAN_SU_DUNG.md** để có hướng dẫn chi tiết
2. Xem file **QUICK_START.txt** để có checklist từng bước
3. Kiểm tra phần "Khắc phục lỗi" ở trên

## ⚖️ BẢN QUYỀN

- Mã nguồn mở, sử dụng tự do
- Phù hợp cho: Phỏng vấn lao động đi Nhật, báo cáo địa lý, thống kê dân cư

---

**Phiên bản:** 1.0  
**Ngày tạo:** 2026-10-16  
**Tương thích:** Excel 2013, 2016, 2019, 2021, Office 365  
**Hệ điều hành:** Windows 10/11

---

## 🎉 HOÀN THÀNH!

File **BanDoUngVien.xlsm** đã sẵn sàng sử dụng trên Desktop!

**Bước tiếp theo:**
1. Double-click **AddVBA.vbs** để tự động thêm VBA
2. Hoặc làm theo hướng dẫn thủ công ở trên
3. Mở file Excel và bấm nút "TẠO BẢN ĐỒ"

**Chúc bạn sử dụng hiệu quả! 🇻🇳🗺️**
