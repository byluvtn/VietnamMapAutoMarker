"""
Simple Auto-Generate Vietnam Map Excel File
============================================
"""
import os
import sys
from pathlib import Path

# Cài đặt thư viện cần thiết
print("Checking and installing required packages...")
os.system('pip install openpyxl python-docx Pillow -q')

import zipfile
import shutil
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font, PatternFill, Alignment

# Đường dẫn
DESKTOP = Path.home() / "Desktop"
WORD_FILE = DESKTOP / "男性_面接者所在地地図2601004.docx"
OUTPUT_DIR = DESKTOP / "VietnamMapAutoMarker"
MAP_IMAGE = OUTPUT_DIR / "BanDoVN.png"
OUTPUT_EXCEL = DESKTOP / "BanDoUngVien.xlsm"

print(f"Desktop: {DESKTOP}")
print(f"Word file: {WORD_FILE}")
print(f"Output: {OUTPUT_EXCEL}")

# Bước 1: Trích xuất ảnh
print("\n[1/3] Extracting map image from Word file...")
if WORD_FILE.exists():
    try:
        temp_dir = OUTPUT_DIR / "temp"
        temp_dir.mkdir(exist_ok=True, parents=True)
        
        with zipfile.ZipFile(WORD_FILE, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        media_dir = temp_dir / "word" / "media"
        if media_dir.exists():
            images = list(media_dir.glob("*.*"))
            if images:
                largest = max(images, key=lambda x: x.stat().st_size)
                shutil.copy(largest, MAP_IMAGE)
                print(f"   ✓ Extracted: {MAP_IMAGE.name}")
        
        shutil.rmtree(temp_dir)
    except Exception as e:
        print(f"   ✗ Error: {e}")
else:
    print(f"   ✗ Word file not found")

# Bước 2: Tạo Excel
print("\n[2/3] Creating Excel file...")
wb = Workbook()

# Sheet DanhSach
ws1 = wb.active
ws1.title = "DanhSach"
ws1.append(["STT", "Họ và tên", "Tỉnh/Thành", "Loại đơn vị"])

# Format header
for col in range(1, 5):
    cell = ws1.cell(1, col)
    cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    cell.font = Font(bold=True, color="FFFFFF")
    cell.alignment = Alignment(horizontal="center", vertical="center")

# Sample data
data = [
    ["01", "NGUYEN DIN TRUNG", "HA NOI", "首都"],
    ["02", "PHAM CHUNG HIEU", "HUNG YEN", "省"],
    ["03", "LE PHUOC QUAN", "BINH DINH", "省"],
    ["04", "LE HOANG THACH", "HA TINH", "省"],
    ["05", "TRIEU ANH MINH", "PHU THO", "省"],
    ["06", "LOC HOANG NHAN", "LAM DONG", "省"],
    ["07", "TUONG XUAN LOC", "HA NOI", "首都"],
    ["08", "DAM QUANG VINH", "HA NOI", "首都"],
    ["09", "TRIEU KHANH HUNG", "PHU THO", "省"],
    ["10", "NGUYEN VAN AN", "HO CHI MINH", "市"]
]

for row in data:
    ws1.append(row)

ws1.column_dimensions['A'].width = 8
ws1.column_dimensions['B'].width = 25
ws1.column_dimensions['C'].width = 18
ws1.column_dimensions['D'].width = 12

print(f"   ✓ Created sheet 'DanhSach' with {len(data)} candidates")

# Sheet BanDo
ws2 = wb.create_sheet("BanDo")
ws2['A1'] = "企業名：久野金属株式会社 様"
ws2['A1'].font = Font(size=12, bold=True)
ws2['A2'] = "ベトナム全国地図"
ws2['A2'].font = Font(size=14, bold=True)

if MAP_IMAGE.exists():
    try:
        img = ExcelImage(str(MAP_IMAGE))
        img.width = 600
        img.height = int(600 * (img.height / img.width))
        ws2.add_image(img, 'C3')
        print(f"   ✓ Added map image to sheet 'BanDo'")
    except Exception as e:
        print(f"   ✗ Could not add image: {e}")
else:
    print(f"   ✗ Map image not found")

ws2.column_dimensions['A'].width = 30

# Save
wb.save(OUTPUT_EXCEL)
print(f"   ✓ Saved Excel file: {OUTPUT_EXCEL.name}")

# Bước 3: Tạo hướng dẫn
print("\n[3/3] Creating instruction file...")
instruction = DESKTOP / "HUONG_DAN_THEM_VBA.txt"
with open(instruction, 'w', encoding='utf-8') as f:
    f.write("="*60 + "\n")
    f.write("   HƯỚNG DẪN THÊM VBA VÀO FILE EXCEL\n")
    f.write("="*60 + "\n\n")
    f.write(f"✓ File Excel đã được tạo: {OUTPUT_EXCEL.name}\n\n")
    f.write("CÁC BƯỚC TIẾP THEO:\n\n")
    f.write("1. Mở file BanDoUngVien.xlsm\n")
    f.write("2. Nhấn Alt + F11 (mở VBA Editor)\n")
    f.write("3. Menu: Insert → Module\n")
    f.write("4. Mở file MapGenerator.bas trong thư mục VietnamMapAutoMarker\n")
    f.write("5. Copy toàn bộ và paste vào Module\n")
    f.write("6. Đóng VBA Editor (Alt + Q)\n")
    f.write("7. Vào sheet 'BanDo'\n")
    f.write("8. Tab Developer → Insert → Button (Form Control)\n")
    f.write("9. Vẽ nút và chọn macro: GenerateVietnamMap\n")
    f.write("10. Đặt tên nút: TẠO BẢN ĐỒ\n\n")
    f.write("HOÀN TẤT!\n")

print(f"   ✓ Created instruction file: {instruction.name}")

print("\n" + "="*60)
print("   DONE!")
print("="*60)
print(f"\n✓ Excel file: {OUTPUT_EXCEL}")
print(f"✓ Map image: {MAP_IMAGE}")
print(f"✓ Instructions: {instruction}")
print("\nNext: Follow instructions in HUONG_DAN_THEM_VBA.txt")
