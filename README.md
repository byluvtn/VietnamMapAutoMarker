# Vietnam Map Auto Marker - Excel VBA Project

Automatically generate Vietnam map with candidate residence markers for Japanese labor interview forms.

## 🎯 Purpose

This Excel Macro (VBA) project automatically draws a Vietnam map with text boxes and red arrows marking the residence locations of candidates going to Japan, following the labor interview template format.

## ✨ Features

- **63 Provinces Coverage**: Complete coordinate system for all 63 Vietnamese provinces
- **Smart Layout**: Automatic left/right placement to prevent text box overlaps
- **Auto Formatting**: 
  - Hanoi & Ho Chi Minh City: Red background (#DC143C)
  - Other provinces: Green background (#6AA84F)
  - White bold text, red arrows (#E60000)
- **One-Click Generation**: Just click the button to create the entire map
- **Auto Cleanup**: Removes old objects before drawing new ones

## 📋 Requirements

- Microsoft Excel 2013 or later (with VBA support)
- `.xlsm` file format (Macro-Enabled Workbook)

## 🚀 Quick Start

### 1. Setup Excel File

Create two sheets:

**Sheet "DanhSach" (Candidate List):**
| Column | Content | Example |
|--------|---------|---------|
| A | Number (2 digits) | 01, 02, 03 |
| B | Full Name (UPPERCASE) | NGUYEN DIN TRUNG |
| C | Province (no accents) | HA NOI, PHU THO, BINH DINH |
| D | Type (Japanese) | 省 (province), 首都 (Hanoi), 市 (HCMC) |

**Sheet "BanDo" (Map):**
- Insert Vietnam map image with numbered provinces (1-63)
- **Important**: Name the image shape as `BanDoVN`

### 2. Import VBA Code

1. Press `Alt + F11` to open VBA Editor
2. Insert → Module
3. Copy all content from [`MapGenerator.bas`](./MapGenerator.bas)
4. Paste into the module window

### 3. Create Button

1. Go to sheet "BanDo"
2. Developer tab → Insert → Button (Form Control)
3. Assign macro: `GenerateVietnamMap`
4. Label the button: "GENERATE MAP" or "TẠO BẢN ĐỒ"

### 4. Run

1. Enter candidate data in "DanhSach" sheet
2. Go to "BanDo" sheet
3. Click the button
4. Done! Map is automatically generated

## 📝 Sample Data

```
STT   Full Name              Province        Type
01    NGUYEN DIN TRUNG       HA NOI          首都
02    PHAM CHUNG HIEU        HUNG YEN        省
03    LE PHUOC QUAN          BINH DINH       省
04    LE HOANG THACH         HA TINH         省
05    TRIEU ANH MINH         PHU THO         省
10    HO CHI MINH            HO CHI MINH     市
```

## 🗺️ Supported Provinces

All 63 Vietnamese provinces are supported:
- **Northern Vietnam** (25): Hanoi, Vinh Phuc, Bac Ninh, Hai Duong, Hai Phong, etc.
- **North Central Coast** (6): Thanh Hoa, Nghe An, Ha Tinh, Quang Binh, etc.
- **South Central Coast** (8): Da Nang, Quang Nam, Binh Dinh, Khanh Hoa, etc.
- **Central Highlands** (5): Kon Tum, Gia Lai, Dak Lak, Lam Dong, etc.
- **Southeast** (6): Binh Duong, Dong Nai, Ho Chi Minh, etc.
- **Mekong Delta** (13): Can Tho, An Giang, Kien Giang, Ca Mau, etc.

See [`HUONG_DAN_SU_DUNG.md`](./HUONG_DAN_SU_DUNG.md) for complete list.

## 🔧 Customization

### Change Colors

Edit constants in VBA code (lines 13-16):
```vba
Private Const COLOR_GREEN As Long = 6989903   ' Green background
Private Const COLOR_RED As Long = 14423100    ' Red background
Private Const COLOR_ARROW As Long = 15073280  ' Arrow color
```

### Adjust Text Box Size

Edit constants (lines 17-18):
```vba
Private Const TEXTBOX_WIDTH As Single = 220   ' Width
Private Const TEXTBOX_HEIGHT As Single = 30   ' Height
```

### Modify Province Coordinates

Edit the `GetProvinceCoordinates()` function (lines 23-97):
```vba
Case 1: coords = Array(0.48, 0.15)  ' Hanoi
' rx = 0.48 (48% from left), ry = 0.15 (15% from top)
```

## 🐛 Troubleshooting

### "Sheet 'DanhSach' not found"
- Create a sheet named exactly "DanhSach" (case-sensitive)

### "Image 'BanDoVN' not found"
- Select the map image
- In Name Box (top-left), type: `BanDoVN` and press Enter

### "Province not found: XXX"
- Check spelling in column C
- Must be UPPERCASE, NO ACCENTS
- Example: "HA NOI" not "Hà Nội" or "ha noi"

### Text boxes overlap
- Enlarge the map image
- Increase `MIN_VERTICAL_SPACING` constant (line 19) from 35 to 40-45

## 📚 Documentation

- **Vietnamese Guide**: [`HUONG_DAN_SU_DUNG.md`](./HUONG_DAN_SU_DUNG.md) - Detailed Vietnamese documentation
- **VBA Source**: [`MapGenerator.bas`](./MapGenerator.bas) - Main VBA module

## 🎨 Output Format

Text boxes display:
```
[NUMBER]. [FULL NAME] -> [PROVINCE] [TYPE]

Examples:
05. TRIEU ANH MINH -> PHU THO 省
01. NGUYEN DIN TRUNG -> HA NOI 首都
10. HO CHI MINH -> HO CHI MINH 市
```

## 🔑 Key Functions

- `GenerateVietnamMap()` - Main function to generate the map
- `GetProvinceCoordinates()` - Returns coordinates for 63 provinces
- `GetProvinceID()` - Province name lookup (supports multiple variations)
- `DrawCandidatesOnMap()` - Draws text boxes and arrows
- `CreateSampleData()` - Helper function to create test data

## 📦 Project Structure

```
VietnamMapAutoMarker/
├── MapGenerator.bas           # Main VBA source code
├── HUONG_DAN_SU_DUNG.md      # Vietnamese documentation
└── README.md                  # This file (English)
```

## 💡 Use Cases

- Labor interview presentations for Japan recruitment
- Geographic distribution visualization
- Candidate origin tracking
- Print on A3/A2 paper for office display
- Export to PDF for email distribution

## ⚙️ Technical Details

- **Language**: VBA (Visual Basic for Applications)
- **Excel Version**: 2013+, Office 365
- **File Format**: `.xlsm` (Macro-Enabled)
- **Coordinate System**: Relative (0.0 to 1.0) for scalability
- **Layout Algorithm**: Smart positioning to avoid overlaps

## 📄 License

Open source, free to use for personal and commercial purposes.

## 🤝 Contributing

Feel free to improve the code:
- Add more province name variations
- Improve coordinate accuracy
- Enhance layout algorithm
- Add new features

## 📞 Support

For issues or questions:
1. Check the troubleshooting section
2. Review setup steps carefully
3. Verify sheet names and data format
4. See Vietnamese guide for detailed instructions

---

**Version:** 1.0  
**Last Updated:** 2026-10-16  
**Compatible:** Microsoft Excel 2013, 2016, 2019, 2021, Office 365

---

🇻🇳 **Made for Vietnamese labor recruitment to Japan** 🇯🇵
