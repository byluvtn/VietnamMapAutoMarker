Attribute VB_Name = "MapGenerator"
Option Explicit

'==============================================
' MODULE: MapGenerator
' MỤC ĐÍCH: Tự động vẽ bản đồ Việt Nam đánh dấu nơi cư trú ứng viên
' TÁC GIẢ: Auto-generated VBA Module
' NGÀY: 2026-10-16
'==============================================

' Hằng số định dạng
Private Const COLOR_GREEN As Long = 6989903      ' #6AA84F - Màu xanh lá cho tỉnh thường
Private Const COLOR_RED As Long = 14423100       ' #DC143C - Màu đỏ cho Hà Nội, TP.HCM
Private Const COLOR_WHITE As Long = 16777215     ' #FFFFFF - Màu chữ trắng
Private Const COLOR_ARROW As Long = 15073280     ' #E60000 - Màu mũi tên đỏ
Private Const ARROW_WIDTH As Single = 1.5        ' Độ dày mũi tên
Private Const TEXTBOX_WIDTH As Single = 220      ' Chiều rộng Text Box
Private Const TEXTBOX_HEIGHT As Single = 30      ' Chiều cao Text Box
Private Const MIN_VERTICAL_SPACING As Single = 35 ' Khoảng cách tối thiểu giữa các Text Box
Private Const MAP_ELEM_PREFIX As String = "MAP_ELEM_" ' Tiền tố nhận diện đối tượng do macro tạo ra

'==============================================
' TYPE: CandidateInfo
' MỤC ĐÍCH: Cấu trúc dữ liệu lưu thông tin ứng viên
'==============================================
Private Type CandidateInfo
    STT As String
    HoTen As String
    TinhThanh As String
    LoaiDonVi As String
    ProvinceID As Integer
End Type

'==============================================
' FUNCTION: GetProvinceCoordinates
' MỤC ĐÍCH: Trả về tọa độ tương đối (rx, ry) của 63 tỉnh thành VN
'           Tọa độ được đo trực tiếp theo vị trí tâm số thứ tự in trên
'           ảnh BanDoVN.png (rx, ry theo tỷ lệ 0.0 - 1.0 của chiều
'           rộng/chiều cao ảnh).
' THAM SỐ: provinceID - Số thứ tự tỉnh (1-63)
' TRẢ VỀ: Array(rx, ry) - Tọa độ tương đối từ 0.0 đến 1.0
'==============================================
Private Function GetProvinceCoordinates(ByVal provinceID As Integer) As Variant
    Dim coords As Variant
    
    Select Case provinceID
        ' MIỀN BẮC - Vùng Đồng bằng Sông Hồng
        Case 1: coords = Array(0.4, 0.23)    ' Hà Nội (首都)
        Case 2: coords = Array(0.38, 0.21)   ' Vĩnh Phúc
        Case 3: coords = Array(0.42, 0.23)   ' Bắc Ninh
        Case 4: coords = Array(0.44, 0.23)   ' Hải Dương
        Case 5: coords = Array(0.46, 0.25)   ' Hải Phòng
        Case 6: coords = Array(0.42, 0.25)   ' Hưng Yên
        Case 7: coords = Array(0.44, 0.26)   ' Hà Nam
        Case 8: coords = Array(0.41, 0.26)   ' Nam Định
        Case 9: coords = Array(0.45, 0.27)   ' Thái Bình
        Case 10: coords = Array(0.42, 0.28)  ' Ninh Bình
        
        ' MIỀN BẮC - Vùng Trung du Miền núi phía Bắc
        Case 11: coords = Array(0.49, 0.21)  ' Lạng Sơn (nhãn "11" trên ảnh)
        Case 12: coords = Array(0.33, 0.14)  ' Cao Bằng
        Case 13: coords = Array(0.43, 0.14)  ' Bắc Kạn
        Case 14: coords = Array(0.40, 0.17)  ' (nhãn "14" trên ảnh)
        Case 15: coords = Array(0.37, 0.17)  ' Tuyên Quang
        Case 16: coords = Array(0.28, 0.16)  ' Thái Nguyên
        Case 17: coords = Array(0.31, 0.19)  ' Lào Cai
        Case 18: coords = Array(0.40, 0.20)  ' Yên Bái
        Case 19: coords = Array(0.45, 0.18)  ' Phú Thọ
        Case 20: coords = Array(0.47, 0.21)  ' Bắc Giang
        Case 21: coords = Array(0.35, 0.21)  ' Quảng Ninh
        
        ' MIỀN BẮC - Vùng Tây Bắc
        Case 22: coords = Array(0.21, 0.19)  ' Điện Biên
        Case 23: coords = Array(0.21, 0.16)  ' Lai Châu
        Case 24: coords = Array(0.22, 0.21)  ' Sơn La
        Case 25: coords = Array(0.38, 0.25)  ' Hòa Bình
        
        ' BẮC TRUNG BỘ
        Case 26: coords = Array(0.37, 0.28)  ' Thanh Hóa
        Case 27: coords = Array(0.32, 0.33)  ' Nghệ An
        Case 28: coords = Array(0.40, 0.39)  ' Hà Tĩnh
        Case 29: coords = Array(0.44, 0.43)  ' Quảng Bình
        Case 30: coords = Array(0.48, 0.48)  ' Quảng Trị
        Case 31: coords = Array(0.53, 0.50)  ' Thừa Thiên Huế
        
        ' NAM TRUNG BỘ
        Case 32: coords = Array(0.58, 0.52)  ' Đà Nẵng
        Case 33: coords = Array(0.56, 0.54)  ' Quảng Nam
        Case 34: coords = Array(0.62, 0.58)  ' Quảng Ngãi
        Case 35: coords = Array(0.64, 0.61)  ' Bình Định
        Case 36: coords = Array(0.65, 0.64)  ' Phú Yên
        Case 37: coords = Array(0.65, 0.67)  ' Khánh Hòa
        Case 38: coords = Array(0.64, 0.71)  ' Ninh Thuận
        Case 39: coords = Array(0.63, 0.73)  ' Bình Thuận
        
        ' TÂY NGUYÊN
        Case 40: coords = Array(0.56, 0.59)  ' Kon Tum
        Case 41: coords = Array(0.59, 0.63)  ' Gia Lai
        Case 42: coords = Array(0.58, 0.67)  ' Đắk Lắk
        Case 43: coords = Array(0.55, 0.69)  ' Đắk Nông
        Case 44: coords = Array(0.58, 0.72)  ' Lâm Đồng
        
        ' ĐÔNG NAM BỘ
        Case 45: coords = Array(0.47, 0.71)  ' Bình Phước
        Case 46: coords = Array(0.42, 0.74)  ' Tây Ninh
        Case 47: coords = Array(0.46, 0.74)  ' Bình Dương
        Case 48: coords = Array(0.50, 0.75)  ' Đồng Nai
        Case 49: coords = Array(0.63, 0.73)  ' Bà Rịa-Vũng Tàu (gần 39 trên bản đồ)
        Case 50: coords = Array(0.47, 0.76)  ' TP. Hồ Chí Minh (市)
        
        ' ĐỒNG BẰNG SÔNG CỬU LONG
        Case 51: coords = Array(0.42, 0.75)  ' Long An
        Case 52: coords = Array(0.42, 0.77)  ' Tiền Giang
        Case 53: coords = Array(0.46, 0.78)  ' Bến Tre
        Case 54: coords = Array(0.42, 0.81)  ' Trà Vinh
        Case 55: coords = Array(0.41, 0.80)  ' Vĩnh Long
        Case 56: coords = Array(0.38, 0.77)  ' Đồng Tháp
        Case 57: coords = Array(0.34, 0.79)  ' An Giang
        Case 58: coords = Array(0.35, 0.82)  ' Kiên Giang
        Case 59: coords = Array(0.37, 0.80)  ' Cần Thơ
        Case 60: coords = Array(0.39, 0.82)  ' Hậu Giang
        Case 61: coords = Array(0.42, 0.83)  ' Sóc Trăng
        Case 62: coords = Array(0.38, 0.85)  ' Bạc Liêu
        Case 63: coords = Array(0.34, 0.87)  ' Cà Mau
        
        Case Else: coords = Array(0.5, 0.5)  ' Mặc định giữa bản đồ
    End Select
    
    GetProvinceCoordinates = coords
End Function

'==============================================
' FUNCTION: GetProvinceID
' MỤC ĐÍCH: Tra cứu mã số tỉnh từ tên tỉnh
' THAM SỐ: provinceName - Tên tỉnh (không dấu, in hoa)
' TRẢ VỀ: Integer - Mã số tỉnh (1-63), 0 nếu không tìm thấy
'==============================================
Private Function GetProvinceID(ByVal provinceName As String) As Integer
    provinceName = UCase(Trim(provinceName))
    
    ' Chuẩn hóa tên tỉnh
    provinceName = Replace(provinceName, "TP.", "")
    provinceName = Replace(provinceName, "THANH PHO", "")
    provinceName = Trim(provinceName)
    
    Select Case provinceName
        ' MIỀN BẮC
        Case "HA NOI", "HANOI", "HÀ NỘI": GetProvinceID = 1
        Case "VINH PHUC", "VĨNH PHÚC": GetProvinceID = 2
        Case "BAC NINH", "BẮC NINH": GetProvinceID = 3
        Case "HAI DUONG", "HẢI DƯƠNG": GetProvinceID = 4
        Case "HAI PHONG", "HẢI PHÒNG": GetProvinceID = 5
        Case "HUNG YEN", "HƯNG YÊN": GetProvinceID = 6
        Case "HA NAM", "HÀ NAM": GetProvinceID = 7
        Case "NAM DINH", "NAM ĐỊNH": GetProvinceID = 8
        Case "THAI BINH", "THÁI BÌNH": GetProvinceID = 9
        Case "NINH BINH", "NINH BÌNH": GetProvinceID = 10
        
        Case "HA GIANG", "HÀ GIANG": GetProvinceID = 11
        Case "CAO BANG", "CAO BẰNG": GetProvinceID = 12
        Case "BAC KAN", "BẮC KẠN": GetProvinceID = 13
        Case "LANG SON", "LẠNG SƠN": GetProvinceID = 14
        Case "TUYEN QUANG", "TUYÊN QUANG": GetProvinceID = 15
        Case "THAI NGUYEN", "THÁI NGUYÊN": GetProvinceID = 16
        Case "LAO CAI", "LÀO CAI": GetProvinceID = 17
        Case "YEN BAI", "YÊN BÁI": GetProvinceID = 18
        Case "PHU THO", "PHÚ THỌ": GetProvinceID = 19
        Case "BAC GIANG", "BẮC GIANG": GetProvinceID = 20
        Case "QUANG NINH", "QUẢNG NINH": GetProvinceID = 21
        
        Case "DIEN BIEN", "ĐIỆN BIÊN": GetProvinceID = 22
        Case "LAI CHAU", "LAI CHÂU": GetProvinceID = 23
        Case "SON LA", "SƠN LA": GetProvinceID = 24
        Case "HOA BINH", "HÒA BÌNH": GetProvinceID = 25
        
        ' BẮC TRUNG BỘ
        Case "THANH HOA", "THANH HÓA": GetProvinceID = 26
        Case "NGHE AN", "NGHỆ AN": GetProvinceID = 27
        Case "HA TINH", "HÀ TĨNH": GetProvinceID = 28
        Case "QUANG BINH", "QUẢNG BÌNH": GetProvinceID = 29
        Case "QUANG TRI", "QUẢNG TRỊ": GetProvinceID = 30
        Case "THUA THIEN HUE", "THỪA THIÊN HUẾ", "HUE", "HUẾ": GetProvinceID = 31
        
        ' NAM TRUNG BỘ
        Case "DA NANG", "ĐÀ NẴNG": GetProvinceID = 32
        Case "QUANG NAM", "QUẢNG NAM": GetProvinceID = 33
        Case "QUANG NGAI", "QUẢNG NGÃI": GetProvinceID = 34
        Case "BINH DINH", "BÌNH ĐỊNH": GetProvinceID = 35
        Case "PHU YEN", "PHÚ YÊN": GetProvinceID = 36
        Case "KHANH HOA", "KHÁNH HÒA": GetProvinceID = 37
        Case "NINH THUAN", "NINH THUẬN": GetProvinceID = 38
        Case "BINH THUAN", "BÌNH THUẬN": GetProvinceID = 39
        
        ' TÂY NGUYÊN
        Case "KON TUM", "KONTUM": GetProvinceID = 40
        Case "GIA LAI": GetProvinceID = 41
        Case "DAK LAK", "ĐẮK LẮK", "DAKLAK": GetProvinceID = 42
        Case "DAK NONG", "ĐẮK NÔNG", "DAKNONG": GetProvinceID = 43
        Case "LAM DONG", "LÂM ĐỒNG": GetProvinceID = 44
        
        ' ĐÔNG NAM BỘ
        Case "BINH PHUOC", "BÌNH PHƯỚC": GetProvinceID = 45
        Case "TAY NINH", "TÂY NINH": GetProvinceID = 46
        Case "BINH DUONG", "BÌNH DƯƠNG": GetProvinceID = 47
        Case "DONG NAI", "ĐỒNG NAI": GetProvinceID = 48
        Case "BA RIA VUNG TAU", "BÀ RỊA VŨNG TÀU", "VUNG TAU": GetProvinceID = 49
        Case "HO CHI MINH", "HỒ CHÍ MINH", "HOCHIMINH", "SAI GON", "SAIGON": GetProvinceID = 50
        
        ' ĐỒNG BẰNG SÔNG CỬU LONG
        Case "LONG AN": GetProvinceID = 51
        Case "TIEN GIANG", "TIỀN GIANG": GetProvinceID = 52
        Case "BEN TRE", "BẾN TRE": GetProvinceID = 53
        Case "TRA VINH", "TRÀ VINH": GetProvinceID = 54
        Case "VINH LONG", "VĨNH LONG": GetProvinceID = 55
        Case "DONG THAP", "ĐỒNG THÁP": GetProvinceID = 56
        Case "AN GIANG": GetProvinceID = 57
        Case "KIEN GIANG", "KIÊN GIANG": GetProvinceID = 58
        Case "CAN THO", "CẦN THƠ": GetProvinceID = 59
        Case "HAU GIANG", "HẬU GIANG": GetProvinceID = 60
        Case "SOC TRANG", "SÓC TRĂNG": GetProvinceID = 61
        Case "BAC LIEU", "BẠC LIÊU": GetProvinceID = 62
        Case "CA MAU", "CÀ MAU": GetProvinceID = 63
        
        Case Else: GetProvinceID = 0
    End Select
End Function

'==============================================
' SUB: GenerateVietnamMap (MAIN PROCEDURE)
' MỤC ĐÍCH: Hàm chính để tạo bản đồ Việt Nam với đánh dấu ứng viên
' CÁCH DÙNG: Gán cho nút Button trên sheet "BanDo"
'==============================================
Public Sub GenerateVietnamMap()
    On Error GoTo ErrorHandler
    
    Dim wsDanhSach As Worksheet
    Dim wsBanDo As Worksheet
    Dim mapImage As Shape
    Dim candidates() As CandidateInfo
    Dim candidateCount As Integer
    Dim i As Long
    Dim lastRow As Long
    
    ' Tắt cập nhật màn hình để tăng tốc độ
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Lấy các sheet cần thiết
    On Error Resume Next
    Set wsDanhSach = ThisWorkbook.Worksheets("DanhSach")
    Set wsBanDo = ThisWorkbook.Worksheets("BanDo")
    On Error GoTo ErrorHandler
    
    If wsDanhSach Is Nothing Then
        MsgBox "Không tìm thấy sheet 'DanhSach'!", vbCritical
        GoTo CleanUp
    End If
    
    If wsBanDo Is Nothing Then
        MsgBox "Không tìm thấy sheet 'BanDo'!", vbCritical
        GoTo CleanUp
    End If
    
    ' Tìm ảnh bản đồ
    On Error Resume Next
    Set mapImage = wsBanDo.Shapes("BanDoVN")
    On Error GoTo ErrorHandler
    
    If mapImage Is Nothing Then
        MsgBox "Không tìm thấy ảnh bản đồ 'BanDoVN' trên sheet 'BanDo'!" & vbCrLf & _
               "Vui lòng đặt tên ảnh bản đồ là 'BanDoVN'.", vbCritical
        GoTo CleanUp
    End If
    
    ' Xóa tất cả Text Box và mũi tên cũ do macro tạo ra (không đụng đến
    ' ảnh bản đồ gốc BanDoVN và nút bấm TAO BAN DO)
    Call ClearOldObjects(wsBanDo)
    
    ' Đọc dữ liệu từ sheet DanhSach
    lastRow = wsDanhSach.Cells(wsDanhSach.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "Không có dữ liệu trong sheet 'DanhSach'!", vbInformation
        GoTo CleanUp
    End If
    
    ' Thu thập thông tin ứng viên
    candidateCount = 0
    ReDim candidates(1 To lastRow - 1)
    
    For i = 2 To lastRow
        If Trim(wsDanhSach.Cells(i, 2).Value) <> "" Then
            candidateCount = candidateCount + 1
            
            With candidates(candidateCount)
                .STT = Format(wsDanhSach.Cells(i, 1).Value, "00")
                .HoTen = UCase(Trim(wsDanhSach.Cells(i, 2).Value))
                .TinhThanh = UCase(Trim(wsDanhSach.Cells(i, 3).Value))
                .LoaiDonVi = Trim(wsDanhSach.Cells(i, 4).Value)
                .ProvinceID = GetProvinceID(.TinhThanh)
                
                If .ProvinceID = 0 Then
                    MsgBox "Không tìm thấy tỉnh: " & .TinhThanh & " (Dòng " & i & ")", vbExclamation
                End If
            End With
        End If
    Next i
    
    If candidateCount = 0 Then
        MsgBox "Không có dữ liệu hợp lệ!", vbInformation
        GoTo CleanUp
    End If
    
    ' Vẽ Text Box và mũi tên
    ReDim Preserve candidates(1 To candidateCount)
    Call DrawCandidatesOnMap(wsBanDo, mapImage, candidates, candidateCount)
    
    MsgBox "Đã tạo bản đồ thành công với " & candidateCount & " ứng viên!", vbInformation
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Lỗi: " & Err.Description, vbCritical
End Sub

'==============================================
' SUB: ClearOldObjects
' MỤC ĐÍCH: Xóa tất cả Text Box và mũi tên do macro tạo ra ở lần chạy
'           trước, dựa theo tiền tố tên "MAP_ELEM_" đã gán khi tạo.
'           Nhờ vậy, ảnh bản đồ gốc "BanDoVN" và nút bấm "TAO BAN DO"
'           (không có tiền tố này) sẽ không bao giờ bị xóa nhầm, và
'           các Text Box/mũi tên cũ sẽ không bị vẽ chồng lên lần vẽ mới.
'==============================================
Private Sub ClearOldObjects(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim i As Long
    
    ' Xóa ngược từ cuối để tránh lỗi index khi Delete làm thay đổi Count
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        
        ' Chỉ xóa các đối tượng do macro tạo ra (tên bắt đầu bằng MAP_ELEM_)
        If Left(shp.Name, Len(MAP_ELEM_PREFIX)) = MAP_ELEM_PREFIX Then
            shp.Delete
        End If
    Next i
End Sub

'==============================================
' SUB: DrawCandidatesOnMap
' MỤC ĐÍCH: Vẽ Text Box và mũi tên cho tất cả ứng viên
'==============================================
Private Sub DrawCandidatesOnMap(ByVal ws As Worksheet, ByVal mapImage As Shape, _
                                ByRef candidates() As CandidateInfo, ByVal count As Integer)
    Dim i As Integer
    Dim coords As Variant
    Dim mapX As Single, mapY As Single
    Dim mapWidth As Single, mapHeight As Single
    Dim targetX As Single, targetY As Single
    Dim textBoxX As Single, textBoxY As Single
    Dim txtBox As Shape
    Dim arrow As Shape
    Dim boxColor As Long
    Dim isLeftSide As Boolean
    Dim leftPositions() As Single
    Dim rightPositions() As Single
    Dim leftCount As Integer, rightCount As Integer
    
    ' Lấy kích thước và vị trí ảnh bản đồ
    mapX = mapImage.Left
    mapY = mapImage.Top
    mapWidth = mapImage.Width
    mapHeight = mapImage.Height
    
    ' Khởi tạo mảng vị trí để tránh chồng đè
    ReDim leftPositions(1 To count)
    ReDim rightPositions(1 To count)
    leftCount = 0
    rightCount = 0
    
    ' Vẽ từng ứng viên
    For i = 1 To count
        If candidates(i).ProvinceID > 0 Then
            ' Lấy tọa độ tương đối của tỉnh
            coords = GetProvinceCoordinates(candidates(i).ProvinceID)
            
            ' Tính tọa độ tuyệt đối trên bản đồ
            ' targetX = mapImage.Left + rx * mapImage.Width
            ' targetY = mapImage.Top + ry * mapImage.Height
            targetX = mapX + (coords(0) * mapWidth)
            targetY = mapY + (coords(1) * mapHeight)
            
            ' Xác định màu sắc Text Box
            If candidates(i).ProvinceID = 1 Or candidates(i).ProvinceID = 50 Then
                boxColor = COLOR_RED ' Hà Nội và TP.HCM
            Else
                boxColor = COLOR_GREEN ' Tỉnh thường
            End If
            
            ' Smart Layout: Xác định vị trí Text Box (trái hoặc phải)
            isLeftSide = (coords(0) > 0.5) ' Nếu tỉnh ở bên phải bản đồ -> đặt Text Box bên trái
            
            If isLeftSide Then
                leftCount = leftCount + 1
                textBoxX = mapX - TEXTBOX_WIDTH - 20
                textBoxY = CalculateSmartYPosition(leftPositions, leftCount, targetY)
                leftPositions(leftCount) = textBoxY
            Else
                rightCount = rightCount + 1
                textBoxX = mapX + mapWidth + 20
                textBoxY = CalculateSmartYPosition(rightPositions, rightCount, targetY)
                rightPositions(rightCount) = textBoxY
            End If
            
            ' Tạo Text Box (đặt tên theo tiền tố MAP_ELEM_ để nhận diện khi xóa)
            Set txtBox = CreateCandidateTextBox(ws, textBoxX, textBoxY, candidates(i), boxColor)
            txtBox.Name = MAP_ELEM_PREFIX & "BOX_" & i
            
            ' Tạo mũi tên từ Text Box đến vị trí trên bản đồ (đặt tên tương tự)
            Set arrow = CreateArrow(ws, txtBox, targetX, targetY, isLeftSide)
            arrow.Name = MAP_ELEM_PREFIX & "ARROW_" & i
        End If
    Next i
End Sub

'==============================================
' FUNCTION: CalculateSmartYPosition
' MỤC ĐÍCH: Tính toán vị trí Y thông minh để tránh chồng đè Text Box
'==============================================
Private Function CalculateSmartYPosition(ByRef positions() As Single, ByVal currentCount As Integer, _
                                        ByVal targetY As Single) As Single
    Dim i As Integer
    Dim proposedY As Single
    Dim hasOverlap As Boolean
    Dim offset As Single
    
    proposedY = targetY - (TEXTBOX_HEIGHT / 2)
    offset = 0
    
    ' Kiểm tra chồng đè với các Text Box đã vẽ
    Do
        hasOverlap = False
        For i = 1 To currentCount - 1
            If Abs(proposedY - positions(i)) < MIN_VERTICAL_SPACING Then
                hasOverlap = True
                Exit For
            End If
        Next i
        
        If hasOverlap Then
            offset = offset + MIN_VERTICAL_SPACING
            proposedY = targetY - (TEXTBOX_HEIGHT / 2) + offset
            
            ' Nếu offset quá lớn, thử hướng ngược lại
            If offset > 200 Then
                offset = -offset
                proposedY = targetY - (TEXTBOX_HEIGHT / 2) + offset
            End If
        End If
    Loop While hasOverlap
    
    CalculateSmartYPosition = proposedY
End Function

'==============================================
' FUNCTION: CreateCandidateTextBox
' MỤC ĐÍCH: Tạo Text Box chứa thông tin ứng viên
'==============================================
Private Function CreateCandidateTextBox(ByVal ws As Worksheet, ByVal x As Single, _
                                        ByVal y As Single, ByRef candidate As CandidateInfo, _
                                        ByVal bgColor As Long) As Shape
    Dim txtBox As Shape
    Dim displayText As String
    
    ' Format: [STT]. [HỌ TÊN] -> [TÊN TỈNH] [省/首都/市]
    displayText = candidate.STT & ". " & candidate.HoTen & " -> " & _
                  candidate.TinhThanh & " " & candidate.LoaiDonVi
    
    ' Tạo Text Box
    Set txtBox = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, TEXTBOX_WIDTH, TEXTBOX_HEIGHT)
    
    With txtBox
        .TextFrame.Characters.Text = displayText
        
        ' Định dạng chữ
        With .TextFrame.Characters.Font
            .Name = "Arial"
            .Size = 10
            .Bold = True
            .Color = COLOR_WHITE
        End With
        
        ' Căn giữa chữ
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.MarginLeft = 3
        .TextFrame.MarginRight = 3
        .TextFrame.MarginTop = 2
        .TextFrame.MarginBottom = 2
        
        ' Định dạng viền và nền
        .Fill.ForeColor.RGB = bgColor
        .Fill.Solid
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Weight = 1
    End With
    
    Set CreateCandidateTextBox = txtBox
End Function

'==============================================
' FUNCTION: CreateArrow
' MỤC ĐÍCH: Tạo mũi tên kết nối từ Text Box đến vị trí trên bản đồ
'==============================================
Private Function CreateArrow(ByVal ws As Worksheet, ByVal txtBox As Shape, _
                            ByVal targetX As Single, ByVal targetY As Single, _
                            ByVal isLeftSide As Boolean) As Shape
    Dim arrow As Shape
    Dim startX As Single, startY As Single
    
    ' Xác định điểm xuất phát của mũi tên (từ mép của Text Box)
    If isLeftSide Then
        startX = txtBox.Left + txtBox.Width ' Mép phải của Text Box
    Else
        startX = txtBox.Left ' Mép trái của Text Box
    End If
    startY = txtBox.Top + (txtBox.Height / 2) ' Giữa Text Box theo chiều cao
    
    ' Tạo Connector (mũi tên)
    Set arrow = ws.Shapes.AddConnector(msoConnectorStraight, startX, startY, targetX, targetY)
    
    With arrow.Line
        .ForeColor.RGB = COLOR_ARROW
        .Weight = ARROW_WIDTH
        .EndArrowheadStyle = msoArrowheadTriangle
        .EndArrowheadLength = msoArrowheadLengthMedium
        .EndArrowheadWidth = msoArrowheadWidthMedium
    End With
    
    Set CreateArrow = arrow
End Function

'==============================================
' SUB: CreateSampleData (HELPER - TÙY CHỌN)
' MỤC ĐÍCH: Tạo dữ liệu mẫu để test
'==============================================
Public Sub CreateSampleData()
    Dim wsDanhSach As Worksheet
    Dim sampleData As Variant
    Dim i As Long
    
    On Error Resume Next
    Set wsDanhSach = ThisWorkbook.Worksheets("DanhSach")
    On Error GoTo 0
    
    If wsDanhSach Is Nothing Then
        Set wsDanhSach = ThisWorkbook.Worksheets.Add
        wsDanhSach.Name = "DanhSach"
    End If
    
    ' Xóa dữ liệu cũ
    wsDanhSach.Cells.Clear
    
    ' Tạo header
    With wsDanhSach
        .Cells(1, 1).Value = "STT"
        .Cells(1, 2).Value = "Họ và tên"
        .Cells(1, 3).Value = "Tỉnh/Thành"
        .Cells(1, 4).Value = "Loại đơn vị"
        
        ' Format header
        .Range("A1:D1").Font.Bold = True
        .Range("A1:D1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:D1").Font.Color = RGB(255, 255, 255)
    End With
    
    ' Dữ liệu mẫu
    sampleData = Array( _
        Array("01", "NGUYEN DIN TRUNG", "HA NOI", "首都"), _
        Array("02", "PHAM CHUNG HIEU", "HUNG YEN", "省"), _
        Array("03", "LE PHUOC QUAN", "BINH DINH", "省"), _
        Array("04", "LE HOANG THACH", "HA TINH", "省"), _
        Array("05", "TRIEU ANH MINH", "PHU THO", "省"), _
        Array("06", "LOC HOANG NHAN", "LAM DONG", "省"), _
        Array("07", "TUONG XUAN LOC", "HA NOI", "首都"), _
        Array("08", "DAM QUANG VINH", "HA NOI", "首都"), _
        Array("09", "TRIEU KHANH HUNG", "PHU THO", "省"), _
        Array("10", "HO CHI MINH", "HO CHI MINH", "市") _
    )
    
    ' Ghi dữ liệu vào sheet
    For i = LBound(sampleData) To UBound(sampleData)
        wsDanhSach.Cells(i + 2, 1).Value = sampleData(i)(0)
        wsDanhSach.Cells(i + 2, 2).Value = sampleData(i)(1)
        wsDanhSach.Cells(i + 2, 3).Value = sampleData(i)(2)
        wsDanhSach.Cells(i + 2, 4).Value = sampleData(i)(3)
    Next i
    
    ' Auto-fit columns
    wsDanhSach.Columns("A:D").AutoFit
    
    MsgBox "Đã tạo dữ liệu mẫu thành công!", vbInformation
End Sub
