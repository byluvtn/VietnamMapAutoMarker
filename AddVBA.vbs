' ============================================================
' Auto Add VBA Code to Excel File
' ============================================================
' Script tự động nhúng mã VBA từ MapGenerator.bas vào BanDoUngVien.xlsm
' và tạo nút Button để chạy macro
' ============================================================

Option Explicit

Dim xlApp, xlBook, vbComp, vbaCode, fso, file
Dim excelPath, vbaPath, ws, btn, shp
Dim desktopPath, scriptPath

' Lấy đường dẫn Desktop
Set fso = CreateObject("Scripting.FileSystemObject")
desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)

' Đường dẫn file
excelPath = desktopPath & "\BanDoUngVien.xlsm"
vbaPath = scriptPath & "\MapGenerator.bas"

' Kiểm tra file tồn tại
If Not fso.FileExists(excelPath) Then
    MsgBox "Không tìm thấy file Excel: " & vbCrLf & excelPath, vbCritical, "Lỗi"
    WScript.Quit
End If

If Not fso.FileExists(vbaPath) Then
    MsgBox "Không tìm thấy file VBA: " & vbCrLf & vbaPath, vbCritical, "Lỗi"
    WScript.Quit
End If

On Error Resume Next

' Tạo Excel Application
Set xlApp = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Không thể khởi động Excel!" & vbCrLf & Err.Description, vbCritical, "Lỗi"
    WScript.Quit
End If

xlApp.Visible = True
xlApp.DisplayAlerts = False

' Mở file Excel
Set xlBook = xlApp.Workbooks.Open(excelPath)
If Err.Number <> 0 Then
    MsgBox "Không thể mở file Excel!" & vbCrLf & Err.Description, vbCritical, "Lỗi"
    xlApp.Quit
    WScript.Quit
End If

' Kiểm tra trust settings
On Error Resume Next
Dim testAccess
testAccess = xlBook.VBProject.VBComponents.Count
If Err.Number <> 0 Then
    MsgBox "Cần bật 'Trust access to VBA project object model':" & vbCrLf & vbCrLf & _
           "1. File → Options → Trust Center" & vbCrLf & _
           "2. Trust Center Settings" & vbCrLf & _
           "3. Macro Settings → " & vbCrLf & _
           "   ✓ Trust access to the VBA project object model" & vbCrLf & vbCrLf & _
           "Sau đó chạy lại script này.", vbExclamation, "Cài đặt Trust Center"
    xlBook.Close False
    xlApp.Quit
    WScript.Quit
End If
On Error GoTo 0

' Đọc mã VBA
Set file = fso.OpenTextFile(vbaPath, 1, False, -1)
vbaCode = file.ReadAll
file.Close

' Thêm Module với VBA code
Set vbComp = xlBook.VBProject.VBComponents.Add(1)
vbComp.CodeModule.AddFromString vbaCode
vbComp.Name = "MapGenerator"

' Đặt tên cho ảnh bản đồ
On Error Resume Next
Set ws = xlBook.Worksheets("BanDo")
If Not ws Is Nothing Then
    Dim imgCount, i
    imgCount = ws.Shapes.Count
    
    For i = 1 To imgCount
        Set shp = ws.Shapes(i)
        If shp.Type = 13 Then ' msoLinkedPicture or msoPicture
            shp.Name = "BanDoVN"
            Exit For
        End If
    Next
    
    ' Tạo Button
    Dim btnLeft, btnTop, btnWidth, btnHeight
    btnLeft = 50
    btnTop = 20
    btnWidth = 150
    btnHeight = 35
    
    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
    btn.OnAction = "GenerateVietnamMap"
    btn.Text = "TẠO BẢN ĐỒ"
    btn.Font.Size = 12
    btn.Font.Bold = True
End If
On Error GoTo 0

' Lưu và thông báo
xlBook.Save

MsgBox "Đã thêm VBA thành công!" & vbCrLf & vbCrLf & _
       "✓ Module: MapGenerator" & vbCrLf & _
       "✓ Ảnh bản đồ: BanDoVN" & vbCrLf & _
       "✓ Nút: TẠO BẢN ĐỒ" & vbCrLf & vbCrLf & _
       "Vào sheet 'BanDo' và bấm nút để test!", vbInformation, "Hoàn thành"

' Cleanup
xlBook.Close True
xlApp.Quit

Set btn = Nothing
Set shp = Nothing
Set ws = Nothing
Set vbComp = Nothing
Set xlBook = Nothing
Set xlApp = Nothing
Set file = Nothing
Set fso = Nothing

WScript.Echo "Done!"
