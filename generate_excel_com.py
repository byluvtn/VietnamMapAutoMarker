"""
Generate BanDoUngVien.xlsm using native Excel COM automation (win32com).
This guarantees a genuinely valid macro-enabled .xlsm file because Excel
itself writes the file (unlike openpyxl which cannot create a real
vbaProject.bin stream).

All console output is ASCII-only to avoid Windows codepage crashes.
"""
import os
import sys
import time
import subprocess
import threading

import win32com.client as win32
import pythoncom

DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
PROJECT_DIR = os.path.join(DESKTOP, "VietnamMapAutoMarker")
IMAGE_PATH = os.path.join(PROJECT_DIR, "BanDoVN.png")
BAS_PATH = os.path.join(PROJECT_DIR, "MapGenerator.bas")
OUTPUT_PATH = os.path.join(DESKTOP, "BanDoUngVien.xlsm")

EXCEL_EXE = r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"

xlOpenXMLWorkbookMacroEnabled = 52
vbext_ct_StdModule = 1

SAMPLE_DATA = [
    ("01", "NGUYEN DIN TRUNG", "HA NOI", "\u9996\u90fd"),       # 首都
    ("02", "PHAM CHUNG HIEU", "HUNG YEN", "\u7701"),            # 省
    ("03", "LE PHUOC QUAN", "BINH DINH", "\u7701"),
    ("04", "LE HOANG THACH", "HA TINH", "\u7701"),
    ("05", "TRIEU ANH MINH", "PHU THO", "\u7701"),
    ("06", "LOC HOANG NHAN", "LAM DONG", "\u7701"),
    ("07", "TUONG XUAN LOC", "HA NOI", "\u9996\u90fd"),
    ("08", "DAM QUANG VINH", "HA NOI", "\u9996\u90fd"),
    ("09", "TRIEU KHANH HUNG", "PHU THO", "\u7701"),
    ("10", "HO CHI MINH", "HO CHI MINH", "\u5e02"),             # 市
]

HEADERS = ["STT", "Ho va ten", "Tinh/Thanh", "Loai don vi"]


def log(msg):
    print(msg)
    sys.stdout.flush()


def ensure_excel_process_running():
    """Pre-launch EXCEL.EXE /automation as a background COM server using
    subprocess.Popen with an argument list. This is required in this
    environment: dispatching Excel.Application via COM alone is not
    reliable unless an /automation process is already alive first."""
    try:
        subprocess.Popen([EXCEL_EXE, "/automation"])
    except Exception as e:
        log("WARNING: could not pre-launch EXCEL.EXE: %s" % e)
    time.sleep(3)


def dispatch_excel_with_retry(max_attempts=10, delay=2):
    last_err = RuntimeError("dispatch_excel_with_retry: max_attempts must be >= 1")
    for attempt in range(1, max_attempts + 1):
        try:
            return win32.Dispatch("Excel.Application")
        except Exception as e:
            last_err = e
            log("      Dispatch attempt %d/%d failed, retrying..." % (attempt, max_attempts))
            time.sleep(delay)
    raise last_err


def read_vba_code():
    with open(BAS_PATH, "r", encoding="utf-8") as f:
        lines = f.readlines()
    # Strip the Attribute VB_Name line - name is set via vbcomp.Name instead
    lines = [l for l in lines if not l.strip().startswith("Attribute VB_Name")]
    return "".join(lines)


def build_workbook(excel):
    wb = excel.Workbooks.Add()

    # Normalize sheets: first sheet -> DanhSach, add BanDo after it,
    # delete any extra default sheets.
    ws_danh = wb.Sheets(1)
    ws_danh.Name = "DanhSach"

    ws_bando = wb.Sheets.Add(After=ws_danh)
    ws_bando.Name = "BanDo"

    excel.DisplayAlerts = False
    for i in range(wb.Sheets.Count, 0, -1):
        s = wb.Sheets(i)
        if s.Name not in ("DanhSach", "BanDo"):
            s.Delete()
    excel.DisplayAlerts = True

    # --- DanhSach sheet ---
    for col, header in enumerate(HEADERS, start=1):
        ws_danh.Cells(1, col).Value = header
    header_range = ws_danh.Range("A1:D1")
    header_range.Font.Bold = True
    header_range.Interior.Color = 0xC47244  # BGR for RGB(68,114,196)
    header_range.Font.Color = 0xFFFFFF

    for row_idx, row in enumerate(SAMPLE_DATA, start=2):
        for col_idx, value in enumerate(row, start=1):
            ws_danh.Cells(row_idx, col_idx).Value = value

    ws_danh.Columns("A:D").AutoFit()

    # --- BanDo sheet: insert map image ---
    shp = ws_bando.Shapes.AddPicture(
        IMAGE_PATH, 0, 1, 20, 50, -1, -1
    )
    # Cap size so the map fits reasonably on screen
    shp.LockAspectRatio = 1
    if shp.Width > 900:
        shp.Width = 900
    shp.Name = "BanDoVN"

    # --- Button wired to GenerateVietnamMap ---
    btn = ws_bando.Buttons().Add(20, 10, 160, 30)
    btn.OnAction = "GenerateVietnamMap"
    btn.Characters.Text = "TAO BAN DO"

    return wb, ws_danh, ws_bando


def inject_vba(wb):
    vbcomp = wb.VBProject.VBComponents.Add(vbext_ct_StdModule)
    vbcomp.Name = "MapGenerator"
    code = read_vba_code()
    vbcomp.CodeModule.AddFromString(code)
    return vbcomp


def auto_dismiss_dialogs(stop_event):
    """Background thread: periodically press Enter to dismiss any VBA
    MsgBox popups that would otherwise block automation. Tries several
    possible window title variants since the MsgBox title can differ
    (app title, workbook name, or default "Microsoft Excel")."""
    pythoncom.CoInitialize()
    try:
        shell = win32.Dispatch("WScript.Shell")
        titles = ["Microsoft Excel", "Excel"]
        while not stop_event.is_set():
            time.sleep(0.3)
            for title in titles:
                try:
                    activated = shell.AppActivate(title)
                except Exception:
                    activated = False
                if activated:
                    try:
                        shell.SendKeys("{ENTER}")
                    except Exception:
                        pass
                    break
    finally:
        pythoncom.CoUninitialize()


def test_generated_macro(excel, wb):
    """Run GenerateVietnamMap once to verify it executes without a VBA
    runtime error, auto-dismissing its completion MsgBox."""
    stop_event = threading.Event()
    t = threading.Thread(target=auto_dismiss_dialogs, args=(stop_event,), daemon=True)
    t.start()
    try:
        excel.Run("GenerateVietnamMap")
        ok = True
        err = None
    except Exception as e:
        ok = False
        err = str(e)
    finally:
        stop_event.set()
        t.join(timeout=2)
    return ok, err


def verify_saved_file():
    """Reopen the saved file fresh and confirm structure is intact."""
    excel2 = dispatch_excel_with_retry()
    excel2.Visible = False
    excel2.DisplayAlerts = False
    try:
        wb2 = excel2.Workbooks.Open(OUTPUT_PATH)
    except Exception as e:
        excel2.Quit()
        return False, ["Cannot reopen saved file: %s" % e]

    problems = []
    sheet_names = [wb2.Sheets(i).Name for i in range(1, wb2.Sheets.Count + 1)]
    if "DanhSach" not in sheet_names:
        problems.append("Missing sheet DanhSach")
    if "BanDo" not in sheet_names:
        problems.append("Missing sheet BanDo")

    shape_found = False
    if "BanDo" in sheet_names:
        ws = wb2.Sheets("BanDo")
        for i in range(1, ws.Shapes.Count + 1):
            if ws.Shapes(i).Name == "BanDoVN":
                shape_found = True
                break
    if not shape_found:
        problems.append("Missing shape BanDoVN on sheet BanDo")

    try:
        comp_names = [wb2.VBProject.VBComponents(i).Name
                      for i in range(1, wb2.VBProject.VBComponents.Count + 1)]
        if "MapGenerator" not in comp_names:
            problems.append("Missing VBA module MapGenerator")
    except Exception as e:
        problems.append("Cannot access VBProject: %s" % e)

    wb2.Close(False)
    excel2.Quit()
    return (len(problems) == 0), problems


def main():
    pythoncom.CoInitialize()

    if not os.path.exists(IMAGE_PATH):
        log("ERROR: map image not found at %s" % IMAGE_PATH)
        sys.exit(1)
    if not os.path.exists(BAS_PATH):
        log("ERROR: VBA source not found at %s" % BAS_PATH)
        sys.exit(1)

    if os.path.exists(OUTPUT_PATH):
        try:
            os.remove(OUTPUT_PATH)
        except Exception as e:
            log("ERROR: cannot remove old output file: %s" % e)
            sys.exit(1)

    log("[1/6] Starting Excel...")
    ensure_excel_process_running()
    excel = dispatch_excel_with_retry()
    excel.Visible = True
    excel.DisplayAlerts = False

    log("[2/6] Building workbook (sheets, data, image, button)...")
    wb, ws_danh, ws_bando = build_workbook(excel)

    log("[3/6] Injecting VBA module...")
    inject_vba(wb)

    log("      Saving as macro-enabled workbook...")
    wb.SaveAs(OUTPUT_PATH, FileFormat=xlOpenXMLWorkbookMacroEnabled)

    # A VBA module injected via CodeModule.AddFromString in the same session
    # is not reliably callable via Application.Run until the workbook is
    # closed and reopened. Close and reopen before testing the macro.
    log("[4/6] Closing and reopening workbook to activate injected macro...")
    wb.Close(False)
    wb = excel.Workbooks.Open(OUTPUT_PATH)

    log("[5/6] Testing GenerateVietnamMap macro...")
    ok, err = test_generated_macro(excel, wb)
    if ok:
        log("      OK: macro ran without error")
    else:
        log("      FAILED: macro raised an error: %s" % err)

    wb.Save()
    wb.Close(False)
    excel.Quit()
    time.sleep(1)

    log("[6/6] Verifying saved file by reopening it...")
    passed, problems = verify_saved_file()

    if passed and ok:
        log("VERIFICATION PASSED: file opens correctly, all parts exist, macro runs OK.")
        log("Output file: %s" % OUTPUT_PATH)
    else:
        log("VERIFICATION FAILED:")
        if not ok:
            log("  - Macro test failed: %s" % err)
        for p in problems:
            log("  - %s" % p)
        sys.exit(1)


if __name__ == "__main__":
    main()
