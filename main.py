"""
ReadAloud v2 — Cleaner, lighter version
Nathan's TTS + OCR tool for Windows

F6  → Read clipboard aloud
F7  → OCR screenshot → read aloud
Tray icon → change voice / rate / exit
"""

import os, sys, asyncio, threading, subprocess, tempfile, ctypes
from ctypes import wintypes

# ── Dependencies ─────────────────────────────────────────────────
MISSING = []
try:
    import edge_tts
except ImportError:
    MISSING.append("edge-tts")

try:
    import mss
except ImportError:
    MISSING.append("mss")

try:
    from PIL import Image
except ImportError:
    MISSING.append("Pillow")

try:
    import win32gui, win32clipboard, win32api, win32ui
    import win32con
except ImportError:
    print("ERROR: pywin32 not installed. Run: pip install pywin32")
    sys.exit(1)

# Workaround: add missing shell notify constants if not present
for _c, _v in [("NIF_MESSAGE", 0x00000000), ("NIF_ICON", 0x00000002),
               ("NIF_TIP", 0x00000004), ("NIM_ADD", 0x00000000),
               ("NIM_MODIFY", 0x00000001), ("NIM_DELETE", 0x00000002),
               ("WS_EX_TRANSPARENT", 0x00000020), ("WS_EX_TOPMOST", 0x00000008),
               ("WS_POPUP", 0x80000000), ("WS_EX_TOOLWINDOW", 0x00000080),
               ("LWA_ALPHA", 0x00000002), ("WM_USER", 0x0400),
               ("IDI_APPLICATION", 0x7F00), ("IDC_CROSS", 0x7F02),
               ("CW_USEDEFAULT", 0x80000000), ("MF_STRING", 0x00000000),
               ("MF_POPUP", 0x00000010), ("MF_SEPARATOR", 0x00000800),
               ("MF_CHECKED", 0x00000008), ("MF_STRING", 0x00000000),
               ("TPM_LEFTALIGN", 0x0000), ("WM_NULL", 0x0000)]:
    if not hasattr(win32con, _c):
        setattr(win32con, _c, _v)
if MISSING:
    print(f"Missing packages: {', '.join(MISSING)}")
    print("Run: pip install " + " ".join(MISSING))
    sys.exit(1)

# ── Constants ───────────────────────────────────────────────────
APP      = "ReadAloud"
TMP      = tempfile.gettempdir()
AUDIO    = os.path.join(TMP, "readaloud.mp3")
OCR_IMG  = os.path.join(TMP, "readaloud_ocr.png")
LOG_FILE = os.path.join(TMP, "readaloud.log")

VOICES = {
    "sonia":   ("en-GB-SoniaNeural",   "🇬🇧 Sonia (UK Female)"),
    "guy":     ("en-GB-GuyNeural",      "🇬🇧 Guy (UK Male)"),
    "sarah":   ("en-US-SarahNeural",    "🇺🇸 Sarah (US)"),
    "jenny":   ("en-US-JennyNeural",    "🇺🇸 Jenny (US)"),
    "aria":    ("en-US-AriaNeural",     "🇺🇸 Aria (US)"),
    "davis":   ("en-US-DavisNeural",    "🇺🇸 Davis (US)"),
    "gera":    ("de-DE-SeraphinaNeural","🇩🇪 Gera (German)"),
    "lucia":   ("es-ES-LuciaNeural",    "🇪🇸 Lucia (Spanish)"),
    "hilaire": ("fr-FR-HilaireNeural",  "🇫🇷 Hilaire (French)"),
    "hoda":    ("ar-EG-HodaNeural",     "🇪🇬 Hoda (Arabic)"),
}

RATES = {
    "slow":   "-20%",
    "normal": "+0%",
    "fast":   "+20%",
}

# ── Logging ─────────────────────────────────────────────────────
def log(msg):
    ts = __import__("datetime").datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass

# ── TTS ─────────────────────────────────────────────────────────
class TTS:
    def __init__(self):
        self.voice_key = "sonia"
        self.rate_key   = "normal"
        self._proc      = None

    @property
    def voice(self) -> str:
        return VOICES[self.voice_key][0]

    def stop(self):
        try:
            if self._proc and self._proc.poll() is None:
                self._proc.terminate()
        except Exception:
            pass

    async def _synth(self, text: str):
        rate = RATES[self.rate_key]
        log(f"TTS: {self.voice}@{rate} → {text[:50]!r}...")
        await edge_tts.Communicate(text, self.voice, rate=rate).save(AUDIO)

    def speak(self, text: str, block: bool = False):
        if not text.strip():
            return
        self.stop()

        def run():
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                loop.run_until_complete(self._synth(text.strip()))
                if not os.path.exists(AUDIO):
                    log("Audio file not created")
                    return
                cmd = ["mplay32", "/play", "/close", AUDIO]
                try:
                    subprocess.run(cmd, check=True,
                                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                except FileNotFoundError:
                    # Fallback: use default player
                    if sys.platform == "win32":
                        os.startfile(AUDIO)
                    else:
                        # Linux/macOS: try mpv, then ffplay, then vlc
                        import shutil
                        for player in ["mpv", "ffplay", "vlc"]:
                            if shutil.which(player):
                                subprocess.run([player, AUDIO], check=True)
                                break
                        else:
                            # Last resort: PowerShell (Windows only)
                            subprocess.run(
                                ["powershell", "-c",
                                 f"Start-Process '{AUDIO}'"],
                                check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except Exception as e:
                log(f"TTS error: {e}")
            finally:
                loop.close()

        if block:
            run()
        else:
            threading.Thread(target=run, daemon=True).start()

# ── Clipboard ───────────────────────────────────────────────────
def clipboard_text() -> str | None:
    try:
        win32clipboard.OpenClipboard()
        if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
            data = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
            win32clipboard.CloseClipboard()
            return data
        win32clipboard.CloseClipboard()
    except Exception as e:
        log(f"Clipboard error: {e}")
    return None

# ── OCR ─────────────────────────────────────────────────────────
def windows_ocr(path: str) -> str:
    """Run Windows.Media.Ocr via PowerShell."""
    p = path.replace("\\", "\\\\")
    ps = f"""
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    $null = [Windows.Media.Ocr.OcrEngine,Windows.Media.Ocr,ContentType=WindowsRuntime]
    $null = [Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics.Imaging,ContentType=WindowsRuntime]
    $null = [Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]
    $file  = [Windows.Storage.StorageFile]::GetFileFromPathAsync('{p}').GetAwaiter().GetResult()
    $stream = $file.OpenReadAsync().GetAwaiter().GetResult()
    $dec   = [Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream).GetAwaiter().GetResult()
    $bmp   = $dec.GetAsyncWaitForResult()
    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    if ($null -eq $engine) {{ ""; exit }}
    ($engine.RecognizeAsync($bmp).GetAwaiter().GetResult()).Lines.Text
    """
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
            capture_output=True, text=True, timeout=30,
            encoding="utf-8", errors="replace")
        return r.stdout.strip()
    except Exception as e:
        log(f"OCR error: {e}")
        return ""

def capture_region(left: int, top: int, w: int, h: int) -> str:
    with mss.mss() as sct:
        sct.grab({"left": left, "top": top, "width": w, "height": h}).save(OCR_IMG)
    return OCR_IMG

# ── Region Selector (transparent overlay) ──────────────────────
class Selector:
    def __init__(self):
        self.rect     = None
        self.dragging = False
        self.start    = (0, 0)
        self._hwnd    = None

    def _wnd(self, h, m, w, l):
        MA  = 0x0201  # WM_LBUTTONDOWN
        MR  = 0x0202  # WM_LBUTTONUP
        MM  = 0x0200  # WM_MOUSEMOVE
        PK  = 0x0014  # WM_ERASEBKGND
        PT  = win32con.WM_PAINT
        DW  = win32con.WM_DESTROY

        if m == MA:
            self.dragging = True
            self.start = (win32api.LoWord(l), win32api.HiWord(l))
            win32gui.SetCapture(h)
            return 0
        if m == MR and self.dragging:
            end = (win32api.LoWord(l), win32api.HiWord(l))
            x1, y1 = self.start; x2, y2 = end
            self.rect = (min(x1,x2), min(y1,y2), max(x1,x2), max(y1,y2))
            self.dragging = False
            win32gui.ReleaseCapture()
            win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
            return 0
        if m == MM and self.dragging:
            win32gui.InvalidateRect(h, None, True)
            return 0
        if m == PT:
            self._paint()
            return 0
        if m == DW:
            win32gui.PostQuitMessage(0)
            return 0
        return win32gui.DefWindowProc(h, m, w, l)

    def _paint(self):
        if not self.dragging:
            return
        x1, y1 = self.start
        x2, y2 = win32api.GetCursorPos()
        dc = win32gui.GetDC(self._hwnd)
        r2 = win32con.R2_NOTXORPEN
        win32gui.SetROP2(dc, r2)
        pen = win32gui.CreatePen(win32con.PS_SOLID, 2, 0x00FF00)
        old = win32gui.SelectObject(dc, pen)
        win32gui.SelectObject(dc, win32gui.GetStockObject(win32con.NULL_BRUSH))
        win32gui.Rectangle(dc, x1, y1, x2, y2)
        win32gui.SelectObject(dc, old)
        win32gui.DeleteObject(pen)
        win32gui.ReleaseDC(self._hwnd, dc)

    def select(self):
        sw = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
        sh = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
        CN = f"{APP}Overlay"
        wc = win32gui.WNDCLASS()
        wc.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
        wc.lpfnWndProc   = self._wnd
        wc.hCursor       = win32gui.LoadCursor(0, win32con.IDC_CROSS)
        wc.lpszClassName = CN
        try:
            win32gui.RegisterClass(wc)
        except Exception:
            pass
        EX = win32con.WS_EX_TRANSPARENT | win32con.WS_EX_TOPMOST | win32con.WS_EX_LAYERED
        ST = win32con.WS_POPUP | win32con.WS_EX_TOOLWINDOW
        self._hwnd = win32gui.CreateWindowEx(EX, CN, APP, ST, 0, 0, sw, sh, 0, 0, 0, None)
        win32gui.SetLayeredWindowAttributes(self._hwnd, 0, 160, win32con.LWA_ALPHA)
        win32gui.ShowWindow(self._hwnd, win32con.SW_SHOW)
        win32gui.SetForegroundWindow(self._hwnd)
        _u = ctypes.windll.user32
        msg = wintypes.MSG()
        pmsg = ctypes.byref(msg)
        while _u.GetMessageW(pmsg, None, 0, 0) != 0:
            _u.TranslateMessage(pmsg)
            _u.DispatchMessageW(pmsg)
            if self.rect:
                break
        try:
            win32gui.DestroyWindow(self._hwnd)
        except Exception:
            pass
        return self.rect

# ── System Tray ─────────────────────────────────────────────────
def setup_tray(tts: TTS):
    hinst = win32api.GetModuleHandle(None)

    # Hidden message window
    WC = win32gui.WNDCLASS()
    WC.lpfnWndProc   = lambda h, m, w, l: 0
    WC.lpszClassName = f"{APP}TrayHost"
    WC.hInstance     = hinst
    try:
        win32gui.RegisterClass(WC)
    except Exception:
        pass
    hwnd = win32gui.CreateWindow(WC.lpszClassName, APP, 0, 0, 0, 0, 0, 0, 0, hinst, None)

    # Build menu
    hmenu = win32gui.CreateMenu()
    win32gui.AppendMenu(hmenu, win32con.MF_STRING, 1,  "▶  Speak Clipboard (F6)")
    win32gui.AppendMenu(hmenu, win32con.MF_STRING, 2,  "🖼  OCR Region (F7)")
    win32gui.AppendMenu(hmenu, win32con.MF_SEPARATOR, 0, "")
    win32gui.AppendMenu(hmenu, win32con.MF_STRING, 3,  "⏹  Stop")
    win32gui.AppendMenu(hmenu, win32con.MF_SEPARATOR, 0, "")

    # Voices submenu
    sub_v = win32gui.CreateMenu()
    for i, (k, (_, label)) in enumerate(VOICES.items()):
        mid = 10 + i
        check = win32con.MF_CHECKED if tts.voice_key == k else win32con.MF_STRING
        win32gui.AppendMenu(sub_v, check, mid, label)
    win32gui.AppendMenu(hmenu, win32con.MF_POPUP, sub_v, "🎤 Voice")

    # Rate submenu
    sub_r = win32gui.CreateMenu()
    for i, (k, label) in enumerate(RATES.items()):
        mid = 30 + i
        check = win32con.MF_CHECKED if tts.rate_key == k else win32con.MF_STRING
        win32gui.AppendMenu(sub_r, check, mid, f"{label} ({k})")
    win32gui.AppendMenu(hmenu, win32con.MF_POPUP, sub_r, "⚡ Speed")

    win32gui.AppendMenu(hmenu, win32con.MF_SEPARATOR, 0, "")
    win32gui.AppendMenu(hmenu, win32con.MF_STRING, 99, "❌  Exit")

    # NID tuple: (hwnd, id, flags, msg, hicon, tip)
    hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
    nid = (hwnd, 0, win32con.NIF_MESSAGE | win32con.NIF_ICON | win32con.NIF_TIP,
           win32con.WM_USER + 1, hicon, APP)

    win32gui.Shell_NotifyIcon(win32con.NIM_ADD, nid)

    # Tray message loop — pure ctypes to avoid pywintypes mismatch
    _user32 = ctypes.windll.user32

    def tray_loop():
        msg = wintypes.MSG()
        pmsg = ctypes.byref(msg)
        while _user32.GetMessageW(pmsg, None, 0, 0) != 0:
            if msg.message == 0x0111:  # WM_COMMAND
                wid = msg.wParam & 0xFFFF
                if wid == 1:
                    text = clipboard_text()
                    if text:
                        tts.speak(text)
                elif wid == 2:
                    sel = Selector()
                    r = sel.select()
                    if r:
                        x1,y1,x2,y2 = r
                        if x2-x1 > 5 and y2-y1 > 5:
                            p = capture_region(x1, y1, x2-x1, y2-y1)
                            text = windows_ocr(p)
                            if text:
                                tts.speak(text, block=True)
                            else:
                                log("OCR: no text found")
                elif wid == 3:
                    tts.stop()
                elif wid == 99:
                    win32gui.Shell_NotifyIcon(win32con.NIM_DELETE, nid)
                    win32gui.PostQuitMessage(0)
                elif 10 <= wid < 10 + len(VOICES):
                    key = list(VOICES.keys())[wid - 10]
                    tts.voice_key = key
                    log(f"Voice: {VOICES[key][1]}")
                elif 30 <= wid < 30 + len(RATES):
                    key = list(RATES.keys())[wid - 30]
                    tts.rate_key = key
                    log(f"Rate: {key}")
            elif msg.message == win32con.WM_USER + 1:
                if msg.lparam == win32con.WM_RBUTTONUP:
                    cx, cy = win32api.GetCursorPos()
                    win32gui.SetForegroundWindow(hwnd)
                    win32gui.TrackPopupMenu(hmenu, win32con.TPM_LEFTALIGN,
                                            cx, cy, 0, hwnd, None)
                    win32gui.PostMessage(hwnd, win32con.WM_NULL, 0, 0)
            else:
                _user32.TranslateMessage(pmsg)
                _user32.DispatchMessageW(pmsg)

    t = threading.Thread(target=tray_loop, daemon=True)
    t.start()
    return nid, hwnd

# ── Hotkey Pump (pure ctypes) ───────────────────────────────────
user32 = ctypes.windll.user32
WM_HOTKEY, HOTKEY1, HOTKEY2 = 0x0312, 1, 2
F6, F7 = 0x75, 0x76  # VK_F6, VK_F7

def hotkey_loop(tts: TTS, tray_hwnd):
    # Hidden window for hotkey messages
    hinst = win32api.GetModuleHandle(None)
    WC = win32gui.WNDCLASS()
    WC.lpfnWndProc   = lambda h, m, w, l: 0
    WC.lpszClassName = f"{APP}HotkeyWin"
    WC.hInstance     = hinst
    try:
        win32gui.RegisterClass(WC)
    except Exception:
        pass
    hwnd = win32gui.CreateWindow(WC.lpszClassName, APP, 0, 0, 0, 0, 0, 0, 0, hinst, None)

    user32.RegisterHotKey(hwnd, HOTKEY1, 0, F6)
    user32.RegisterHotKey(hwnd, HOTKEY2, 0, F7)
    log("Ready! F6=Speak, F7=OCR")

    msg = wintypes.MSG()
    pmsg = ctypes.byref(msg)
    while user32.GetMessageW(pmsg, None, 0, 0) != 0:
        if msg.message == WM_HOTKEY:
            if msg.wParam == HOTKEY1:
                text = clipboard_text()
                if text:
                    log(f"F6 → {len(text)} chars")
                    tts.speak(text)
                else:
                    log("F6: clipboard empty")
            elif msg.wParam == HOTKEY2:
                log("F7: selecting region...")
                sel = Selector()
                r = sel.select()
                if r:
                    x1,y1,x2,y2 = r
                    if x2-x1 > 5 and y2-y1 > 5:
                        p = capture_region(x1, y1, x2-x1, y2-y1)
                        text = windows_ocr(p)
                        if text:
                            log(f"F7 OCR → {len(text)} chars")
                            tts.speak(text, block=True)
                        else:
                            log("F7: no text found")
                else:
                    log("F7: cancelled")
        else:
            # Dispatch non-hotkey messages via ctypes
            user32.TranslateMessage(pmsg)
            user32.DispatchMessageW(pmsg)

# ── Main ────────────────────────────────────────────────────────
if __name__ == "__main__":
    log(f"{APP} v2 starting...")
    tts = TTS()
    setup_tray(tts)
    hotkey_loop(tts, None)
