import sys
import time
import os
import json
import threading
import winsound
import openvr
import math
import struct
import io
import requests
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from pythonosc import udp_client

# --- Configuration & State ---
SETTINGS_FILE = "settings.json"
DEFAULT_SETTINGS = {
    "wpmDot": 15,
    "wpmDash": 20,
    "freq": 900,
    "playbackGap": 20,
    "dashRatio": 3.0,
    "keyRepeat": True,
    "memorizeMode": False, # Legacy
    "charTimeout": 0.3,
    "triggerThreshold": 0.9,
    "oscInterval": 1.5,
    "requireStickDown": True,
    "debugMode": False,
    # Custom Morse Definitions
    "customMorseMap": {
        ".-.-.-": "、",
        ".-.-": "ー",
        "..--": "！",
        "---.": "？"
    }
}

# Global state wrapper
class AppState:
    def __init__(self):
        self.settings = DEFAULT_SETTINGS.copy()
        # Ensure deep copy for dictionary
        self.settings["customMorseMap"] = DEFAULT_SETTINGS["customMorseMap"].copy()
        self.unit_dot = 60
        self.unit_dash = 60
        self.trigger_threshold = 0.5
        self.char_timeout = 1.0
        
        # Audio
        self.sound_dot = None
        self.sound_dash = None
        
        # Runtime flags
        self.running = False # VR Loop running
        self.vr_initialized = False
        
        # Logic state
        self.last_char_confirmed_time = 0
        self.current_symbol_sequence = ""
        
        # Text Buffers
        self.fixed_text = ""   # Converted Japanese text
        self.text_buffer = ""  # Current Raw Romaji Input
        
        self.last_repeat_time = 0
        self.next_repeat_allowed_at = 0
        self.keys_down = {'dot': False, 'dash': False}
        
        # Input states for edge detection
        self.was_stick_down = False
        self.was_grip_down = False
        self.was_left_grip_down = False
        
        # Conversion Logic
        self.conversion_active = False # Are we currently selecting a candidate?
        self.conversion_request_pending = False # Is a network request flying?
        self.conversion_candidates = [] # List of strings
        self.conversion_index = 0
        
        self.osc = None # Initialized below

state = AppState()

class OSCManager:
    def __init__(self, ip, port, app_state):
        self.client = udp_client.SimpleUDPClient(ip, port)
        self.state = app_state
        self.last_sent_time = 0
        self.pending_update = None # (address, args)
    
    def send(self, address, args):
        interval = float(self.state.settings.get("oscInterval", 1.5))
        now = time.time()
        
        # If enough time passed, send immediately
        if (now - self.last_sent_time) >= interval:
            self.force_send(address, args)
        else:
            # Otherwise queue it (overwrite ensures only latest is pending)
            self.pending_update = (address, args)
            
    def force_send(self, address, args):
        try:
            self.client.send_message(address, args)
        except: pass
        self.last_sent_time = time.time()
        self.pending_update = None
        
    def process_queue(self):
        # Check if we can release pending message
        if self.pending_update:
            interval = float(self.state.settings.get("oscInterval", 1.5))
            now = time.time()
            if (now - self.last_sent_time) >= interval:
                addr, args = self.pending_update
                self.force_send(addr, args)

state.osc = OSCManager("127.0.0.1", 9000, state)

def generate_overlay_text():
    """Generates (main_text_line, candidates_context_line) for overlay."""
    candidates_str = None
    
    if state.conversion_active:
        # Main Line: Fixed Text + [SelectedCandidate] or Placeholder
        if state.conversion_candidates:
            c = state.conversion_candidates[state.conversion_index]
            main_text = f"{state.fixed_text}[{c}]"
            
            # Candidates Line: Show neighbor candidates
            # Format:  cand1  [cand2]  cand3  cand4
            # Let's show a window of candidates around the index
            start = max(0, state.conversion_index - 2)
            end = min(len(state.conversion_candidates), start + 5)
            
            subset = []
            for i in range(start, end):
                val = state.conversion_candidates[i]
                if i == state.conversion_index:
                    subset.append(f"[{val}]")
                else:
                    subset.append(val)
            candidates_str = " ".join(subset)
            
        else:
            main_text = f"{state.fixed_text}[?]"
            candidates_str = "No Candidates"
    else:
        # Main Line: Fixed + Buffer
        # If text is too long, we want to see the END (cursor).
        full = f"{state.fixed_text}{state.text_buffer}"
        if len(full) > 20: # Heuristic limit for visual clarity
             main_text = "..." + full[-20:]
        else:
             main_text = full
        candidates_str = "" # Clear candidates line

    morse_view = state.current_symbol_sequence if state.current_symbol_sequence else ""
    return main_text, candidates_str, morse_view

# --- Helpers ---
def generate_wav_bytes(freq, duration_ms):
    sample_rate = 44100
    num_samples = int(sample_rate * (duration_ms / 1000.0))
    amplitude = 16000
    
    buf = io.BytesIO()
    buf.write(b'RIFF')
    buf.write(struct.pack('<I', 36 + num_samples * 2))
    buf.write(b'WAVEfmt ')
    buf.write(struct.pack('<I', 16))
    buf.write(struct.pack('<H', 1)) 
    buf.write(struct.pack('<H', 1)) 
    buf.write(struct.pack('<I', sample_rate))
    buf.write(struct.pack('<I', sample_rate * 2))
    buf.write(struct.pack('<H', 2)) 
    buf.write(struct.pack('<H', 16))
    buf.write(b'data')
    buf.write(struct.pack('<I', num_samples * 2))
    
    for i in range(num_samples):
        t = float(i) / sample_rate
        value = int(amplitude * math.sin(2.0 * math.pi * freq * t))
        if i > num_samples - 100:
             fade = (num_samples - i) / 100.0
             value = int(value * fade)
        if i < 100:
             fade = i / 100.0
             value = int(value * fade)
        buf.write(struct.pack('<h', value))
        
    return buf.getvalue()

# --- Overlay Manager (Dependency Free) ---
class OverlayManager:
    def __init__(self, key, name, app_state, log_func=None):
        self._log_func = log_func if log_func else print
        self.vr_overlay = None
        self.handle = None
        self.state = app_state
        self.temp_file = os.path.abspath("overlay_temp.png") # Traditionally png, but openvr supports raw data too. 
        # Actually OpenVR SetOverlayFromFile supports PNG/JPG. Raw bytes via SetOverlayRaw is harder from Python without ctypes structures.
        # But we want NO dependencies. Writing a BMP is easy in pure Python.
        self.temp_file = os.path.abspath("overlay_temp.bmp")
        
        self.last_text = None
        self.enabled = False
        
        # Visibility Logic
        self.is_active = False # Logic state (Stick down?)
        self.last_active_time = 0
        self.visibility_timeout = 3.0 # Seconds to wait before hiding
        self.currently_visible = False # Actual VR state
        
        # Simple Bitmap Font (Accessing Windows Font is hard without libraries)
        # So we will use a very basic pixel font defined here or just render text to BMP?
        # Rendering text to BMP in pure python is complex.
        # BUT: We can use `tkinter` (which is already imported) to render text to an image!
        # Tkinter has 'Canvas' which can save to postscript, but not easily to BMP without Ghostscript.
        
        # RESET STRATEGY:
        # We will use valid BMP header generation + a simple 8x8 bitmap font included in code
        # OR better: use the existing Tkinter window to render text to a hidden canvas, catch the pixels? 
        # No, that's flaky.
        
        # Let's go with a lightweight BMP generator + basic pixel font.
        # For Japanese support without libraries... that is extremely hard (Kanji rendering requires a font engine).
        
        # ALTERNATIVE:
        # Application is already using `tkinter`.
        # We can take a screenshot of a specific Label widget? No.
        
        # WAIT. We can use standard Windows APIs via `ctypes` (built-in) to render text to a bitmap context.
        # This is strictly Windows-only but the user is on Windows.
        self.use_gdi = False
        if os.name == 'nt':
            self.use_gdi = True
            import ctypes
            from ctypes import wintypes
            
            self.gdi32 = ctypes.windll.gdi32
            self.user32 = ctypes.windll.user32
            
            # Constants
            self.FW_BOLD = 700
            self.ANSI_CHARSET = 0
            self.DEFAULT_PITCH = 0
            self.OUT_DEFAULT_PRECIS = 0
            self.CLIP_DEFAULT_PRECIS = 0
            self.DEFAULT_QUALITY = 0
            self.FF_DONTCARE = 0
        
        self.log_debug(f"Overlay: overlayEnabled={self.state.settings.get('overlayEnabled', False)}, use_gdi={self.use_gdi}")
        
        if self.state.settings.get("overlayEnabled", False):
            try:
                self.vr_overlay = openvr.IVROverlay()
                self.log_debug("Overlay: IVROverlay取得成功")
                
                # Check if overlay already exists (from previous crash/incomplete shutdown)
                existing_handle = None
                try:
                    result = self.vr_overlay.findOverlay(key)
                    self.log_debug(f"Overlay: findOverlay result = {result}")
                    
                    # Handle different return formats
                    if isinstance(result, tuple):
                        # (handle, error) format - pyopenvr returns tuple
                        existing_handle = result[0] if result[0] and result[0] != 0 else None
                    elif isinstance(result, int) and result != 0:
                        existing_handle = result
                    elif result is not None and result != 0:
                        existing_handle = result
                        
                except openvr.error_code.OverlayError_UnknownOverlay:
                    self.log_debug("Overlay: 既存オーバーレイなし")
                except Exception as find_err:
                    # Any other error during findOverlay
                    self.log_debug(f"Overlay: findOverlay例外: {type(find_err).__name__}: {find_err}")
                
                if existing_handle:
                    self.log(f"Overlay: 既存オーバーレイを破棄中 handle={existing_handle}")
                    try:
                        self.vr_overlay.destroyOverlay(existing_handle)
                        self.log_debug("Overlay: 既存オーバーレイ破棄完了")
                        # Wait a bit for OpenVR to fully release the overlay
                        time.sleep(0.1)
                    except Exception as destroy_err:
                        self.log(f"Overlay: destroyOverlay失敗: {destroy_err}")
                
                self.handle = self.vr_overlay.createOverlay(key, name)
                self.log_debug(f"Overlay: 新規作成, handle={self.handle}")
                
                self.vr_overlay.setOverlayWidthInMeters(self.handle, self.state.settings.get("overlayWidth", 1.0))
                self.vr_overlay.setOverlayColor(self.handle, 1.0, 1.0, 1.0)
                self.vr_overlay.setOverlayAlpha(self.handle, self.state.settings.get("overlayOpacity", 0.8))
                
                self.enabled = True
                self.update_transform()
                self.update_image("Ready")
                self.vr_overlay.hideOverlay(self.handle) # Start hidden (only show on stick down)
                self.log(f"Overlay: 初期化完了, enabled={self.enabled}")
            except Exception as e:
                import traceback
                self.log(f"Overlay Init Failed: {e}")
                self.log(traceback.format_exc())
                self.enabled = False
        else:
            self.log("Overlay: overlayEnabled=Falseのため初期化スキップ")

    def log(self, msg):
        """Always show important logs."""
        self._log_func(msg)
    
    def log_debug(self, msg):
        """Only show if debugMode is enabled."""
        if self.state.settings.get("debugMode", False):
            self._log_func(msg)

    def update_transform(self):
        if not self.enabled: return
        try:
            z = self.state.settings.get("overlayZ", -1.5)
            y = self.state.settings.get("overlayY", -0.2)
            
            self.log_debug(f"Overlay: update_transform z={z}, y={y}")
            
            transform = openvr.HmdMatrix34_t()
            transform.m[0][0] = 1.0; transform.m[0][1] = 0.0; transform.m[0][2] = 0.0; transform.m[0][3] = 0.0
            transform.m[1][0] = 0.0; transform.m[1][1] = 1.0; transform.m[1][2] = 0.0; transform.m[1][3] = y
            transform.m[2][0] = 0.0; transform.m[2][1] = 0.0; transform.m[2][2] = 1.0; transform.m[2][3] = z
            
            self.vr_overlay.setOverlayTransformTrackedDeviceRelative(self.handle, openvr.k_unTrackedDeviceIndex_Hmd, transform)
            self.log_debug("Overlay: setOverlayTransformTrackedDeviceRelative成功")
        except Exception as e:
            self.log(f"Overlay: update_transform失敗: {e}")

    def create_text_bitmap(self, text, candidates_context=None, morse_text=None, width=1024, height=300):
        # uses GDI to render text. Modern Glass-like Style.
        import ctypes
        from ctypes import wintypes
        
        # 1. Create DC
        hdc = self.user32.GetDC(0)
        mem_dc = self.gdi32.CreateCompatibleDC(hdc)
        
        # 2. Create Bitmap (24bit RGB)
        class BITMAPINFOHEADER(ctypes.Structure):
            _fields_ = [('biSize', wintypes.DWORD), ('biWidth', wintypes.LONG), ('biHeight', wintypes.LONG),
                        ('biPlanes', wintypes.WORD), ('biBitCount', wintypes.WORD), ('biCompression', wintypes.DWORD),
                        ('biSizeImage', wintypes.DWORD), ('biXPelsPerMeter', wintypes.LONG), ('biYPelsPerMeter', wintypes.LONG),
                        ('biClrUsed', wintypes.DWORD), ('biClrImportant', wintypes.DWORD)]
        class BITMAPINFO(ctypes.Structure):
            _fields_ = [('bmiHeader', BITMAPINFOHEADER), ('bmiColors', wintypes.DWORD * 3)]

        bmi = BITMAPINFO()
        bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bmi.bmiHeader.biWidth = width
        bmi.bmiHeader.biHeight = -height
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 24
        bmi.bmiHeader.biCompression = 0 
        
        pBits = ctypes.c_void_p()
        hBitmap = self.gdi32.CreateDIBSection(mem_dc, ctypes.byref(bmi), 0, ctypes.byref(pBits), 0, 0)
        self.gdi32.SelectObject(mem_dc, hBitmap)
        
        # 3. Draw Background
        # Dark Modern Grey Background
        bg_brush = self.gdi32.CreateSolidBrush(0x1E1E1E) # RGB(30,30,30) - BGR hex: 1E1E1E
        rect = wintypes.RECT(0, 0, width, height)
        self.user32.FillRect(mem_dc, ctypes.byref(rect), bg_brush)
        self.gdi32.DeleteObject(bg_brush)
        
        # Accent Border at Bottom
        border_brush = self.gdi32.CreateSolidBrush(0xD47800) # RGB(0, 120, 212) -> BGR: D47800 (Blue)
        border_rect = wintypes.RECT(0, height-4, width, height)
        self.user32.FillRect(mem_dc, ctypes.byref(border_rect), border_brush)
        self.gdi32.DeleteObject(border_brush)

        # 4. Setup Fonts
        self.gdi32.SetBkMode(mem_dc, 1) # TRANSPARENT
        
        # Main Text Font (Large, White)
        lfHeight = -72 
        hFontMain = self.gdi32.CreateFontW(lfHeight, 0, 0, 0, self.FW_BOLD, 0, 0, 0, 128, 0, 0, 0, 2, "Yu Gothic UI Semibold") 
        
        # Candidate Font (Medium, Grey/Accent)
        lfHeightSub = -36 
        hFontSub = self.gdi32.CreateFontW(lfHeightSub, 0, 0, 0, 400, 0, 0, 0, 128, 0, 0, 0, 2, "Yu Gothic UI")

        # Morse Font (Monospace, Accent Color)
        lfHeightMorse = -48
        hFontMorse = self.gdi32.CreateFontW(lfHeightMorse, 0, 0, 0, self.FW_BOLD, 0, 0, 0, 0, 0, 0, 0, 2, "Courier New")

        # DT_FLAGS: CENTER(1) | VCENTER(4) | SINGLELINE(32=0x20) => 0x25
        # Actually simpler: DT_CENTER(1) | DT_TOP(0)
        DT_CENTER = 0x1
        
        # 5. Draw Candidates (Upper Area)
        if candidates_context:
            self.gdi32.SelectObject(mem_dc, hFontSub)
            self.gdi32.SetTextColor(mem_dc, 0xAAAAAA) # Light Grey
            c_rect = wintypes.RECT(20, 20, width-20, 80)
            self.user32.DrawTextW(mem_dc, candidates_context, -1, ctypes.byref(c_rect), DT_CENTER)
            
            # Draw Separator Line
            pen = self.gdi32.CreatePen(0, 1, 0x444444)
            self.gdi32.SelectObject(mem_dc, pen)
            self.gdi32.MoveToEx(mem_dc, 20, 85, None)
            self.gdi32.LineTo(mem_dc, width-20, 85)
            self.gdi32.DeleteObject(pen)

        # 6. Draw Main Text (Middle Area)
        self.gdi32.SelectObject(mem_dc, hFontMain)
        self.gdi32.SetTextColor(mem_dc, 0xFFFFFF) # White
        
        # DT_CENTER | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX
        # DT_END_ELLIPSIS = 0x8000, DT_NOPREFIX = 0x800
        # Check standard values: 
        # DT_CENTER(1) | DT_VCENTER(4) | DT_SINGLELINE(32)? 
        # If single line, ellipsis works. If multiline, needed wordbreak.
        # Let's assume single line scrolling is preferred, so ellipsis at end.
        draw_flags = 0x1 | 0x4 | 0x20 | 0x8000 | 0x800 
        
        m_rect = wintypes.RECT(20, 100, width-20, 200)
        self.user32.DrawTextW(mem_dc, text, -1, ctypes.byref(m_rect), draw_flags)
        
        # 7. Draw Morse Code (Bottom Area)
        if morse_text:
             morse_display = morse_text.replace(".", "・").replace("-", "ー")
             
             self.gdi32.SelectObject(mem_dc, hFontMorse)
             self.gdi32.SetTextColor(mem_dc, 0xFFFFFF) # White
             
             mo_rect = wintypes.RECT(20, 210, width-20, 280)
             self.user32.DrawTextW(mem_dc, morse_display, -1, ctypes.byref(mo_rect), DT_CENTER)
        
        # 8. Save BMP (Bottom-up format for OpenVR compatibility)
        row_stride = (width * 3 + 3) & ~3
        image_size = row_stride * height
        bfSize = 14 + 40 + image_size
        
        file_header = b'BM' + struct.pack('<I', bfSize) + b'\x00\x00\x00\x00' + struct.pack('<I', 54)
        # Use positive height for bottom-up BMP (OpenVR compatible)
        info_header = struct.pack('<I', 40) + struct.pack('<i', width) + struct.pack('<i', height) + \
                      struct.pack('<H', 1) + struct.pack('<H', 24) + struct.pack('<I', 0) + \
                      struct.pack('<I', image_size) + struct.pack('<i', 0) + struct.pack('<i', 0) + \
                      struct.pack('<I', 0) + struct.pack('<I', 0)
                      
        buffer = (ctypes.c_char * image_size).from_address(pBits.value)
        raw_data = bytearray(buffer)
        
        # Flip rows for bottom-up BMP (since GDI rendered top-down)
        flipped_data = bytearray(image_size)
        for row in range(height):
            src_offset = row * row_stride
            dst_offset = (height - 1 - row) * row_stride
            flipped_data[dst_offset:dst_offset + row_stride] = raw_data[src_offset:src_offset + row_stride]
        
        with open(self.temp_file, "wb") as f:
            f.write(file_header)
            f.write(info_header)
            f.write(flipped_data)
            
        # Cleanup
        self.gdi32.DeleteObject(hBitmap)
        self.gdi32.DeleteObject(hFontMain)
        self.gdi32.DeleteObject(hFontSub)
        self.gdi32.DeleteObject(hFontMorse)
        self.gdi32.DeleteDC(mem_dc)
        self.user32.ReleaseDC(0, hdc)


    def update_image(self, text, candidates_str=None, morse_text=None):
        if not self.enabled or not self.use_gdi:
            return
        
        # Check if identical to skip
        current_state = (text, candidates_str, morse_text)
        if hasattr(self, 'last_state') and self.last_state == current_state:
            return
            
        try:
            self.log_debug(f"Overlay: 画像生成開始 text='{text}'")
            self.create_text_bitmap(text, candidates_str, morse_text)
            
            # Verify file exists
            if not os.path.exists(self.temp_file):
                self.log(f"Overlay: ファイルが存在しません: {self.temp_file}")
                return
            
            file_size = os.path.getsize(self.temp_file)
            self.log_debug(f"Overlay: BMPファイル生成完了, size={file_size} bytes")
            
            result = self.vr_overlay.setOverlayFromFile(self.handle, self.temp_file)
            # In pyopenvr, None typically means success (no error returned)
            if result is None or result == 0:
                self.log_debug(f"Overlay: setOverlayFromFile成功")
            else:
                self.log(f"Overlay: setOverlayFromFile failed with error code {result}")
            self.last_state = current_state
        except Exception as e:
            import traceback
            self.log(f"Overlay Update Error: {e}")
            self.log(traceback.format_exc())

    def set_active(self, active):
        """Sets the logical active state (e.g. stick is down)."""
        if not self.enabled: return
        
        if self.is_active != active:
            self.log_debug(f"Overlay: set_active({active})")
        
        self.is_active = active
        if active:
            self.last_active_time = time.time()
            # Immediate show if active
            if not self.currently_visible:
                 self.show()

    def process_visibility(self):
        """Called every loop to check timeouts."""
        if not self.enabled or not self.handle: return
        
        if self.is_active:
             self.last_active_time = time.time()
             if not self.currently_visible:
                 self.show()
        else:
             # If inactive, check timeout
             if self.currently_visible:
                 elapsed = time.time() - self.last_active_time
                 if elapsed > self.visibility_timeout:
                     self.hide()

    def show(self):
        try:
            self.log_debug(f"Overlay: show() handle={self.handle}")
            self.vr_overlay.showOverlay(self.handle)
            self.currently_visible = True
        except Exception as e:
            self.log(f"Overlay: showOverlay失敗: {e}")

    def hide(self):
        try:
            self.log_debug("Overlay: hide()")
            self.vr_overlay.hideOverlay(self.handle)
            self.currently_visible = False
        except Exception as e:
            self.log(f"Overlay: hideOverlay失敗: {e}")

    def shutdown(self):
        if self.enabled and self.handle:
            try:
                self.log(f"Overlay: shutdown - destroyOverlay handle={self.handle}")
                self.vr_overlay.destroyOverlay(self.handle)
                self.handle = None
                self.enabled = False
                self.log("Overlay: shutdown完了")
            except Exception as e:
                self.log(f"Overlay: shutdown失敗: {e}")


# --- Japanese Conversion Helper ---
# Simple Romaji to Hiragana Table
ROMAJI_KANA_MAP = {
    'a':'あ', 'i':'い', 'u':'う', 'e':'え', 'o':'お',
    'ka':'か', 'ki':'き', 'ku':'く', 'ke':'け', 'ko':'こ',
    'sa':'さ', 'si':'し', 'shi':'し', 'su':'す', 'se':'せ', 'so':'そ',
    'ta':'た', 'ti':'ち', 'chi':'ち', 'tu':'つ', 'tsu':'つ', 'te':'て', 'to':'と',
    'na':'な', 'ni':'に', 'nu':'ぬ', 'ne':'ね', 'no':'の',
    'ha':'は', 'hi':'ひ', 'hu':'ふ', 'fu':'ふ', 'he':'へ', 'ho':'ほ',
    'ma':'ま', 'mi':'み', 'mu':'む', 'me':'め', 'mo':'も',
    'ya':'や', 'yi':'い', 'yu':'ゆ', 'ye':'い', 'yo':'よ',
    'ra':'ら', 'ri':'り', 'ru':'る', 're':'れ', 'ro':'ろ',
    'wa':'わ', 'wo':'を', 'nn':'ん', 'xn':'ん',
    'ga':'が', 'gi':'ぎ', 'gu':'ぐ', 'ge':'げ', 'go':'ご',
    'za':'ざ', 'ji':'じ', 'zi':'じ', 'zu':'ず', 'ze':'ぜ', 'zo':'ぞ',
    'da':'だ', 'di':'ぢ', 'du':'づ', 'de':'で', 'do':'ど',
    'ba':'ば', 'bi':'び', 'bu':'ぶ', 'be':'べ', 'bo':'ぼ',
    'pa':'ぱ', 'pi':'ぴ', 'pu':'ぷ', 'pe':'ぺ', 'po':'ぽ',
    
    # KYA...
    'kya':'きゃ', 'kyi':'きぃ', 'kyu':'きゅ', 'kye':'きぇ', 'kyo':'きょ',
    'gya':'ぎゃ', 'gyi':'ぎぃ', 'gyu':'ぎゅ', 'gye':'ぎぇ', 'gyo':'ぎょ',
    'sha':'しゃ', 'shu':'しゅ', 'she':'しぇ', 'sho':'しょ',
    'sya':'しゃ', 'syu':'しゅ', 'sye':'しぇ', 'syo':'しょ', # Added
    'ja':'じゃ', 'ju':'じゅ', 'je':'じぇ', 'jo':'じょ',
    'jya':'じゃ', 'jyu':'じゅ', 'jye':'じぇ', 'jyo':'じょ', # Added
    'cha':'ちゃ', 'chu':'ちゅ', 'che':'ちぇ', 'cho':'ちょ',
    'tya':'ちゃ', 'tyi':'ちぃ', 'tyu':'ちゅ', 'tye':'ちぇ', 'tyo':'ちょ', # Added
    'cya':'ちゃ', 'cyi':'ちぃ', 'cyu':'ちゅ', 'cye':'ちぇ', 'cyo':'ちょ', # Added
    'nya':'にゃ', 'nyi':'にぃ', 'nyu':'にゅ', 'nye':'にぇ', 'nyo':'にょ',
    'hya':'ひゃ', 'hyi':'ひぃ', 'hyu':'ひゅ', 'hye':'ひぇ', 'hyo':'ひょ',
    'bya':'びゃ', 'byi':'びぃ', 'byu':'びゅ', 'bye':'びぇ', 'byo':'びょ',
    'pya':'ぴゃ', 'pyi':'ぴぃ', 'pyu':'ぴゅ', 'pye':'ぴぇ', 'pyo':'ぴょ',
    'mya':'みゃ', 'myi':'みぃ', 'myu':'みゅ', 'mye':'みぇ', 'myo':'みょ',
    'rya':'りゃ', 'ryi':'りぃ', 'ryu':'りゅ', 'rye':'りぇ', 'ryo':'りょ',
    
    # F...
    'fa':'ふぁ', 'fi':'ふぃ', 'fe':'ふぇ', 'fo':'ふぉ',
    'fya':'ふゃ', 'fyu':'ふゅ', 'fyo':'ふょ',
    
    # V...
    'va':'ヴぁ', 'vi':'ヴぃ', 'vu':'ヴ', 've':'ヴぇ', 'vo':'ヴぉ',
    
    # W...
    'wi':'うぃ', 'we':'うぇ',
    
    # Small / X / L
    'la':'ぁ', 'li':'ぃ', 'lu':'ぅ', 'le':'ぇ', 'lo':'ぉ', 'ltu':'っ',
    'xa':'ぁ', 'xi':'ぃ', 'xu':'ぅ', 'xe':'ぇ', 'xo':'ぉ', 'xtu':'っ',
    
    '-':'ー'
}

def to_hiragana(romaji):
    """
    Robust Romaji -> Hiragana converter.
    """
    text = romaji.lower()
    res = ""
    i = 0
    length = len(text)
    
    while i < length:
        # 1. Try 3-char match (kya, shu, etc.)
        if i + 3 <= length:
            chunk3 = text[i:i+3]
            if chunk3 in ROMAJI_KANA_MAP:
                res += ROMAJI_KANA_MAP[chunk3]
                i += 3
                continue
        # Special case: End of string can match shorter keys in 3-char block? No, strict slice.
        # But ROMAJI_KANA_MAP keys are specific.
        
        # 2. Try 2-char match (ka, n, etc.)
        if i + 2 <= length:
            chunk2 = text[i:i+2]
            if chunk2 in ROMAJI_KANA_MAP:
                res += ROMAJI_KANA_MAP[chunk2]
                i += 2
                continue
            
            # small tsu (double consonant)
            # check cc, kk, ss, tt etc. (not nn, not aa)
            c1 = chunk2[0]
            c2 = chunk2[1]
            if c1 == c2 and c1 not in 'aeioun':
                res += 'っ'
                i += 1 # Only consume the first consonant 't', next 't' starts 'ta' etc.
                continue
                
            # 'n' special cases
            # 'n' + consonant (not n, not y) -> 'ん'
            # 'n' + vowel/y -> processed by na/ni/nu/ne/no/nya map keys
            # 'n' + space or symbol?
            if c1 == 'n' and c2 not in 'aeiouy':
                res += 'ん'
                i += 1
                continue

        # 3. Try 1-char match (a, i, u, e, o, n at end)
        chunk1 = text[i]
        if chunk1 in ROMAJI_KANA_MAP:
            res += ROMAJI_KANA_MAP[chunk1]
            i += 1
            continue
            
        # Single 'n' at very end
        if chunk1 == 'n':
            res += 'ん'
            i += 1
            continue
            
        # Passthrough
        res += chunk1
        i += 1
        
    return res

def request_conversion(romaji_text, callback, log_func=print):
    """
    Hybrid Conversion:
    1. Local Romaji -> Hiragana (Immediate, Reliable)
    2. Google IME API via HTTP (Candidate Enrichment)
    """
    import requests
    import json
    
    def _request():
        if not romaji_text: return
        
        # 1. Local Conversion (Always works)
        hiragana = to_hiragana(romaji_text)
        local_candidates = [hiragana, romaji_text] # Hiragana first, then raw
        
        # 2. Cloud Conversion (Attempt)
        cloud_candidates = []
        try:
            # Use HTTP to avoid SSL issues (Error 1) and standard User-Agent
            url = "http://www.google.com/transliterate"
            # params_alt already defined in previous Logic but let's restate for clarity in block
            params_alt = {"langpair": "ja-Hira|ja", "text": hiragana}
            
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" 
            }
            
            response = requests.get(url, params=params_alt, headers=headers, timeout=2.0)
            
            if response.status_code == 200:
                data = response.json()
                # API Format: [ [segment_raw, [cand1, cand2...]], [segment_raw2, ...], ... ]
                
                if isinstance(data, list) and len(data) > 0:
                    # Construct ONE combined result by joining the TOP candidate of each segment.
                    # This prevents "loss" of later segments if the API splits the sentence.
                    combined_top = ""
                    for segment in data:
                        if isinstance(segment, list) and len(segment) >= 2:
                             cands = segment[1]
                             if cands:
                                 combined_top += cands[0]
                             else:
                                 combined_top += segment[0] # Fallback to raw if no cands
                        else:
                             # Fallback for unexpected structure
                             if isinstance(segment, list) and len(segment) > 0:
                                 combined_top += str(segment[0])
                    
                    if combined_top:
                        cloud_candidates = [combined_top]
                        
        except Exception:
            pass # Silent fail on network, fall back to local
            
        # Combine: Cloud candidates first (usually Kanji), then distinct Local ones
        final_list = []
        seen = set()
        
        # Prefer Cloud Results (Kanji)
        for c in cloud_candidates:
            if c not in seen:
                final_list.append(c)
                seen.add(c)
        
        # Fallback Local Results (Hiragana)
        for c in local_candidates:
            if c not in seen:
                final_list.append(c)
                seen.add(c)
        
        callback(final_list)

    threading.Thread(target=_request, daemon=True).start()

def confirm_conversion():
    """Confirms the currently selected candidate and exits conversion mode."""
    if state.conversion_active and state.conversion_candidates:
        selected = state.conversion_candidates[state.conversion_index]
        state.fixed_text += selected
        # state.text_buffer is already "consumed" visually, so we just clear state
        state.conversion_active = False
        state.conversion_candidates = []
        state.conversion_index = 0
        state.text_buffer = "" # Clear the source buffer
        return True
    return False

def recalculate_derived_values():
    wpm_dot = state.settings.get("wpmDot", 20)
    wpm_dash = state.settings.get("wpmDash", 20)
    
    state.unit_dot = 1200 / wpm_dot
    state.unit_dash = 1200 / wpm_dash
    
    if "charTimeout" in state.settings:
        state.char_timeout = float(state.settings["charTimeout"])
    
    if "triggerThreshold" in state.settings:
        state.trigger_threshold = float(state.settings["triggerThreshold"])

    # Generate Audio
    freq = state.settings["freq"]
    dash_len = state.unit_dash * state.settings["dashRatio"]
    state.sound_dot = generate_wav_bytes(freq, state.unit_dot)
    state.sound_dash = generate_wav_bytes(freq, dash_len)

def load_settings():
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r') as f:
                loaded = json.load(f)
                state.settings.update(loaded)
    except Exception as e:
        print(f"Error loading settings: {e}")
    recalculate_derived_values()

def save_settings():
    try:
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(state.settings, f, indent=4)
        print("Settings saved.")
    except Exception as e:
        print(f"Error saving settings: {e}")

# --- Core Logic ---
MORSE_TO_CHAR = {
    '.-': 'A', '-...': 'B', '-.-.': 'C', '-..': 'D', '.': 'E', '..-.': 'F',
    '--.': 'G', '....': 'H', '..': 'I', '.---': 'J', '-.-': 'K', '.-..': 'L',
    '--': 'M', '-.': 'N', '---': 'O', '.--.': 'P', '--.-': 'Q', '.-.': 'R',
    '...': 'S', '-': 'T', '..-': 'U', '...-': 'V', '.--': 'W', '-..-': 'X',
    '-.--': 'Y', '--..': 'Z',
    '.----': '1', '..---': '2', '...--': '3', '....-': '4', '.....': '5',
    '-....': '6', '--...': '7', '---..': '8', '----.': '9', '-----': '0'
}

def play_sound_realtime(symbol):
    sound_data = state.sound_dot if symbol == '.' else state.sound_dash
    if sound_data:
        def _play():
            try:
                winsound.PlaySound(sound_data, winsound.SND_MEMORY)
            except: pass
        
        winsound.PlaySound(None, winsound.SND_PURGE) 
        threading.Thread(target=_play, daemon=True).start()

def process_input(symbol):
    now = time.time()
    
    # [Requirement] If input occurs during conversion, confirm the candidate first
    if state.conversion_active:
        confirm_conversion()
        # After confirm, we proceed to process the new dot/dash
        # Update output immediately to show confirmed state before adding new char
        state.osc.send("/chatbox/input", [state.fixed_text, True])
    
    print(f"Input: {symbol}")
    state.current_symbol_sequence += symbol
    state.last_char_confirmed_time = now
    
    play_sound_realtime(symbol)
    
    unit = state.unit_dot if symbol == '.' else state.unit_dash
    ratio = state.settings["dashRatio"]
    duration_ms = unit if symbol == '.' else (unit * ratio)
    
    gap_ms = state.settings["playbackGap"]
    wait_ms = duration_ms + gap_ms
    state.next_repeat_allowed_at = now + (wait_ms / 1000.0)

# --- VR Loop ---
def vr_loop(app_instance):
    """
    Main VR Input Loop.
    app_instance: Reference to MorseApp for logging.
    """
    # Helper wrappers for app_instance.log with colors
    log_info = lambda msg: app_instance.log(msg, "INFO")
    log_warn = lambda msg: app_instance.log(msg, "WARN")
    log_err  = lambda msg: app_instance.log(msg, "ERROR")
    log_in   = lambda msg: app_instance.log(msg, "INPUT")
    
    log_info("OpenVR 初期化中...")
    try:
        try:
            vr_system = openvr.init(openvr.VRApplication_Background)
        except openvr.OpenVRError as e:
            log_err(f"OpenVR Init Failed: {e}")
            app_instance.after(0, lambda: messagebox.showwarning("SteamVR未検出", "SteamVRが起動していないか、検出できませんでした。\nSteamVRを起動してから再度実行してください。"))
            return

        state.vr_initialized = True
        try:
            model = vr_system.getStringTrackedDeviceProperty(openvr.k_unTrackedDeviceIndex_Hmd, openvr.Prop_ModelNumber_String)
        except:
            model = "Unknown"
        log_info(f"OpenVR Initialized. Hardware: {model}")

        log_info("OSC Client 初期化中...")
        # Note: OSCManager is already global in state.osc
        
        # Debug log wrapper
        def log_debug(msg):
            if state.settings.get("debugMode", False):
                log_info(msg)
        
        # Init Overlay
        log_debug(f"Overlay設定: overlayEnabled={state.settings.get('overlayEnabled', False)}")
        log_debug(f"全設定: {state.settings}")
        overlay = OverlayManager(key="morse_input_overlay", name="Morse Input", app_state=state, log_func=log_info)
        state.overlay = overlay # Store in state for global access
        if overlay.enabled:
            log_info("Overlay初期化完了")
        else:
            log_info("Overlay無効 (設定またはエラー)")
        
        log_info("--- ループ開始 (Waiting for Input) ---")
        
        left_index = -1
        right_index = -1
        prev_dot = False
        prev_dash = False
        
        # State tracking to avoid spamming logs
        was_input_enabled = None # Tri-state: None, True, False
        was_stick_up = False # Stick Up Edge Detection

        while state.running:
            # Poll all events to keep connection alive
            event = openvr.VREvent_t()
            while vr_system.pollNextEvent(event):
                pass
            
            # Device Discovery (Periodic or if missing)
            if left_index == -1 or right_index == -1:
                for i in range(openvr.k_unMaxTrackedDeviceCount):
                    try:
                        cls = vr_system.getTrackedDeviceClass(i)
                        if cls == openvr.TrackedDeviceClass_Controller:
                            role = vr_system.getControllerRoleForTrackedDeviceIndex(i)
                            if role == openvr.TrackedControllerRole_LeftHand:
                                left_index = i
                            elif role == openvr.TrackedControllerRole_RightHand:
                                right_index = i
                    except: pass

            # Read Controller States
            state_l = None
            state_r = None
            
            if left_index != -1:
                state_l = vr_system.getControllerState(left_index)[1]
            if right_index != -1:
                state_r = vr_system.getControllerState(right_index)[1]

            # Enable Check logic
            input_enabled = True
            current_stick_y = 0.0
            
            # Stick Logic (Right Hand)
            if state_r:
                current_stick_y = state_r.rAxis[0].y
                
                # 1. Safety Switch (Stick Down to Input)
                if state.settings.get("requireStickDown", False):
                    # Y < -0.7 means Down
                    if current_stick_y > -0.7:
                        input_enabled = False
                
                # 2. Stick Up Clear (Stick Up to Clear)
                # Y > 0.7 means Up
                if current_stick_y > 0.7:
                    if not was_stick_up:
                        # Edge: Up Pressed
                        has_text = len(state.text_buffer) > 0 or len(state.fixed_text) > 0 or state.conversion_active
                        if has_text:
                            state.text_buffer = ""
                            state.fixed_text = ""
                            state.conversion_active = False
                            state.conversion_candidates = []
                            state.conversion_index = 0
                            state.conversion_request_pending = False
                            
                            log_info("EDIT: CLEAR_ALL")
                            state.osc.send("/chatbox/input", ["", True])
                            state.overlay.update_image(*generate_overlay_text())
                    was_stick_up = True
                else:
                    was_stick_up = False

            elif state.settings.get("requireStickDown", False):
                input_enabled = False # No controller means no stick down
            
            # Log State Change for clarity
            if input_enabled != was_input_enabled:
                # Logging suppressed per request
                was_input_enabled = input_enabled
                
                # Overlay Visibility Control
                try:
                    state.overlay.set_active(input_enabled)
                    # Force update content on visibility toggle to ensure freshness
                    if input_enabled:
                         state.overlay.update_image(*generate_overlay_text())
                except: pass

            dot_down = False
            dash_down = False
            
            if input_enabled:
                # --- Inputs ---
                if state_l:
                    if state_l.rAxis[1].x > state.trigger_threshold:
                        dot_down = True
                
                if state_r:
                    if state_r.rAxis[1].x > state.trigger_threshold:
                        dash_down = True

            # --- Aux Controls Logic ---
            if state_r:
                 # 1. Backspace on Right Grip
                 grip_pressed = (state_r.ulButtonPressed & (1 << openvr.k_EButton_Grip)) != 0
                 was_grip = state.was_grip_down
                 
                 if grip_pressed and not was_grip:
                     if state.conversion_active:
                         state.conversion_active = False 
                         log_info("IME: ABORT_CONVERSION")
                     else:
                         if len(state.text_buffer) > 0:
                             state.text_buffer = state.text_buffer[:-1]
                             log_info("EDIT: BACKSPACE_BUFFER")
                         elif len(state.fixed_text) > 0:
                             state.fixed_text = state.fixed_text[:-1]
                             log_info("EDIT: BACKSPACE_FIXED")
                     
                     content = state.fixed_text
                     if state.conversion_active:
                         if state.conversion_candidates:
                             content += state.conversion_candidates[state.conversion_index]
                     else:
                         content += state.text_buffer
                     state.osc.send("/chatbox/input", [content, True])
                     state.overlay.update_image(*generate_overlay_text())
                 
                 state.was_grip_down = grip_pressed
            
            if state_l:
                # 2. Japanese Conversion on Left Grip
                left_grip_pressed = (state_l.ulButtonPressed & (1 << openvr.k_EButton_Grip)) != 0
                was_left_grip = state.was_left_grip_down
                
                if left_grip_pressed and not was_left_grip:
                    # Logic:
                    # If converting: Cycle
                    # If NOT converting and NOT pending: Request
                    
                    if state.conversion_active:
                        # Cycle
                        if state.conversion_candidates:
                            state.conversion_index = (state.conversion_index + 1) % len(state.conversion_candidates)
                            selected = state.conversion_candidates[state.conversion_index]
                            
                            log_info(f"IME: CYCLE_CANDIDATE [{state.conversion_index}] -> {selected}")
                            full_text = state.fixed_text + selected
                            state.osc.send("/chatbox/input", [full_text, True])
                            state.overlay.update_image(*generate_overlay_text())
                    
                    elif not state.conversion_request_pending:
                        # Request
                        if len(state.text_buffer) > 0:
                            romaji = state.text_buffer
                            state.conversion_request_pending = True # Lock
                            
                            def on_candidates(result_list):
                                state.conversion_request_pending = False # Unlock
                                if result_list:
                                    state.conversion_candidates = result_list
                                    state.conversion_index = 0
                                    state.conversion_active = True
                                    log_info(f"IME: CANDIDATES_FETCHED [{len(result_list)}]")
                                    
                                    selected = result_list[0]
                                    full_text = state.fixed_text + selected 
                                    state.osc.send("/chatbox/input", [full_text, True])
                                    state.overlay.update_image(*generate_overlay_text())
                                else:
                                    # Request finished but failed
                                    pass

                            log_info(f"IME: REQUEST_CONVERSION -> '{romaji}'")
                            request_conversion(romaji, on_candidates, log_func=log_warn)
                
                state.was_left_grip_down = left_grip_pressed

            
            now = time.time()
            input_symbol = None
            
            # --- Morse Logic ---
            # 1. Edge Detection
            if dash_down and not prev_dash:
                input_symbol = '-'
            elif dot_down and not prev_dot:
                if not dash_down:
                    input_symbol = '.'
            # 2. Repeat
            elif state.settings["keyRepeat"]:
                if now >= state.next_repeat_allowed_at:
                    if dash_down:
                        input_symbol = '-'
                        log_in("SIGNAL: REPEAT_DASH")
                    elif dot_down:
                        input_symbol = '.'
                        log_in("SIGNAL: REPEAT_DOT")
            
            if input_symbol:
                if state.conversion_active:
                    # Confirm pending candidate
                    if state.conversion_candidates:
                        selected = state.conversion_candidates[state.conversion_index]
                        state.fixed_text += selected
                        log_warn(f"IME: COMMIT -> {selected}")
                    
                    state.conversion_active = False
                    state.conversion_candidates = []
                    state.conversion_index = 0
                    state.text_buffer = "" 
                    
                    state.osc.send("/chatbox/input", [state.fixed_text, True])
                    state.overlay.update_image(*generate_overlay_text())

                log_in(f"SIGNAL: DETECTED '{input_symbol}'")
                state.current_symbol_sequence += input_symbol
                
                # Update overlay immediately to show the new dot/dash
                try:
                    state.overlay.update_image(*generate_overlay_text())
                except: pass
                state.last_char_confirmed_time = now
                
                # Audio
                sound_data = state.sound_dot if input_symbol == '.' else state.sound_dash
                if sound_data:
                    def _play():
                        try: winsound.PlaySound(sound_data, winsound.SND_MEMORY)
                        except: pass
                    winsound.PlaySound(None, winsound.SND_PURGE) 
                    threading.Thread(target=_play, daemon=True).start()
                
                # Calculate Timings
                unit = state.unit_dot if input_symbol == '.' else state.unit_dash
                ratio = state.settings["dashRatio"]
                duration_ms = unit if input_symbol == '.' else (unit * ratio)
                gap_ms = state.settings["playbackGap"]
                state.next_repeat_allowed_at = now + ((duration_ms + gap_ms) / 1000.0)
            
            prev_dot = dot_down
            prev_dash = dash_down
            
            # Confirm Char Logic
            if state.current_symbol_sequence:
                if (now - state.last_char_confirmed_time) > state.char_timeout:
                    if not dot_down and not dash_down:
                        seq = state.current_symbol_sequence
                        
                        # Check Custom Map First
                        custom_map = state.settings.get("customMorseMap", {})
                        char = custom_map.get(seq, None)
                        
                        # Fallback to Standard Map
                        if char is None:
                            char = MORSE_TO_CHAR.get(seq, '?')
                        
                        if char != '?':
                            state.text_buffer += char
                            log_info(f"DECODE: '{seq}' -> '{char}' | BUF: {state.text_buffer}")
                            
                            display_text = state.fixed_text
                            if state.conversion_active:
                                 if state.conversion_candidates:
                                     display_text += state.conversion_candidates[state.conversion_index]
                            else:
                                 display_text += state.text_buffer
                            
                            state.osc.send("/chatbox/input", [display_text, True])
                            state.overlay.update_image(*generate_overlay_text())
                        else:
                            log_warn(f"DECODE: UNKNOWN_SEQ '{seq}'")
                        
                        state.current_symbol_sequence = ""
                        state.last_char_confirmed_time = now

            # OSC Queue Processing
            state.osc.process_queue()
            
            # Overlay Visibility Loop
            if hasattr(state, 'overlay') and state.overlay:
                 state.overlay.process_visibility()

            # Sleep
            time.sleep(0.0105)

    except Exception as e:
        import traceback
        log_err(f"SYS_FAILURE: {traceback.format_exc()}")
    
    finally:
        try:
            state.overlay.shutdown()
        except: pass
        try:
            openvr.shutdown()
        except: pass
        log_info("SYSTEM: SHUTDOWN_SEQUENCE_COMPLETE")
        state.running = False
        app_instance.after(0, app_instance.reset_ui_state)

# --- GUI ---
class Colors:
    BG_MAIN = "#0F0F0F"      # Darker
    BG_SIDE = "#181818"      
    BG_CARD = "#1A1A1A"      
    BG_ENTRY = "#000000"     # Pure Black for Terminal
    ACCENT = "#0078D4"       
    ACCENT_HOVER = "#006CC0"
    TEXT_MAIN = "#E0E0E0"
    TEXT_SUB = "#808080"
    BORDER = "#333333"

# ... (Previous GUI classes remain similar but we update _init_console_view inline) ...
# Actually need to target _init_console_view specifically to remove header


class MorseApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("VR Morse Input")
        self.geometry("960x720")
        self.configure(bg=Colors.BG_MAIN)
        self.minsize(800, 600)

        # --- Style Engine ---
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure common styles
        style.configure(".", 
            background=Colors.BG_MAIN, 
            foreground=Colors.TEXT_MAIN, 
            font=("Yu Gothic UI", 10),
            borderwidth=0
        )
        
        # TFrame defaults
        style.configure("TFrame", background=Colors.BG_MAIN)
        style.configure("Card.TFrame", background=Colors.BG_CARD)
        style.configure("Sidebar.TFrame", background=Colors.BG_SIDE)
        
        # Labels
        style.configure("TLabel", background=Colors.BG_MAIN, foreground=Colors.TEXT_MAIN)
        style.configure("Card.TLabel", background=Colors.BG_CARD, foreground=Colors.TEXT_MAIN)
        style.configure("Header.TLabel", 
            background=Colors.BG_MAIN, 
            foreground=Colors.TEXT_MAIN, 
            font=("Yu Gothic UI", 18)
        )
        style.configure("SubHeader.TLabel", 
            background=Colors.BG_CARD, 
            foreground=Colors.ACCENT, 
            font=("Yu Gothic UI", 11)
        )
        style.configure("SidebarTitle.TLabel", 
            background=Colors.BG_SIDE, 
            foreground=Colors.TEXT_MAIN, 
            font=("Yu Gothic UI", 14)
        )
        style.configure("Description.TLabel",
            background=Colors.BG_CARD,
            foreground=Colors.TEXT_SUB,
            font=("Yu Gothic UI", 9)
        )

        # Buttons (Navigation)
        style.configure("Nav.TButton", 
            background=Colors.BG_SIDE, 
            foreground=Colors.TEXT_SUB, 
            borderwidth=0, 
            focuscolor=Colors.BG_SIDE, 
            font=("Yu Gothic UI", 11),
            anchor="w",
            padding=(20, 12)
        )
        style.map("Nav.TButton", 
            background=[("active", "#2D2D2D"), ("pressed", Colors.ACCENT)],
            foreground=[("active", Colors.TEXT_MAIN), ("pressed", Colors.TEXT_MAIN)]
        )
        
        # Buttons (Action)
        style.configure("Action.TButton",
            background=Colors.ACCENT,
            foreground="#FFFFFF",
            borderwidth=0,
            font=("Yu Gothic UI", 10),
            padding=(20, 10)
        )
        style.map("Action.TButton",
            background=[("active", Colors.ACCENT_HOVER), ("pressed", "#005A9E")]
        )
        
        # Inputs
        style.configure("Modern.TEntry", 
            fieldbackground=Colors.BG_ENTRY, 
            foreground=Colors.TEXT_MAIN, 
            insertcolor=Colors.TEXT_MAIN,
            borderwidth=0, 
            padding=5
        )
        

        
        # Treeview (Custom List)
        style.configure("Treeview",
            background=Colors.BG_ENTRY,
            foreground=Colors.TEXT_MAIN,
            fieldbackground=Colors.BG_ENTRY,
            borderwidth=0,
            font=("Yu Gothic UI", 10),
            rowheight=25
        )
        style.configure("Treeview.Heading",
            background=Colors.BG_SIDE,
            foreground=Colors.TEXT_MAIN,
            relief="flat",
            font=("Yu Gothic UI", 9, "bold"),
            padding=(10, 5)
        )
        style.map("Treeview.Heading",
            background=[("active", "#2D2D2D")]
        )
        style.map("Treeview",
            background=[("selected", Colors.ACCENT)],
            foreground=[("selected", "#FFFFFF")]
        )

        # Checkbuttons
        style.configure("Modern.TCheckbutton", 
            background=Colors.BG_CARD, 
            foreground=Colors.TEXT_MAIN, 
            font=("Yu Gothic UI", 10)
        )
        style.map("Modern.TCheckbutton",
            background=[("active", Colors.BG_CARD)],
            indicatorcolor=[("selected", Colors.ACCENT)]
        )

        # --- Layout Architecture ---
        # 1. Sidebar (Left)
        self.sidebar = ttk.Frame(self, style="Sidebar.TFrame", width=240)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar.pack_propagate(False) # Force width
        
        # 2. Main Content (Right)
        self.main_area = ttk.Frame(self, style="TFrame")
        self.main_area.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.frames = {}
        self.vars = {}

        self._init_sidebar()
        self._init_console_view()
        self._init_settings_view()
        self._init_about_view()
        
        # Initialize
        load_settings()
        self.populate_settings()
        
        # Don't capture stdout globally to avoid recursion if printing inside methods, 
        # but we use explicit log function now.
        # sys.stdout = self
        
        # Start at Console
        self.show_frame("Console")
        
        # Auto-start system
        self.after(200, self.toggle_vr)

    def _init_sidebar(self):
        # App Title
        title_frame = ttk.Frame(self.sidebar, style="Sidebar.TFrame")
        title_frame.pack(fill=tk.X, pady=(40, 40), padx=20)
        ttk.Label(title_frame, text="VR MORSE", style="SidebarTitle.TLabel").pack(anchor="w")
        ttk.Label(title_frame, text="INPUT SYSTEM", style="SidebarTitle.TLabel", font=("Yu Gothic UI", 8), foreground=Colors.TEXT_SUB).pack(anchor="w")

        # Navigation
        self.btn_nav_console = ttk.Button(self.sidebar, text="📊  ログ・コンソール", style="Nav.TButton", command=lambda: self.show_frame("Console"))
        self.btn_nav_console.pack(fill=tk.X, pady=1)

        self.btn_nav_settings = ttk.Button(self.sidebar, text="⚙  システム設定", style="Nav.TButton", command=lambda: self.show_frame("Settings"))
        self.btn_nav_settings.pack(fill=tk.X, pady=1)
        
        # Spacer
        ttk.Frame(self.sidebar, style="Sidebar.TFrame").pack(fill=tk.BOTH, expand=True)

        # About Tab (Bottom)
        self.btn_nav_about = ttk.Button(self.sidebar, text="ℹ  情報", style="Nav.TButton", command=lambda: self.show_frame("About"))
        self.btn_nav_about.pack(fill=tk.X, pady=1)
        
        # Status / Action
        status_frame = ttk.Frame(self.sidebar, style="Sidebar.TFrame")
        status_frame.pack(fill=tk.X, padx=20, pady=30)
        
        self.btn_start = ttk.Button(status_frame, text="▶  システム起動", style="Action.TButton", command=self.toggle_vr)
        self.btn_start.pack(fill=tk.X)

    def _init_console_view(self):
        container = ttk.Frame(self.main_area, style="TFrame", padding=40)
        
        # Console Card
        card = ttk.Frame(container, style="Card.TFrame", padding=1) # Border feeling
        card.pack(fill=tk.BOTH, expand=True)
        
        # Inner text
        self.console = scrolledtext.ScrolledText(
            card, 
            state='disabled', 
            bg=Colors.BG_ENTRY, 
            fg="#E0E0E0", 
            font=("Courier New", 10, "normal"), 
            borderwidth=0,
            highlightthickness=0,
            padx=15,
            pady=15
        )
        self.console.pack(fill=tk.BOTH, expand=True, padx=1, pady=1) # 1px margin for border effect
        
        # Configure Tags for Color
        self.console.tag_config("INFO", foreground="#CCCCCC")
        self.console.tag_config("WARN", foreground="#FFD700") # Gold
        self.console.tag_config("ERROR", foreground="#FF4444") # Red
        self.console.tag_config("INPUT", foreground="#00BFFF") # Deep Sky Blue
        
        self.frames["Console"] = container

    def _init_settings_view(self):
        container = ttk.Frame(self.main_area, style="TFrame", padding=40)
        container.pack_propagate(False) # Let it fill parent but not shrink
        container.pack(fill=tk.BOTH, expand=True) # Ensure it takes space
        
        # Header (Fixed at top)
        header_frame = ttk.Frame(container, style="TFrame")
        header_frame.pack(fill=tk.X, pady=(0, 20))
        ttk.Label(header_frame, text="設定", style="Header.TLabel").pack(side=tk.LEFT)
        ttk.Label(header_frame, text="変更は自動的に保存されます", style="Description.TLabel").pack(side=tk.RIGHT, anchor="s", pady=5)

        # Scrollable Area Setup
        canvas_frame = ttk.Frame(container, style="TFrame")
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(canvas_frame, bg=Colors.BG_MAIN, highlightthickness=0, bd=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        
        content = ttk.Frame(canvas, style="TFrame")
        
        content_window = canvas.create_window((0, 0), window=content, anchor="nw")
        
        def _configure_content(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(content_window, width=event.width)
            
        def _mouse_scroll(event):
            if content.winfo_height() > canvas.winfo_height():
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        content.bind("<Configure>", _configure_content)
        canvas.bind("<Configure>", _configure_content) # Sync width
        
        # Bind Mouse Wheel to all children
        def _bind_to_children(widget):
            widget.bind("<MouseWheel>", _mouse_scroll)
            for child in widget.winfo_children():
                _bind_to_children(child)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # --- Settings Content Calls ---
        self._create_card(content, "入力速度・タイミング", [
            ("短点速度 (WPM Dot)", "wpmDot", int, "短音の単位速度。値を上げると速くなります。"),
            ("長点速度 (WPM Dash)", "wpmDash", int, "長音の単位速度。通常は短点と同じか、少し遅くします。"),
            ("長点比率 (Ratio)", "dashRatio", float, "短点に対する長点の長さの比率 (標準: 3.0)"),
            ("連続入力 (Key Repeat)", "keyRepeat", bool, "トリガー長押しで連続入力を有効にします。")
        ])
        
        # --- Spacer ---
        ttk.Frame(content, height=20, style="TFrame").pack()

        # --- Audio & Logic Card ---
        self._create_card(content, "システム・判定", [
            ("周波数 (Hz)", "freq", int, "再生されるビープ音の高さ。"),
            ("トリガー感度 (0.0-1.0)", "triggerThreshold", float, "入力と判定されるトリガーの押し込み深さ。"),
            ("文字確定待ち時間 (秒)", "charTimeout", float, "入力停止から文字が確定・送信されるまでの時間。"),
            ("右スティックトリガー", "requireStickDown", bool, "右スティックを下に倒している間のみ入力を受け付けます。"),
            ("デバッグモード", "debugMode", bool, "詳細なログをコンソールに出力します。")
        ])

        # --- Spacer ---
        ttk.Frame(content, height=20, style="TFrame").pack()

        # --- Custom Map Card ---
        self._create_dict_editor_card(content, "カスタム定義", "customMorseMap")

        # --- Spacer ---
        ttk.Frame(content, height=20, style="TFrame").pack()

        # --- Overlay Card ---
        overlay_desc = "VR空間内に現在の入力文字を表示します。"
            
        self._create_card(content, "オーバーレイ設定", [
            ("有効にする", "overlayEnabled", bool, overlay_desc),
            ("幅 (メートル)", "overlayWidth", float, "オーバーレイの物理的な幅 (1.0 = 1メートル)"),
            ("高さ位置 (Y)", "overlayY", float, "視界に対する上下位置 (マイナスで下)"),
            ("奥行き (Z)", "overlayZ", float, "視界に対する奥行き (マイナスで前方)"),
            ("不透明度 (0.0-1.0)", "overlayOpacity", float, "背景の透け具合")
        ])
        
        # Final binding after widgets creation
        _bind_to_children(content)
        
        self.frames["Settings"] = container

    def _init_about_view(self):
        container = ttk.Frame(self.main_area, style="TFrame", padding=40)
        container.pack_propagate(False)
        container.pack(fill=tk.BOTH, expand=True)

        # Header
        header_frame = ttk.Frame(container, style="TFrame")
        header_frame.pack(fill=tk.X, pady=(0, 20))
        ttk.Label(header_frame, text="情報", style="Header.TLabel").pack(side=tk.LEFT)

        # Content Card
        card = ttk.Frame(container, style="Card.TFrame", padding=40)
        card.pack(fill=tk.BOTH, expand=True)
        
        # Center Content
        center_frame = ttk.Frame(card, style="Card.TFrame")
        center_frame.place(relx=0.5, rely=0.5, anchor="center")

        ttk.Label(center_frame, text="VR Morse Input", style="Header.TLabel", font=("Yu Gothic UI", 24, "bold")).pack(pady=(0, 10))
        ttk.Label(center_frame, text="Version 1.0.0", style="Card.TLabel").pack(pady=(0, 20))
        
        ttk.Label(center_frame, text="VR空間でモールス信号を使用してテキストを入力するためのツールです。", style="Card.TLabel").pack(pady=5)
        ttk.Label(center_frame, text="Created by Antigravity", style="Description.TLabel").pack(pady=(20, 0))
        
        self.frames["About"] = container
    
    def _create_card(self, parent, title, fields):
        card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        card.pack(fill=tk.X)
        
        ttk.Label(card, text=title, style="SubHeader.TLabel").pack(anchor="w", pady=(0, 15))
        
        # Grid layout for fields
        grid_frame = ttk.Frame(card, style="Card.TFrame")
        grid_frame.pack(fill=tk.X)
        grid_frame.columnconfigure(1, weight=1)
        grid_frame.columnconfigure(2, weight=0) # Help text or unit
        
        row = 0
        for label, key, dtype, desc in fields:
            # Label
            lbl = ttk.Label(grid_frame, text=label, style="Card.TLabel")
            lbl.grid(row=row, column=0, sticky="nw", pady=10, padx=(0, 20))
            
            # Input
            val_frame = ttk.Frame(grid_frame, style="Card.TFrame")
            val_frame.grid(row=row, column=1, sticky="ew", pady=5)
            
            if dtype == bool:
                var = tk.BooleanVar()
                chk = ttk.Checkbutton(val_frame, text="ON/OFF", variable=var, style="Modern.TCheckbutton")
                chk.pack(anchor="w")
                self.vars[key] = (var, dtype)
                var.trace_add("write", self.schedule_save)
            else:
                var = tk.StringVar()
                entry = ttk.Entry(val_frame, textvariable=var, style="Modern.TEntry")
                entry.pack(fill=tk.X)
                self.vars[key] = (var, dtype)
                var.trace_add("write", self.schedule_save)
                
            # Description (Small)
            desc_lbl = ttk.Label(grid_frame, text=desc, style="Description.TLabel")
            desc_lbl.grid(row=row+1, column=0, columnspan=2, sticky="w", pady=(0, 15))
            
            row += 2

    def _create_dict_editor_card(self, parent, title, setting_key):
        card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        card.pack(fill=tk.X)
        
        ttk.Label(card, text=title, style="SubHeader.TLabel").pack(anchor="w", pady=(0, 15))
        
        # --- List Area ---
        list_frame = ttk.Frame(card)
        list_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Treeview for custom map
        columns = ("morse", "char")
        tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=5)
        tree.heading("morse", text="モールス信号 (例: ...)")
        tree.heading("char", text="文字 (例: ?)")
        tree.column("morse", width=150)
        tree.column("char", width=100)
        tree.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # --- Edit Area ---
        edit_frame = ttk.Frame(card, style="Card.TFrame")
        edit_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(edit_frame, text="信号:", style="Card.TLabel").pack(side=tk.LEFT, padx=(0, 5))
        ent_morse = ttk.Entry(edit_frame, width=15)
        ent_morse.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Label(edit_frame, text="文字:", style="Card.TLabel").pack(side=tk.LEFT, padx=(0, 5))
        ent_char = ttk.Entry(edit_frame, width=10)
        ent_char.pack(side=tk.LEFT, padx=(0, 10))
        
        def _refresh_list():
            for item in tree.get_children():
                tree.delete(item)
            current_map = state.settings.get(setting_key, {})
            # Sort by key length for easier reading
            for m in sorted(current_map.keys(), key=lambda k: (len(k), k)):
                c = current_map[m]
                tree.insert("", "end", values=(m, c))
        
        def _add_item():
            m = ent_morse.get().strip()
            c = ent_char.get().strip()
            if m and c:
                # Update State
                if setting_key not in state.settings:
                    state.settings[setting_key] = {}
                state.settings[setting_key][m] = c
                _refresh_list()
                ent_morse.delete(0, tk.END)
                ent_char.delete(0, tk.END)
                self.schedule_save()
        
        def _delete_item():
            selected = tree.selection()
            if selected:
                item = tree.item(selected[0])
                m = str(item['values'][0]) # values coming back can be int if looking like int
                if setting_key in state.settings and m in state.settings[setting_key]:
                    del state.settings[setting_key][m]
                    _refresh_list()
                    self.schedule_save()

        btn_add = ttk.Button(edit_frame, text="追加/更新", style="Action.TButton", command=_add_item)
        btn_add.pack(side=tk.LEFT)
        
        btn_del = ttk.Button(card, text="選択項目を削除", style="Nav.TButton", command=_delete_item)
        btn_del.pack(anchor="e", pady=(5, 0))

        # Initial Load
        _refresh_list()
        
        # Store refresh function to call it later if needed (e.g. on populate)
        # Using a special key format to distinguish from normal vars
        self.vars[setting_key + "_refresh"] = (_refresh_list, None)

    def populate_settings(self):
        for key, (var, _) in self.vars.items():
            if key.endswith("_refresh"):
                # Call refresh function for custom editors
                var()
                continue
                
            val = state.settings.get(key, "")
            var.set(val)

    def schedule_save(self, *args):
        if hasattr(self, "_save_timer"):
            self.after_cancel(self._save_timer)
        self._save_timer = self.after(500, self.apply_settings)

    def apply_settings(self, *args):
        try:
            for key, (var, t_func) in self.vars.items():
                if t_func is None: continue

                try:
                    val = var.get()
                    if t_func == bool:
                        state.settings[key] = val
                    else:
                        state.settings[key] = t_func(val)
                except ValueError:
                    pass # Ignore invalid input during typing
            
            save_settings()
            recalculate_derived_values()
            
            # Update Overlay Live
            if hasattr(state, 'overlay') and state.overlay:
                # Update enabled state if changed (might need re-init but simpler to just toggle visibility logic)
                state.overlay.enabled = state.settings.get("overlayEnabled", False) and state.overlay.use_gdi
                if state.overlay.enabled:
                    if state.overlay.vr_overlay:
                         try:
                             state.overlay.vr_overlay.showOverlay(state.overlay.handle)
                             state.overlay.vr_overlay.setOverlayWidthInMeters(state.overlay.handle, state.settings.get("overlayWidth", 1.0))
                             state.overlay.vr_overlay.setOverlayAlpha(state.overlay.handle, state.settings.get("overlayOpacity", 0.8))
                             state.overlay.update_transform()
                             # Re-render current text to apply new size/alpha potential
                             state.overlay.last_text = None 
                             current_text = state.fixed_text + state.text_buffer
                             if state.conversion_active and state.conversion_candidates:
                                 current_text = state.fixed_text + state.conversion_candidates[state.conversion_index]
                             state.overlay.update_image(current_text if current_text else "Ready")
                         except: pass
                else:
                    if state.overlay.vr_overlay:
                        try: state.overlay.vr_overlay.hideOverlay(state.overlay.handle)
                        except: pass

            self.log("設定を保存しました。", "INFO")
        except Exception as e:
            self.log(f"設定エラー: {e}", "ERROR")

    def show_frame(self, name):
        for f in self.frames.values():
            f.pack_forget()
        self.frames[name].pack(fill=tk.BOTH, expand=True)

    def log(self, text, level="INFO"):
        """Thread-safe logging"""
        def _update():
            self.console.configure(state='normal')
            timestamp = time.strftime("[%H:%M:%S] ")
            self.console.insert(tk.END, timestamp + f"{level.ljust(5)} ", level)
            self.console.insert(tk.END, text + "\n", level)
            self.console.see(tk.END)
            self.console.configure(state='disabled')
        self.after(0, _update)

    def reset_ui_state(self):
        self.btn_start.configure(text="▶  システム起動", style="Action.TButton")
        
    def toggle_vr(self):
        if not state.running:
            state.running = True
            self.btn_start.configure(text="■  システム停止", style="Action.TButton")
            threading.Thread(target=vr_loop, args=(self,), daemon=True).start()
        else:
            state.running = False
            self.btn_start.configure(text="停止中...")

if __name__ == "__main__":
    app = MorseApp()
    app.mainloop()
    state.running = False
