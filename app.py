"""
Christiani & Nielsen (Thai) Public Company Limited
CNT STAFF TIME AUTOMATION SUITE (Universal 2026 Edition)
=========================================================
ระบบบันทึกเวลาทำงานและประมวลผลการลาอัตโนมัติ สำหรับทุกหน่วยงานและโครงการของ CNT
ออกแบบตามหลัก Modern Enterprise UI 2026 & Senior-Friendly Typography

Features:
1. 🏖️ บันทึกการลาที่ได้รับการอนุมัติอัตโนมัติ (Leave Automation)
2. ⏰ บันทึกใบแจ้งการมาทำงาน (Attendance Notification / ลืมสแกนนิ้ว) พร้อมกฎปัดเศษลงช่วงละ 30 นาที
3. ⚙️ ระบบตั้งค่าคอลัมน์เป้าหมายในไฟล์ STAFF TIME แบบยืดหยุ่น (Custom Columns)
4. 🛡️ ระบบกันพลาด 5 ชั้น (Safeguards: สแกนครบ, Double-Check ชื่อ-รหัส, ตรวจข้อมูลขัดแย้ง, สแกนเดี่ยวซ้ำซ้อน, Audit Report)
5. 🔍 ตรวจสอบและแสดงรายการแก้ไขทั้งหมดก่อนบันทึกจริง (Pre-Check Preview All Changes)
6. 👤 ระบบตรวจสอบรายคน (Employee Quick Inspector)
7. 📈 สถิติภาพรวมโครงการ (STAFF TIME Project Analytics)
8. 📖 คู่มือการใช้งานแบบละเอียดฝังในโปรแกรม (Interactive User Manual)
9. 💾 บันทึกและจดจำการตั้งค่าและไฟล์ล่าสุดอัตโนมัติ (Auto-save Settings)
10. 🕒 ไอคอนนาฬิกาประจำโปรแกรม ไร้ไอคอนขนนก
11. 🔍 ฟอนต์ Tahoma ตัวใหญ่คมชัด อ่านสบายตาทั่วทั้งโปรแกรม
"""

import os
import sys
import io
import re
import json
import math
import zipfile
import shutil
import subprocess
import threading
import urllib.request
import urllib.error
from datetime import datetime, timedelta, time, date
import tkinter as tk
from tkinter import filedialog, messagebox

# เปิด High-DPI ให้ตัวหนังสือคมชัด ไม่เบลอบนจอความละเอียดสูง
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

# ตั้งค่า AppUserModelID เพื่อให้ Windows Taskbar แสดงไอคอนนาฬิกา
try:
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("cnt.stafftime.automation.ohm282")
except Exception:
    pass

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from PIL import Image, ImageDraw, ImageTk
import customtkinter as ctk

# ตั้งค่าธีม CustomTkinter
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# ข้อมูลเวอร์ชันและการอัปเดต (Auto-Update Configuration)
APP_VERSION = "1.2.1"
DEFAULT_GITHUB_REPO = "OHM282/LeaveAutomation-CNT"


# -------------------------------------------------------------
# ตัวช่วยจัดการไอคอนนาฬิกา & พาธไฟล์ (Clock Icon & Path Helpers)
# -------------------------------------------------------------

def get_resource_path(relative_name):
    """ดึง Path ของไฟล์ Resource รองรับทั้งขณะ Dev และใน PyInstaller .exe"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_name)
    return os.path.join(os.path.abspath(os.path.dirname(__file__)), relative_name)


def get_settings_file_path():
    """ดึง Path ของไฟล์ settings.json วางข้าง .exe หรือสคริปต์หลัก"""
    if hasattr(sys, '_MEIPASS'):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, "settings.json")


def load_app_settings():
    """โหลดการตั้งค่าที่บันทึกไว้ล่าสุด"""
    try:
        p = get_settings_file_path()
        if os.path.exists(p):
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def save_app_settings(data):
    """บันทึกการตั้งค่าลง settings.json"""
    try:
        p = get_settings_file_path()
        with open(p, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def center_window(window, width, height, parent=None):
    """จัดตำแหน่งหน้าต่างให้อยู่กึ่งกลางหน้าจอหรือกึ่งกลางหน้าต่างแม่"""
    window.update_idletasks()
    if parent:
        parent.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - width) // 2
        y = parent.winfo_y() + (parent.winfo_height() - height) // 2
    else:
        sw = window.winfo_screenwidth()
        sh = window.winfo_screenheight()
        x = (sw - width) // 2
        y = (sh - height) // 2
    x = max(20, x)
    y = max(20, y)
    window.geometry(f"{width}x{height}+{x}+{y}")


# -------------------------------------------------------------
# ระบบคำแนะนำแบบลอยเมื่อเมาส์ชี้ (Modern Tooltip - UI/UX Pro Max)
# -------------------------------------------------------------

class ModernToolTip:
    """กล่องคำแนะนำแบบลอยเมื่อนำเมาส์ไปชี้ (Hover Tooltip - UI/UX Pro Max)"""
    def __init__(self, widget, text, delay=350):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tipwindow = None
        self.id = None

        # รองรับทั้งวิดเจ็ตมาตรฐานและ CustomTkinter (เช่น CTkSegmentedButton, CTkEntry, CTkTextbox)
        targets = []
        if hasattr(widget, '_canvas'):
            targets.append(widget._canvas)
        if hasattr(widget, '_entry'):
            targets.append(widget._entry)
        if hasattr(widget, '_textbox'):
            targets.append(widget._textbox)
        if hasattr(widget, '_buttons_dict'):
            targets.extend(widget._buttons_dict.values())
        targets.append(widget)

        for t in targets:
            try:
                t.bind("<Enter>", self._on_enter, add="+")
                t.bind("<Leave>", self._on_leave, add="+")
                t.bind("<ButtonPress>", self._on_leave, add="+")
            except Exception:
                pass

    def _on_enter(self, event=None):
        self._schedule()

    def _on_leave(self, event=None):
        self._unschedule()
        self._hide()

    def _schedule(self):
        self._unschedule()
        self.id = self.widget.after(self.delay, self._show)

    def _unschedule(self):
        tid = self.id
        self.id = None
        if tid:
            try:
                self.widget.after_cancel(tid)
            except Exception:
                pass

    def _show(self):
        if not self.widget.winfo_exists():
            return
        x = self.widget.winfo_rootx() + 10
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        try:
            tw.attributes("-topmost", True)
        except Exception:
            pass
        frm = tk.Frame(tw, bg="#0f172a", padx=9, pady=5, relief="flat", highlightthickness=1, highlightbackground="#334155")
        frm.pack()
        lbl = tk.Label(frm, text=self.text, justify='left', bg="#0f172a", fg="#f8fafc", font=("Tahoma", 10, "normal"))
        lbl.pack()

    def _hide(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            try:
                tw.destroy()
            except Exception:
                pass


def add_tooltip(widget, text):
    """ฟังก์ชันผู้ช่วยผูก Tooltip เข้ากับ Widget"""
    if widget and text:
        return ModernToolTip(widget, text)
    return None


def ensure_clock_icon():
    """สร้างไฟล์ clock.ico หากยังไม่มี เพื่อใช้เป็นไอคอนประจำโปรแกรม"""
    icon_path = get_resource_path("clock.ico")
    if os.path.exists(icon_path):
        return icon_path
        
    try:
        target_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "clock.ico")
        size = 512
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        
        # Outer slate ring & blue border
        margin = 24
        draw.ellipse([margin, margin, size - margin, size - margin], fill="#1e293b", outline="#2563eb", width=20)
        
        # Dial
        inner_m = margin + 26
        draw.ellipse([inner_m, inner_m, size - inner_m, size - inner_m], fill="#ffffff", outline="#e2e8f0", width=8)
        
        center = size // 2
        radius = (size - 2 * inner_m) // 2
        
        # Ticks
        for h in range(12):
            ang = math.radians(h * 30 - 90)
            is_maj = (h % 3 == 0)
            t_len = 32 if is_maj else 18
            t_w = 10 if is_maj else 5
            col = "#1e293b" if is_maj else "#94a3b8"
            x1 = center + int((radius - 16) * math.cos(ang))
            y1 = center + int((radius - 16) * math.sin(ang))
            x2 = center + int((radius - 16 - t_len) * math.cos(ang))
            y2 = center + int((radius - 16 - t_len) * math.sin(ang))
            draw.line([x1, y1, x2, y2], fill=col, width=t_w)
            
        # Hands (7:30)
        h_ang = math.radians(225 - 90)
        h_len = radius * 0.52
        draw.line([center, center, center + int(h_len * math.cos(h_ang)), center + int(h_len * math.sin(h_ang))], fill="#1e293b", width=16)
        
        m_ang = math.radians(180 - 90)
        m_len = radius * 0.72
        draw.line([center, center, center + int(m_len * math.cos(m_ang)), center + int(m_len * math.sin(m_ang))], fill="#f59e0b", width=12)
        
        # Center gold pin
        draw.ellipse([center - 18, center - 18, center + 18, center + 18], fill="#f59e0b", outline="#b45309", width=4)
        draw.ellipse([center - 8, center - 8, center + 8, center + 8], fill="#ffffff")
        
        sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
        img.save(target_path, format="ICO", sizes=sizes)
        return target_path
    except Exception:
        return None


CNT_LOGO_B64 = "iVBORw0KGgoAAAANSUhEUgAAAGsAAABOCAYAAAAn3uiJAAAJYklEQVR4nO1d3XHbOBCGM3kPOzilgjAVnFxBZDZg+cHPsSqwVYHsZz2IaUBWKhBTQXgVHK+C41WQm+XumiAIgKAEUSSjneFoSJEAiA/7A2B3eSWGSNF6IoSYCiE+CSFCOgLLExkdP4QQaXFs7+F8UHQlhkLRGgC5FULMhBAA1rEEoH0TQuyGAlz/wYrWcyHEV+KeU1EihFiK7T389paueg7SoycuGgVoVz0VdyvSSeeinRBi0Tfx2C+wovUTcZMrJWQ0sAGRVToYgQ+IOz/RAHAVpzkBFoueUD/AitbQoa+O3LQjwyAR2/v8wLrASPlCv00UE2jt6xodWDj6Xxt0Uy6EeBFCPHvtNJwCACeDfmyyHK/PDdhVD4DaN8yRnknp5ydsx4T05KzPgF31GCjonDuxvU8dOjqUdNGf0r+gz1i3pY0dHa1BDG8sXH5WwM4DFnbwTwtQdj0RHTxBTgm4F6OlhzptY+GyswHWPVjYGXuLVXZntMCiYuQ/ejLrExoQes6N1huLLoNVjxvRMb3rukLSDe2AAk6M1gDw3uP8a2rVldv7uwJMPc1omjFizorWM7L82gD1QAD7JjD9rxvvsnPY50ad6pHed1WRpAt0tDQAZesolaDTQA/9JV0DYyMwcPJSuBBwWLQWhnZA+z6L0YGFuiYwjPCnA4HaOU2Qy4nwLYm/pOX630KxOJnCgvO39zC9GIkYROvvb80/0MEfax3dDFRM3Nh+7Q4tSdFafOFzYMG6vcOAOcu03lc3z6P1ygJURrotUe63GSxVQHUg4Qo/cF2VZJ0Gz0XrpeZdgGtBrz4NHywUQbrOT2p6Ck1zePE285vQYiEChzYbEThXc7EynwlUdW532wVYXZjups5ftjBA0gMnolMaAH4I69cZJjC1cDWEeg1WXbzg0k+iAXVyghWDlfBJKA10bYFV/AGDhUpZB8CLhqtg614l6JSbI5V3eIJRHxsmyrYF6d5z1sxicqtcpXvRwyy+Oq08d2R1sJU0HTJY8gq4vK6WO4jKzOP8JbDozvaEAyh1fN/BgKUbad8dRaXbCoM7PdJ8zxcl4+EsnnzWKT1QVPogn8YG75XJFA6Vs/Q6oj4pdRWVbSk3GAG+Rr9payUcIlhTR9ExcRy1bcm0vdHGe8pMZsMnGNN+lkqmuZWP+VBimCj7MuV17ZyMBazqaDSb076cKxcWY8MHB+RjBusf5Vwv3315wm4L/RgbOtSfKd8R9UEMnpoWBg746tmUHx1Yf1TOTBuAPi2qrXHxNfBmbIwULNeR7Nei2hYrITpjYH6kKX8a4+gMYOWOIGQdrQQsDNcfPYOVDxEs3QjTibdu1ti2hciNDQOj/faGWVRnQwQrc3zJH4a5UNChsREeUNZpLdlOwcJG5w4izrQG6N+03hqNjUNIx/3JkA0MXeO/OG43fD0Jd20LY8PH6Nfp1XTIYLmKuJeT70FVCVyjDyfzto6PNc2zgWUScXNHvwZYFpqeyNjYHVGC3gVhe3+KbZ2OwLKJOHfT+rVjY8NOpXevSicFqqtJ8YuT6xZylw5YDBGKPAOGA8nkS2Ejk78IuHEPHqydUcTpdYnJtP7pfWMPfezdjQ2zF1bWRe6M04OF5rKJu540q+QmcTghDns4o7GxMnphjSgwIaDAhMApxgn9122gpOSm1l5PAHceElNlji0DrvooRhVMZ05Iove4dQv5yUjMfreKIbQov1B5QesgOPtgAyfUkxsX54h8hJAZnd6JKSxUvX9u8X/XEQCQK6LzOJdsewx0p7HFXW+RmPQDbFXUQUEL8bqFERDSygIfE8t9rqvtppAiTBfUIXULFoqeZUvAEgoFffbcmofGCbddFNdjv0aaB+PV4typF4nV9D3w7DHzLnt6oeZcUvBcp1x1TrCacmGkpLgzh2RZU0fgoKyEjBGzQYBzuY2lbebBNOJ0QE2AuaeQiwqO47RAgVIGRvG7iKzmFHq/YTogd8C6ya4ZFbpr5dCOY2PFBp/CzjXXYOIdtDLwu6nus4m+foHlvmqhToS/HbgSwRNk1yRdi67yXAwHrHJJZ9PC0stJj8ip6lQCcD40RPUfnkLvtwWrFIuPZ3RvzinFXeeJtIYHVtXC23SchTr2GMf8G4HlfyJs46TYmjCyJ9R/sI6bCB83Qe4ZDQcs/UoDexmxD586KeZ870LK/570nYMudKELXehCF7rQhS50oQtd6EJdTorLxdOQll9g+2En7ffgZyDkc/zlzbqMclzs6L49Zetc0Eo63BvQffAsX9P5i99S/bGSsPjubTKL1yZ0LadzQfXtK88zYTvA9RkD6uqOpfxFvBd6dz7nNOG3RYJjfQJlaDe0bVpbAC6XyyaGvtV9NHT19g5SO9i7aS+tck8pcoNdueSlHflc3nLgBu2lMkLp21j8cnPFq1UtP6i4kOEgeqBrc+k5KG8mtVnOuV53QcMO43r580sqcd0bqpfP+ct2/K66/O5cr84lQP4IDfZHGd+lW6TmeldqO94RwiFxzzW5fbnvxsIz+FysWa/jcyjvjo4fxYjBZzjygutW944YoNzAiW45BMuQWe5429a8S36M9O298bD5iYSSN9UNueKV9VfLWSjtWKmpwvllU+nFML126VcHAQH6pqDYAZpqOhs/4IydPpfEoCuB2OLQnFUhysqFV7jO2yh2wpEMHcS/ti36jDj2P8s9EL35Szq3pSOPqbxXqrsaVVMtZykxSkZ99o/dybPu/Ij7PHqOAzCmb8mHZULnkgVdh61x9LdwibUqxXAmZaa5le7IqE0uH93k3ecd/UL40L+Ge3nk60J71Lr5sC0Mf1M4igeuXN9S078c7Pcog5VWOgLFInCSrChB2T1pY2aRdZc6tqUy9tJ93BgX3wcZmJA6BJI7ys9yMHcT+CGttt9IHGkS9czJ9s+/Q3+UR9bwUTcwTp4kqVKWXS1HblMts8A7EisxuS//olHIG3JuhA1JNZkyOVb4lcqekbxv8mv4QPfiZ5NQny3ov7mGc5toR20DbmKw4Xze8D52MVge5cCWr2Nbue5fkpEjf55DLmevtONZvve99Nmh75JMj8kk5gSL3HD5fKGMvhtJeV9TQHQmovVnaZeXPZOY1PLRiRLLgPaUI3ZbmLv8X1lG9TpQPZABOAqBYbGKHVg1DLhubgu/D3/qiTtNfW8grk/lVnQuLevGd8ZrcplM8A5sM3A7QL8W7fgfBq7SKQwAdzYAAAAASUVORK5CYII="

def ensure_cnt_logo():
    """โหลดหรือสร้างไฟล์ cnt_logo.png เพื่อใช้เป็นโลโก้บริษัท Christiani & Nielsen"""
    logo_path = get_resource_path("cnt_logo.png")
    if os.path.exists(logo_path):
        return logo_path
    try:
        target_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "cnt_logo.png")
        import base64
        with open(target_path, "wb") as f:
            f.write(base64.b64decode(CNT_LOGO_B64))
        return target_path
    except Exception:
        return None


def get_default_column_config():
    """คืนค่าคอนฟิกตำแหน่งคอลัมน์มาตรฐาน"""
    return {
        'emp_id_col': 'A',
        'emp_name_col': 'B',
        'scan_col': 'E',
        'leave_type_col': 'G',
        'time_in_col': 'J',
        'time_out_col': 'K',
        'checkin_col': 'L',
        'checkout_col': 'M',
        'leave_code_col': 'Q',
        'start_row': 2
    }


# -------------------------------------------------------------
# ฟังก์ชันคำนวณและประมวลผลพื้นฐาน (Backend Helpers)
# -------------------------------------------------------------

def extract_emp_id(emp_str):
    """
    ดึงเฉพาะรหัสพนักงานที่เป็นตัวเลข ตัด 0 นำหน้าออก เช่น 'A05114 : นายนพดล' -> '5114'
    """
    if not emp_str or str(emp_str) == 'nan':
        return ""
    match = re.search(r'\d+', str(emp_str))
    if match:
        digits = match.group(0)
        return digits.lstrip('0')
    return str(emp_str).strip()


def to_decimal_hours(val):
    """
    แปลงค่าเวลาในเซลล์เป็นตัวเลขชั่วโมง เช่น 7:30 -> 7.5, 18:30 -> 18.5
    """
    if val is None or str(val) == 'nan':
        return None
    if isinstance(val, time):
        return val.hour + val.minute / 60.0 + val.second / 3600.0
    if isinstance(val, datetime):
        return val.hour + val.minute / 60.0 + val.second / 3600.0
    if isinstance(val, (int, float)):
        if 0 <= val <= 1.0:
            return val * 24.0
        if 1.0 < val < 24.0:
            return float(val)
        return None
    if isinstance(val, str):
        val = val.strip()
        m = re.search(r'(\d{1,2}):(\d{2})(?::(\d{2}))?', val)
        if m:
            h = int(m.group(1))
            mi = int(m.group(2))
            return h + mi / 60.0
    return None


def calculate_checkout_time(time_val):
    """
    คำนวณเวลาออก (Check-out) ตามระเบียบบริษัท:
    - เริ่มต้นที่ 16:30 เป็นต้นไป
    - ปัดเศษนาทีลงเป็นช่วงละ 30 นาที (:00 หรือ :30)
      เช่น 16:55 -> 16:30
          17:00 -> 17:00
          17:10 -> 17:00 (ปัดเศษลงตามหลักเกณฑ์)
          17:33 -> 17:30
          18:01 -> 18:00
          18:31 -> 18:30
    """
    h, m = None, None
    if isinstance(time_val, (time, datetime)):
        h = time_val.hour
        m = time_val.minute
    elif isinstance(time_val, (int, float)):
        if 0 <= time_val <= 1.0:
            tot_sec = int(round(time_val * 86400))
            h = (tot_sec // 3600) % 24
            m = (tot_sec % 3600) // 60
        else:
            h = int(time_val)
            m = int(round((time_val - h) * 60))
    elif isinstance(time_val, str):
        match = re.search(r'(\d{1,2}):(\d{2})', time_val)
        if match:
            h = int(match.group(1))
            m = int(match.group(2))
            
    if h is None or m is None:
        return time(16, 30)
        
    # เริ่มต้นที่ 16:30 เป็นต้นไป
    if h < 16 or (h == 16 and m < 30):
        return time(16, 30)
        
    # ปัดเศษลงเป็นช่วงละ 30 นาที (:00 หรือ :30)
    if m < 30:
        return time(h, 0)
    else:
        return time(h, 30)


def check_fingerprint_scans(val):
    """
    ตรวจสอบข้อมูลการสแกนนิ้วดิบ (ช่อง E / Time ของไฟล์ STAFF TIME)
    ครอบคลุมช่วงเช้าตั้งแต่ 04:00 - 11:59 (เพื่อรองรับคนมาเช้าตรู่)
    และออกช่วงบ่าย/เย็น >= 12:00
    คืนค่าเป็น (has_in, has_out, scans_list, earliest_dec, latest_dec)
    """
    if val is None or str(val) == 'nan':
        return False, False, [], None, None
    
    val_str = str(val).strip()
    matches = re.findall(r'(\d{1,2}):(\d{2})', val_str)
    scans_list = []
    has_in = False
    has_out = False
    dec_times = []
    
    if matches:
        for m in matches:
            h = int(m[0])
            mi = int(m[1])
            time_str = f"{h:02d}:{mi:02d}"
            if time_str not in scans_list:
                scans_list.append(time_str)
            t_dec = h + mi / 60.0
            dec_times.append(t_dec)
            if 4.0 <= t_dec < 12.0:  # ขยายช่วงเช้าเริ่มต้นที่ 04:00 น.
                has_in = True
            elif t_dec >= 12.0:
                has_out = True
    elif isinstance(val, (datetime, time)):
        h = val.hour
        mi = val.minute
        time_str = f"{h:02d}:{mi:02d}"
        scans_list.append(time_str)
        t_dec = h + mi / 60.0
        dec_times.append(t_dec)
        if 4.0 <= t_dec < 12.0:
            has_in = True
        elif t_dec >= 12.0:
            has_out = True
            
    earliest_dec = min(dec_times) if dec_times else None
    latest_dec = max(dec_times) if dec_times else None

    # กรณีมีสแกนตั้งแต่ 2 เวลาขึ้นไปและห่างกัน >= 1 ชม. (เช่น เข้า 12:54 ออก 18:31 หรือ เข้า 07:18 ออก 11:31)
    # แสดงว่ามีสแกนทั้งขาเข้าและขาออกครบถ้วน
    if len(scans_list) >= 2 and earliest_dec is not None and latest_dec is not None and (latest_dec - earliest_dec) >= 1.0:
        has_in = True
        has_out = True

    return has_in, has_out, scans_list, earliest_dec, latest_dec


def is_valid_time(val):
    """
    ตรวจสอบว่าเป็นเวลาที่ถูกต้อง (ไม่ใช่ None, ไม่ใช่ 'ไม่ได้แสกน'/'ไม่ได้สแกน', และไม่ใช่ 0:00)
    รองรับทั้ง time, datetime, float/int (decimal fraction ของวันใน Excel), และ string
    """
    if val is None:
        return False
    if isinstance(val, (time, datetime)):
        return not (val.hour == 0 and val.minute == 0)
    if isinstance(val, (int, float)):
        return val > 0.001
    s = str(val).strip()
    if not s or s in ['None', 'ไม่ได้แสกน', 'ไม่ได้สแกน', '0:00', '00:00', '0:00:00', '0', '#VALUE!', '-']:
        return False
    m = re.match(r'^(\d{1,2}):(\d{2})', s)
    if m:
        h, mi = int(m.group(1)), int(m.group(2))
        return not (h == 0 and mi == 0)
    return False


def format_time_val(val):
    """จัดรูปแบบค่าเวลาให้เป็น HH:MM สำหรับแสดงผลและใส่ในหมายเหตุ"""
    if isinstance(val, (time, datetime)):
        return f"{val.hour:02d}:{val.minute:02d}"
    s = str(val).strip()
    m = re.match(r'^(\d{1,2}):(\d{2})', s)
    if m:
        return f"{int(m.group(1)):02d}:{int(m.group(2)):02d}"
    return s


def clean_thai_name(name):
    """
    ตัดคำนำหน้าชื่อภาษาไทยและภาษาอังกฤษ และตัดช่องว่างส่วนเกินออก
    เพื่อใช้เปรียบเทียบชื่อพนักงานอย่างแม่นยำ (Double-check)
    """
    if not name or str(name) == 'nan':
        return ""
    s = str(name).strip()
    s = re.sub(r'^[A-Za-z]?\d+\s*[:\-\.]\s*', '', s).strip()
    prefixes = [
        'นางสาว', 'น.ส.', 'น.ส', 'นาย', 'นาง', 
        'MR.', 'MR ', 'MISS.', 'MISS ', 'MRS.', 'MRS '
    ]
    upper_s = s.upper()
    for prefix in prefixes:
        if upper_s.startswith(prefix.upper()):
            s = s[len(prefix):].strip()
            break
    return re.sub(r'\s+', ' ', s).strip()


def check_name_similarity(name1, name2):
    """
    ตรวจสอบความสอดคล้องกันของชื่อพนักงานเพื่อทำ Double-Check
    ป้องกันการบันทึกผิดคนกรณีชีท STAFF TIME พิมพ์รหัสพนักงานผิด
    """
    c1 = clean_thai_name(name1)
    c2 = clean_thai_name(name2)
    if not c1 or not c2:
        return True  # ถ้าไม่มีชื่อในชีท ให้ถือว่าผ่านการตรวจ
        
    parts1 = c1.split()
    parts2 = c2.split()
    
    first1 = parts1[0] if parts1 else ""
    first2 = parts2[0] if parts2 else ""
    
    if first1 and first2 and first1 == first2:
        return True
    if (len(first1) >= 3 and first1 in c2) or (len(first2) >= 3 and first2 in c1):
        return True
        
    return False


def set_cell_time(cell, target_time):
    """
    เขียนค่าเวลาลงเซลล์ให้เป็น datetime.time พร้อมฟอร์แมต 'h:mm'
    เพื่อให้ Excel แสดงผลเป็น 7:30, 16:30, 18:30 ตรงตามต้นฉบับทุกจุด
    """
    cell.value = target_time
    cell.number_format = 'h:mm'


def read_excel_safely(filepath):
    """
    อ่านไฟล์ Excel อย่างปลอดภัย โดยซ่อมแซมแท็ก <fill/> ที่ผิดรูปแบบของระบบภายนอก
    ในหน่วยความจำก่อนโหลด แก้ปัญหา expected <class 'openpyxl.styles.fills.Fill'>
    """
    try:
        with open(filepath, "rb") as f:
            raw_bytes = f.read()

        if raw_bytes[:4] == b"PK\x03\x04":
            zin = zipfile.ZipFile(io.BytesIO(raw_bytes), "r")
            zout_buf = io.BytesIO()
            with zipfile.ZipFile(zout_buf, "w") as zout:
                for item in zin.infolist():
                    buf = zin.read(item.filename)
                    if item.filename == "xl/styles.xml":
                        content = buf.decode("utf-8", errors="ignore")
                        content = content.replace("<fill/>", '<fill><patternFill patternType="gray125"/></fill>')
                        buf = content.encode("utf-8")
                    zout.writestr(item, buf)
            zin.close()
            zout_buf.seek(0)
            return pd.read_excel(zout_buf, header=None)
        else:
            return pd.read_excel(filepath, header=None)
    except Exception:
        return pd.read_excel(filepath, header=None)


def map_workbook_sheets(wb):
    """
    สแกนหาชีทรายวันทั้งหมดในไฟล์ STAFF TIME แล้วจับคู่กับ datetime.date
    """
    sheet_date_map = {}
    for name in wb.sheetnames:
        ws = wb[name]
        clean = name.strip()
        
        m = re.match(r"^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$", clean)
        if m:
            d = int(m.group(1)); mo = int(m.group(2)); y = int(m.group(3))
            if y < 100:
                y += 2000
            try:
                sheet_date_map[date(y, mo, d)] = name
                continue
            except ValueError:
                pass
                
        m2 = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$", clean)
        if m2:
            d = int(m2.group(1)); mo = int(m2.group(2)); y = int(m2.group(3))
            if y < 100:
                y += 2000
            try:
                sheet_date_map[date(y, mo, d)] = name
                continue
            except ValueError:
                pass

        # ตรวจสอบเนื้อหาในเซลล์หัวตาราง
        found_cell_date = None
        for cell_coord in ["D2", "D3", "E2", "C2", "B2"]:
            cval = ws[cell_coord].value
            if cval is None:
                continue
            if isinstance(cval, datetime):
                found_cell_date = cval.date()
                break
            elif isinstance(cval, date):
                found_cell_date = cval
                break
            else:
                m_str = re.search(r"(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})", str(cval))
                if m_str:
                    cd = int(m_str.group(1)); cmo = int(m_str.group(2)); cy = int(m_str.group(3))
                    if cy < 100:
                        cy += 2000
                    try:
                        found_cell_date = date(cy, cmo, cd)
                        break
                    except ValueError:
                        pass
        if found_cell_date:
            sheet_date_map[found_cell_date] = name

    return sheet_date_map


def clean_leave_text(raw_val):
    """
    ตัดชื่อวันลาภาษาไทยเอาเฉพาะข้อความก่อนเครื่องหมาย '-' เช่น 'ลาป่วย - Sick Leave' -> 'ลาป่วย'
    """
    if not raw_val or str(raw_val) == 'nan':
        return ""
    text = str(raw_val).strip()
    if '-' in text:
        text = text.split('-')[0].strip()
    return text


def get_leave_code(leave_name):
    """
    แปลงประเภทการลาเป็นตัวอักษรลงช่อง Q:
    - ลาพักร้อน -> พ
    - ลาป่วย -> ป
    - ลากิจ -> อ
    - ลาไม่รับเงินเดือน / ลาโดยไม่ได้รับค่าจ้าง / อื่นๆ -> อ
    """
    if "ลาพักร้อน" in leave_name:
        return "พ"
    elif "ลาป่วย" in leave_name:
        return "ป"
    elif "ลากิจ" in leave_name:
        return "อ"
    elif "ไม่รับเงินเดือน" in leave_name or "ไม่ได้รับค่าจ้าง" in leave_name:
        return "อ"
    return "อ"


def parse_leave_duration(dur_raw):
    """
    แยกวิเคราะห์ระยะเวลาการลา เช่น '0 วัน 4 ชั่วโมง 150 นาที' -> is_partial=True, hrs=4.25
    แก้ปัญหาระบบภายนอกส่งนาทีมีเลขศูนย์ต่อท้าย (150 -> 15 นาที, 450 -> 45 นาที, 300 -> 30 นาที)
    """
    if dur_raw is None:
        return False, None
    dur_str = str(dur_raw).strip()
    if not dur_str or dur_str.lower() == 'nan':
        return False, None

    m_d = re.search(r'(\d+)\s*วัน', dur_str)
    m_h = re.search(r'(\d+)\s*ชั่วโมง', dur_str)
    m_m = re.search(r'(\d+)\s*นาที', dur_str)

    if m_d or m_h or m_m:
        days_cnt = int(m_d.group(1)) if m_d else 0
        hours_cnt = int(m_h.group(1)) if m_h else 0
        mins_cnt = int(m_m.group(1)) if m_m else 0
        # แก้ปัญหาระบบภายนอกส่งนาทีมีเลข 0 เกินมา เช่น 150 นาที = 15 นาที, 450 นาที = 45 นาที
        if mins_cnt in [150, 450, 300] or (mins_cnt >= 60 and mins_cnt % 10 == 0 and (mins_cnt // 10) < 60):
            mins_cnt = mins_cnt // 10
        if days_cnt == 0 and (hours_cnt > 0 or mins_cnt > 0):
            hrs = round(hours_cnt + mins_cnt / 60.0, 2)
            return True, hrs
        return False, None

    try:
        f_dur = float(dur_str)
        if 0 < f_dur < 1.0:
            hrs = round(f_dur * 8.0, 2)
            return True, hrs
    except Exception:
        pass

    return False, None


COMPANY_HOLIDAYS_2026 = {
    date(2026, 1, 1): "วันขึ้นปีใหม่",
    date(2026, 1, 2): "วันหยุดทำการเพิ่มเป็นกรณีพิเศษ",
    date(2026, 3, 3): "วันมาฆบูชา",
    date(2026, 4, 13): "วันสงกรานต์",
    date(2026, 4, 14): "วันสงกรานต์",
    date(2026, 4, 15): "วันสงกรานต์",
    date(2026, 5, 1): "วันแรงงานแห่งชาติ",
    date(2026, 6, 1): "ชดเชยวันวิสาขบูชา",
    date(2026, 6, 3): "วันเฉลิมพระชนมพรรษา สมเด็จพระนางเจ้าสุทิดาฯ",
    date(2026, 7, 28): "วันเฉลิมพระชนมพรรษา พระบาทสมเด็จพระเจ้าอยู่หัว",
    date(2026, 8, 12): "วันเฉลิมพระชนมพรรษา สมเด็จพระนางเจ้าสิริกิติ์ฯ และวันแม่แห่งชาติ",
    date(2026, 10, 13): "วันนวมินทรมหาราช",
    date(2026, 10, 23): "วันปิยมหาราช",
    date(2026, 12, 5): "วันคล้ายวันพระบรมราชสมภพ ร.9 / วันชาติ / วันพ่อแห่งชาติ",
    date(2026, 12, 31): "วันสิ้นปี",
}


def is_company_holiday(target_date, ws=None):
    """
    ตรวจสอบว่าเป็นวันหยุดบริษัทหรือไม่:
    1. ตรวจสอบตามปฏิทินวันหยุดประจำปี 2569 (2026) ของ บมจ.คริสเตียนีและนีลเส็น (ไทย)
    2. ตรวจสอบจากในชีทว่ามีการกรอก 'น' ในช่อง Q เป็นจำนวนมาก (>= 3 แถว) หรือไม่
    """
    if isinstance(target_date, datetime):
        target_date = target_date.date()
    elif not isinstance(target_date, date):
        return False, None

    if target_date in COMPANY_HOLIDAYS_2026:
        return True, COMPANY_HOLIDAYS_2026[target_date]

    # ตรวจสอบวันหยุดประจำปีแบบระบุเดือน-วัน
    month_day_holidays = {
        (1, 1): "วันขึ้นปีใหม่",
        (1, 2): "วันหยุดทำการเพิ่มเป็นกรณีพิเศษ",
        (4, 13): "วันสงกรานต์",
        (4, 14): "วันสงกรานต์",
        (4, 15): "วันสงกรานต์",
        (5, 1): "วันแรงงานแห่งชาติ",
        (6, 3): "วันเฉลิมพระชนมพรรษา สมเด็จพระนางเจ้าสุทิดาฯ",
        (7, 28): "วันเฉลิมพระชนมพรรษา พระบาทสมเด็จพระเจ้าอยู่หัว",
        (8, 12): "วันเฉลิมพระชนมพรรษา สมเด็จพระนางเจ้าสิริกิติ์ฯ และวันแม่แห่งชาติ",
        (10, 13): "วันนวมินทรมหาราช",
        (10, 23): "วันปิยมหาราช",
        (12, 5): "วันพ่อแห่งชาติ / วันชาติ",
        (12, 31): "วันสิ้นปี",
    }
    md = (target_date.month, target_date.day)
    if md in month_day_holidays:
        return True, month_day_holidays[md]

    # ตรวจสอบจากข้อมูลในชีท หากมีตัวย่อ 'น' ในช่อง Q ตั้งแต่ 3 แถวขึ้นไป
    if ws is not None:
        count_nor = 0
        max_check = min(ws.max_row, 30)
        for r in range(2, max_check + 1):
            val_q = str(ws[f"Q{r}"].value or '').strip()
            if val_q == 'น':
                count_nor += 1
                if count_nor >= 3:
                    return True, "วันหยุดบริษัท (ตรวจพบ 'น' ในชีท)"

    return False, None


def create_audit_report(filepath, audit_rows):
    """
    สร้างไฟล์รายงานผลการตรวจสอบ (Audit & Discrepancy Report) ในรูปแบบ Excel (.xlsx)
    สรุปรายการทั้งหมดที่อัปเดต, ข้ามเพราะสแกนครบ, ไม่พบพนักงาน, หรือข้อมูลขัดแย้ง
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Audit Report"
    
    headers = [
        "ลำดับ", "ชีทวันที่", "รหัสพนักงาน", "ชื่อในไฟล์นำเข้า", 
        "ชื่อในชีท STAFF TIME", "ประเภทข้อมูล", "สถานะผลลัพธ์", "เวลาสแกนจริงในชีท", "รายละเอียดการทำงาน"
    ]
    ws.append(headers)
    
    header_fill = PatternFill(start_color="1A73E8", end_color="1A73E8", fill_type="solid")
    header_font = Font(name="Tahoma", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(1, col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    
    row_font = Font(name="Tahoma", size=10)
    
    for idx, item in enumerate(audit_rows, 1):
        r_vals = [
            idx,
            item.get("sheet", ""),
            item.get("emp_id", ""),
            item.get("input_name", ""),
            item.get("sheet_name", ""),
            item.get("type", ""),
            item.get("status", ""),
            item.get("scans", ""),
            item.get("detail", "")
        ]
        ws.append(r_vals)
        cur_row = ws.max_row
        ws.row_dimensions[cur_row].height = 22
        
        status_val = item.get("status", "")
        status_fill = None
        if "ข้าม" in status_val or "ไม่พบ" in status_val:
            status_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
        elif "ขัดแย้ง" in status_val or "ไม่ตรง" in status_val:
            status_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
        elif "สำเร็จ" in status_val:
            status_fill = PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid")
            
        for c_idx in range(1, len(headers) + 1):
            c = ws.cell(cur_row, c_idx)
            c.font = row_font
            c.border = thin_border
            if status_fill and c_idx == 7:
                c.fill = status_fill
            if c_idx in [1, 2, 3, 6, 7, 8]:
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
    ws.auto_filter.ref = ws.dimensions
    wb.save(filepath)
    wb.close()


def parse_approval_records(df_appr, target_dates, log_callback):
    """
    อ่านไฟล์การลาที่ได้รับการอนุมัติ คัดกรองเฉพาะแถวที่ตรงกับช่วงวันที่ในไฟล์ STAFF TIME
    รองรับทั้งการลาวันเดียว และการลาช่วงยาวข้ามเดือน (เช่น 2026-08-10 To 2026-11-07 รวม 90 วัน)
    """
    header_idx = -1
    for i in range(min(15, len(df_appr))):
        row_str = " ".join([str(x) for x in df_appr.iloc[i].values if pd.notna(x)])
        if "ประเภทการลา" in row_str or "ชื่อพนักงาน" in row_str or "ชื่อ-นามสกุล" in row_str:
            header_idx = i
            break
            
    if header_idx == -1:
        log_callback("⚠️ ไม่พบแถวหัวตารางมาตรฐานในการลาที่ได้รับการอนุมัติ ใช้แถวที่ 2 เป็นค่าเริ่มต้น")
        header_idx = 1

    df_data = df_appr.iloc[header_idx + 1:].copy().reset_index(drop=True)
    
    col_emp = 1    # Col B: ชื่อพนักงาน
    col_type = 2   # Col C: ประเภทการลา
    col_date = 3   # Col D: วันที่ลา
    col_dur = 5    # Col F: จำนวนวันที่ลา
    
    records = []
    seen_entries = set()
    skipped_count = 0
    
    for row_idx, row in df_data.iterrows():
        if len(row) <= max(col_emp, col_type, col_date):
            continue
            
        emp_raw = row[col_emp]
        type_raw = row[col_type]
        date_raw = row[col_date]
        dur_raw = row[col_dur] if len(row) > col_dur else ""
        
        if pd.isna(emp_raw) or pd.isna(type_raw) or pd.isna(date_raw):
            continue
            
        emp_id = extract_emp_id(emp_raw)
        if not emp_id:
            continue
            
        date_matches = re.findall(r'\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\b|\b\d{1,2}[-/]\d{1,2}[-/]\d{4}\b', str(date_raw))
        parsed_dates = []
        for m_str in date_matches:
            for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y']:
                try:
                    parsed_dates.append(datetime.strptime(m_str, fmt).date())
                    break
                except ValueError:
                    pass
                    
        if len(parsed_dates) == 1 and len(row) > 4 and pd.notna(row[4]):
            m_col4 = re.findall(r'\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\b|\b\d{1,2}[-/]\d{1,2}[-/]\d{4}\b', str(row[4]))
            for m_str in m_col4:
                for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y']:
                    try:
                        parsed_dates.append(datetime.strptime(m_str, fmt).date())
                        break
                    except ValueError:
                        pass
                        
        if not parsed_dates:
            continue
            
        start_date = min(parsed_dates)
        end_date = max(parsed_dates)
        
        is_partial, hrs = parse_leave_duration(dur_raw)
                
        cur_d = start_date
        while cur_d <= end_date:
            if cur_d in target_dates:
                entry_key = (emp_id, cur_d.isoformat(), str(type_raw).strip())
                if entry_key in seen_entries:
                    skipped_count += 1
                else:
                    seen_entries.add(entry_key)
                    records.append({
                        'emp_id': emp_id,
                        'emp_raw': str(emp_raw).strip(),
                        'date': cur_d,
                        'leave_type_raw': type_raw,
                        'is_partial': is_partial,
                        'hours': hrs,
                        'start_time_raw': str(date_raw),
                        'end_time_raw': str(date_raw)
                    })
            cur_d += timedelta(days=1)
            
    if skipped_count > 0:
        log_callback(f"ℹ️ พบข้อมูลวันลาที่ซ้ำกันในการลาที่ได้รับการอนุมัติและข้ามไป: {skipped_count} รายการ")
    return records


def parse_attendance_records(df_att, target_dates, log_callback):
    """
    อ่านไฟล์ใบแจ้งการมาทำงาน (CNTIWEB) คัดกรองเฉพาะวันที่ตรงกับไฟล์ STAFF TIME
    """
    header_idx = -1
    for i in range(min(15, len(df_att))):
        row_str = " ".join([str(x) for x in df_att.iloc[i].values if pd.notna(x)])
        if "ชื่อ-นามสกุล" in row_str and "เวลา" in row_str:
            header_idx = i
            break
            
    if header_idx == -1:
        log_callback("⚠️ ไม่พบแถวหัวตารางมาตรฐานในใบแจ้งการมาทำงาน ใช้แถวที่ 4 เป็นค่าเริ่มต้น")
        header_idx = 3

    df_data = df_att.iloc[header_idx + 1:].copy().reset_index(drop=True)
    
    col_emp = 2    # Col C: ชื่อ-นามสกุล
    col_type = 4   # Col E: ประเภท
    col_date = 5   # Col F: วันที่
    col_time = 6   # Col G: เวลาเข้า-ออก
    
    records = []
    seen_keys = set()
    
    for row_idx, row in df_data.iterrows():
        emp_raw = row[col_emp]
        type_raw = row[col_type]
        date_val = row[col_date]
        time_raw = row[col_time]
        
        if pd.isna(emp_raw) or pd.isna(date_val) or pd.isna(time_raw):
            continue
            
        emp_id = extract_emp_id(emp_raw)
        if not emp_id:
            continue
            
        dt = None
        if isinstance(date_val, (datetime, date)):
            dt = date_val.date() if isinstance(date_val, datetime) else date_val
        else:
            m_d = re.search(r'\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\b|\b\d{1,2}[-/]\d{1,2}[-/]\d{4}\b', str(date_val))
            if m_d:
                for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y']:
                    try:
                        dt = datetime.strptime(m_d.group(0), fmt).date()
                        break
                    except ValueError:
                        pass
        if not dt or dt not in target_dates:
            continue
            
        t_dec = to_decimal_hours(time_raw)
        if t_dec is None:
            continue
            
        target_col = None
        target_time_obj = None
        
        # กฎเวลา: เช้าเริ่มตั้งแต่ 04:00 - 11:30 น. (เพื่อรองรับคนมาเช้าตรู่)
        if 4.0 <= t_dec <= 11.5:
            target_col = 'L'
            target_time_obj = time(7, 30)
        elif t_dec > 11.5:  # เย็น เริ่ม 16:30 ปัดเศษลงช่วงละ 30 นาที
            target_col = 'M'
            target_time_obj = calculate_checkout_time(time_raw)
            
        if not target_col:
            continue
            
        rec_key = (emp_id, dt.isoformat(), target_col, str(target_time_obj))
        if rec_key in seen_keys:
            continue
        seen_keys.add(rec_key)
        
        records.append({
            'emp_id': emp_id,
            'emp_raw': emp_raw,
            'date': dt,
            'target_col': target_col,
            'target_time': target_time_obj,
            'orig_time': time_raw
        })
        
    return records


# -------------------------------------------------------------
# แกนกลางประเมินและตัดสินใจ (Shared Business Logic Engine)
# รวมลอจิกเป็นหนึ่งเดียว: ป้องกัน Preview กับ Real Run ไม่ตรงกัน
# -------------------------------------------------------------

def evaluate_all_records(wb, sheet_date_map, leave_records, attendance_records, col_config=None):
    """
    ฟังก์ชันแกนกลางตัวเดียวที่ตัดสินใจตรวจสอบ Safeguard, ตรวจชื่อ-รหัส, ตรวจความขัดแย้ง
    และคำนวณเซลล์ที่จะแก้ไข คืนค่า (summary, evaluated_items) ใช้ร่วมกันทั้ง Preview และ Run จริง
    """
    cfg = col_config or get_default_column_config()
    emp_id_col = cfg.get('emp_id_col', 'A').upper().strip() or 'A'
    emp_name_col = cfg.get('emp_name_col', 'B').upper().strip() or 'B'
    scan_col = cfg.get('scan_col', 'E').upper().strip() or 'E'
    leave_type_col = cfg.get('leave_type_col', 'G').upper().strip() or 'G'
    checkin_col = cfg.get('checkin_col', 'L').upper().strip() or 'L'
    checkout_col = cfg.get('checkout_col', 'M').upper().strip() or 'M'
    leave_code_col = cfg.get('leave_code_col', 'Q').upper().strip() or 'Q'
    try:
        start_row = int(cfg.get('start_row', 2))
    except Exception:
        start_row = 2

    # ดัชนีพนักงานประจำแต่ละชีท
    # ดัชนีพนักงานประจำแต่ละชีท (จับคู่ตรงด้วยรหัสพนักงานเป็นหลัก 100%)
    sheet_emp_id_map = {}
    all_employees = set()

    for dt, sheet_name in sheet_date_map.items():
        ws = wb[sheet_name]
        emp_id_rows = {}
        for row in range(start_row, ws.max_row + 1):
            val_id = ws[f"{emp_id_col}{row}"].value
            val_name = ws[f"{emp_name_col}{row}"].value
            eid = extract_emp_id(val_id) if val_id is not None else ""
            raw_name = str(val_name).strip() if val_name is not None and str(val_name) != 'nan' else ""

            if eid:
                emp_id_rows[eid] = (row, raw_name)
                all_employees.add((eid, raw_name))

        sheet_emp_id_map[sheet_name] = emp_id_rows

    def find_emp_row(sheet_name, emp_id, emp_raw_name):
        id_map = sheet_emp_id_map.get(sheet_name, {})
        if emp_id in id_map:
            row, s_name = id_map[emp_id]
            return row, s_name, 'ok', ""
        return None, "", 'not_found', f"ไม่พบรหัส {emp_id} ในตารางชีทนี้"

    full_day_leaves = {}
    for rec in leave_records:
        if not rec.get('is_partial', False):
            full_day_leaves[(rec['emp_id'], rec['date'])] = rec

    evaluated_items = []
    counts = {
        'total': len(leave_records) + len(attendance_records),
        'update': 0,
        'skip_scan': 0,
        'skip_dup': 0,
        'conflict': 0,
        'mismatch': 0,
        'not_found': 0,
        'leave_updates': 0,
        'leave_skips': 0,
        'attend_updates': 0,
        'attend_skips': 0,
    }

    # 1. ประเมินการลาที่ได้รับการอนุมัติ
    for rec in leave_records:
        sheet_name = sheet_date_map.get(rec['date'])
        if not sheet_name:
            continue
        ws = wb[sheet_name]
        clean_name = clean_leave_text(rec['leave_type_raw'])
        leave_code = get_leave_code(clean_name)
        is_partial = rec.get('is_partial', False)
        hrs = rec.get('hours', None)

        row, s_name, m_type, m_msg = find_emp_row(sheet_name, rec['emp_id'], rec['emp_raw'])

        if m_type == 'not_found':
            counts['not_found'] += 1
            counts['leave_skips'] += 1
            evaluated_items.append({
                'sheet': sheet_name, 'date': rec['date'], 'emp_id': rec['emp_id'],
                'input_name': rec['emp_raw'], 'sheet_name': '-', 'type': 'วันลา',
                'leave_type': clean_name, 'status': '⚠️ ไม่พบพนักงานในชีท', 'status_type': 'not_found',
                'scans': '-', 'target_cells': '-', 'action': 'ข้าม: ไม่พบรหัสและชื่อพนักงานในชีทนี้',
                'detail': m_msg, 'row': None, 'cell_updates': {},
                'audit_entry': {"sheet": sheet_name, "emp_id": rec['emp_id'], "input_name": rec['emp_raw'], "sheet_name": "-", "type": "วันลา", "status": "⚠️ ไม่พบพนักงานในชีท", "scans": "-", "detail": m_msg}
            })
            continue


        scan_val = ws[f"{scan_col}{row}"].value
        h_in, h_out, scans_list, e_min, e_max = check_fingerprint_scans(scan_val)
        scans_str = ", ".join(scans_list) if scans_list else "-"

        if h_in and h_out:
            if not is_partial:
                counts['skip_scan'] += 1
                counts['leave_skips'] += 1
                evaluated_items.append({
                    'sheet': sheet_name, 'date': rec['date'], 'emp_id': rec['emp_id'],
                    'input_name': rec['emp_raw'], 'sheet_name': s_name, 'type': 'วันลา',
                    'leave_type': clean_name, 'status': '⚠️ ข้าม: สแกนนิ้วครบแล้ว', 'status_type': 'skip_scan',
                    'scans': scans_str, 'target_cells': f"{leave_type_col}{row}", 'action': f"ข้าม: พบสแกนเข้า-ออกครบ ({scans_str}) อาจกรอกลาผิดวัน",
                    'detail': f"มีการสแกนนิ้วทั้งเข้าและออก ข้ามการบันทึก '{clean_name}'", 'row': row, 'cell_updates': {},
                    'audit_entry': {"sheet": sheet_name, "emp_id": rec['emp_id'], "input_name": rec['emp_raw'], "sheet_name": s_name, "type": "วันลา", "status": "⚠️ ข้าม: สแกนนิ้วครบแล้ว", "scans": scans_str, "detail": f"สแกนนิ้วครบทั้งเข้า-ออก ({scans_str}) ข้าม '{clean_name}'"}
                })
                continue
            elif e_min is not None and e_max is not None and e_min <= 7.58 and e_max >= 16.5:
                counts['skip_scan'] += 1
                counts['leave_skips'] += 1
                evaluated_items.append({
                    'sheet': sheet_name, 'date': rec['date'], 'emp_id': rec['emp_id'],
                    'input_name': rec['emp_raw'], 'sheet_name': s_name, 'type': 'วันลา',
                    'leave_type': clean_name, 'status': '⚠️ ข้าม: สแกนครบเต็มเวลาปกติ', 'status_type': 'skip_scan',
                    'scans': scans_str, 'target_cells': f"{leave_type_col}{row}", 'action': f"ข้าม: สแกนเต็มเวลาปกติ ({scans_str}) ข้าม '{clean_name}'",
                    'detail': f"สแกนเข้า {e_min:.2f} และออก {e_max:.2f} เต็มเวลาปกติ", 'row': row, 'cell_updates': {},
                    'audit_entry': {"sheet": sheet_name, "emp_id": rec['emp_id'], "input_name": rec['emp_raw'], "sheet_name": s_name, "type": "วันลา", "status": "⚠️ ข้าม: สแกนครบเต็มเวลาปกติ", "scans": scans_str, "detail": f"สแกนนิ้วเต็มเวลาปกติ ({scans_str}) ข้าม '{clean_name}'"}
                })
                continue

        cell_g_val = str(ws[f"{leave_type_col}{row}"].value or '').strip()
        cell_q_val = str(ws[f"{leave_code_col}{row}"].value or '').strip()
        is_unpaid = ("ไม่ได้รับค่าจ้าง" in clean_name) or ("unpaid" in clean_name.lower())

        # ตรวจสอบเงื่อนไขวันหยุดบริษัท:
        # กฎ: ในกรณีที่เป็นวันหยุดบริษัท จะไม่เอาเงื่อนไขการลาใดๆ ไปกรอกลงช่อง G เลย
        # - คนที่ไม่มีสแกน: บังคับกรอก 'น' ในช่อง Q และช่อง G ต้องว่างเปล่า (ล้างค่าวันลาออก)
        # - คนที่มีสแกน: ปล่อยตามสูตรคำนวณของ Excel (ไม่กรอก 'น' และไม่ลงวันลาในช่อง G)
        is_holiday, holiday_name = is_company_holiday(rec['date'], ws)
        if is_holiday:
            if h_in or h_out:
                # มีสแกนทำงาน: ปล่อยตามสูตรคำนวณ ไม่ลงวันลาใดๆ
                cell_updates = {}
                target_cells_list = []
                mod_actions = []

                if cell_g_val and cell_g_val != "None":
                    cell_updates[f"{leave_type_col}{row}"] = None
                    mod_actions.append(f"ล้างข้อความวันลาช่อง {leave_type_col}{row} (วันหยุดบริษัทมีสแกนทำงาน)")
                    target_cells_list.append(f"{leave_type_col}{row}")

                if cell_q_val and not cell_q_val.startswith("=") and cell_q_val in ['พ', 'ป', 'อ', 'น']:
                    cell_updates[f"{leave_code_col}{row}"] = f"=+P{row}-1"
                    mod_actions.append(f"คืนค่าสูตรคำนวณช่อง {leave_code_col}{row} ➔ '=+P{row}-1'")
                    target_cells_list.append(f"{leave_code_col}{row}")

                if not cell_updates:
                    counts['leave_skips'] += 1
                    evaluated_items.append({
                        'sheet': sheet_name, 'date': rec['date'], 'emp_id': rec['emp_id'],
                        'input_name': rec['emp_raw'], 'sheet_name': s_name, 'type': 'วันหยุดบริษัท',
                        'leave_type': '-', 'status': '⚠️ ข้าม: วันหยุดบริษัท (มีสแกน)', 'status_type': 'skip_scan',
                        'scans': scans_str, 'target_cells': '-', 'action': f"ข้าม: พบสแกน ({scans_str}) ในวันหยุดบริษัท ({holiday_name}) ปล่อยคำนวณตามสูตร",
                        'detail': f"วันหยุดบริษัท ({holiday_name}) และมีสแกนเข้า/ออก ({scans_str}) ไม่บันทึกการลา ปล่อยคำนวณตามสูตร",
                        'row': row, 'cell_updates': {},
                        'audit_entry': {"sheet": sheet_name, "emp_id": rec['emp_id'], "input_name": rec['emp_raw'], "sheet_name": s_name, "type": "วันหยุดบริษัท", "status": "⚠️ ข้าม (มีสแกน)", "scans": scans_str, "detail": f"มีสแกนทำงานในวันหยุด ({holiday_name}) ปล่อยตามสูตร"}
                    })
                else:
                    counts['update'] += 1
                    counts['leave_updates'] += 1
                    action_str = " | ".join(mod_actions)
                    evaluated_items.append({
                        'sheet': sheet_name, 'date': rec['date'], 'emp_id': rec['emp_id'],
                        'input_name': rec['emp_raw'], 'sheet_name': s_name, 'type': 'วันหยุดบริษัท',
                        'leave_type': '-', 'status': '✅ จะอัปเดต', 'status_type': 'update',
                        'scans': scans_str, 'target_cells': ", ".join(target_cells_list),
                        'action': f"วันหยุดบริษัท ({holiday_name}): {action_str}",
                        'detail': f"พบสแกน ({scans_str}) ในวันหยุด ➔ ล้างวันลาและคืนค่าสูตรคำนวณ",
                        'row': row, 'cell_updates': cell_updates,
                        'audit_entry': {"sheet": sheet_name, "emp_id": rec['emp_id'], "input_name": rec['emp_raw'], "sheet_name": s_name, "type": "วันหยุดบริษัท", "status": "✅ ปรับคืนสูตร", "scans": scans_str, "detail": action_str}
                    })
                continue
            else:
                # ไม่มีสแกน: บังคับกรอก 'น' ในช่อง Q และช่อง G ต้องว่างเปล่า (ไม่ใส่อะไร)
                cell_updates = {}
                mod_actions = []
                target_cells_list = []

                if cell_q_val != 'น':
                    cell_updates[f"{leave_code_col}{row}"] = 'น'
                    mod_actions.append(f"บันทึกตัวย่อช่อง {leave_code_col}{row} ➔ 'น'")
                    target_cells_list.append(f"{leave_code_col}{row}")

                if cell_g_val and cell_g_val != "None":
                    cell_updates[f"{leave_type_col}{row}"] = None
                    mod_actions.append(f"ล้างข้อความช่อง {leave_type_col}{row} (ไม่ต้องใส่อะไร)")
                    target_cells_list.append(f"{leave_type_col}{row}")

                if not cell_updates:
                    counts['skip_dup'] += 1
                    counts['leave_skips'] += 1
                    evaluated_items.append({
                        'sheet': sheet_name, 'date': rec['date'], 'emp_id': rec['emp_id'],
                        'input_name': rec['emp_raw'], 'sheet_name': s_name, 'type': 'วันหยุดบริษัท',
                        'leave_type': 'วันหยุด (น)', 'status': 'ℹ️ ข้าม: วันหยุดบริษัท (ลง "น" แล้ว)', 'status_type': 'skip_dup',
                        'scans': scans_str, 'target_cells': f"{leave_code_col}{row}",
                        'action': f"ข้าม: วันหยุดบริษัท ({holiday_name}) ช่อง {leave_code_col}{row} มี 'น' และช่อง {leave_type_col}{row} ว่างอยู่แล้ว",
                        'detail': f"วันหยุดบริษัท ({holiday_name}) ลง 'น' ครบถ้วนแล้ว", 'row': row, 'cell_updates': {},
                        'audit_entry': {"sheet": sheet_name, "emp_id": rec['emp_id'], "input_name": rec['emp_raw'], "sheet_name": s_name, "type": "วันหยุดบริษัท", "status": "ℹ️ ข้าม (มี น อยู่แล้ว)", "scans": scans_str, "detail": f"วันหยุด {holiday_name} มี 'น' เรียบร้อย"}
                    })
                else:
                    counts['update'] += 1
                    counts['leave_updates'] += 1
                    action_str = " | ".join(mod_actions)
                    evaluated_items.append({
                        'sheet': sheet_name, 'date': rec['date'], 'emp_id': rec['emp_id'],
                        'input_name': rec['emp_raw'], 'sheet_name': s_name, 'type': 'วันหยุดบริษัท',
                        'leave_type': 'วันหยุด (น)', 'status': '✅ จะอัปเดต', 'status_type': 'update',
                        'scans': scans_str, 'target_cells': ", ".join(target_cells_list),
                        'action': f"วันหยุดบริษัท ({holiday_name}): {action_str}",
                        'detail': f"วันหยุดบริษัท ({holiday_name}) ไม่มีสแกน ➔ บังคับกรอก 'น' และช่อง G ว่างเปล่า",
                        'row': row, 'cell_updates': cell_updates,
                        'audit_entry': {"sheet": sheet_name, "emp_id": rec['emp_id'], "input_name": rec['emp_raw'], "sheet_name": s_name, "type": "วันหยุดบริษัท", "status": "✅ บันทึก 'น'", "scans": scans_str, "detail": action_str}
                    })
                continue

        display_leave = f"{clean_name} {hrs:g} ชม." if (is_partial and hrs is not None) else clean_name

        # วางแผนบันทึกค่าลงในแต่ละช่อง
        mod_actions = []
        cell_updates = {}
        target_cells_list = []

        # 1. จัดการข้อความประเภทการลาในช่อง G (leave_type_col)
        if display_leave not in cell_g_val:
            if cell_g_val and cell_g_val != "None":
                new_g = f"{cell_g_val} / {display_leave}"
                mod_actions.append(f"รวมช่อง {leave_type_col}{row} ➔ '{new_g}'")
                cell_updates[f"{leave_type_col}{row}"] = new_g
            else:
                mod_actions.append(f"บันทึกช่อง {leave_type_col}{row} ➔ '{display_leave}'")
                cell_updates[f"{leave_type_col}{row}"] = display_leave
            target_cells_list.append(f"{leave_type_col}{row}")

        # 2. จัดการตัวอักษรย่อในช่อง Q (leave_code_col เช่น ป, พ, อ)
        # กฎ: สถานะใช้กับ "ลาเต็มวัน" อย่างเดียว (รวมลาโดยไม่ได้รับค่าจ้าง ให้ลงตัวย่อ 'อ' ในช่อง Q ด้วย)
        if not is_partial:
            if leave_code and (leave_code not in cell_q_val or cell_q_val == "None"):
                new_q = f"{cell_q_val}/{leave_code}" if (cell_q_val and cell_q_val != "None" and not cell_q_val.startswith("=")) else leave_code
                mod_actions.append(f"บันทึกตัวย่อช่อง {leave_code_col}{row} ➔ '{new_q}'")
                cell_updates[f"{leave_code_col}{row}"] = new_q
                target_cells_list.append(f"{leave_code_col}{row}")

        # 3. จัดการเวลาเข้า-ออก กรณีลาไม่เต็มวัน (Partial Leave เช่น ครึ่งวัน/รายชั่วโมง)
        if is_partial and hrs is not None:
            # แยกแยะว่าเป็นการลาครึ่งวันบ่าย หรือ ลาครึ่งวันเช้า
            # - ลาครึ่งวันบ่าย: เช้ามาทำงาน แต่ช่วงบ่ายลา (สแกนครั้งสุดท้ายก่อน 13:00)
            # - ลาครึ่งวันเช้า: เช้าลา แต่ช่วงบ่ายมาทำงาน (มีสแกนหลัง 13:00)
            is_afternoon_leave = False
            if e_max is not None:
                is_afternoon_leave = (e_max < 13.0)
            else:
                start_raw_str = str(rec.get('start_time_raw', ''))
                m_t = re.search(r'(\d{1,2}):(\d{2})', start_raw_str)
                is_afternoon_leave = (int(m_t.group(1)) >= 12) if m_t else True

            if is_afternoon_leave:
                # ถ้าลาครึ่งวันบ่าย: ไม่ใส่สถานะช่อง Q และไม่แตะต้องช่อง L (ปล่อยตามเวลาสแกนเช้าจริง/สูตรเดิม) แต่จะปรับช่อง M ให้เลิก 16:30
                cell_m_cur = ws[f"{checkout_col}{row}"].value
                is_m_1630 = (isinstance(cell_m_cur, time) and cell_m_cur.hour == 16 and cell_m_cur.minute == 30) or (isinstance(cell_m_cur, str) and re.match(r'^16:30', cell_m_cur.strip()))
                if not is_m_1630:
                    cell_updates[f"{checkout_col}{row}"] = time(16, 30)
                    mod_actions.append(f"ปรับเวลาออก (ลาบ่าย) {checkout_col}{row} ➔ 16:30")
                    target_cells_list.append(f"{checkout_col}{row}")
            else:
                # ถ้าลาเช้า บ่ายมาทำ: แก้เวลาเข้า เป็น 7:30 ไม่แตะต้องขาออก (ให้ใช้เวลาตามจริงในชีท) และไม่ลงสถานะช่อง Q
                cell_l_cur = ws[f"{checkin_col}{row}"].value
                is_l_set = (isinstance(cell_l_cur, time) and cell_l_cur.hour == 7 and cell_l_cur.minute == 30) or (isinstance(cell_l_cur, str) and re.match(r'^0?7:30', cell_l_cur.strip()))
                if not is_l_set:
                    cell_updates[f"{checkin_col}{row}"] = time(7, 30)
                    mod_actions.append(f"ปรับเวลาเข้า (ลาเช้า) {checkin_col}{row} ➔ 7:30")
                    target_cells_list.append(f"{checkin_col}{row}")

        # หากมีข้อมูลเป้าหมายครบถ้วนอยู่แล้วทุกช่อง ให้ถือว่าข้าม
        if not cell_updates:
            counts['skip_dup'] += 1
            counts['leave_skips'] += 1
            evaluated_items.append({
                'sheet': sheet_name, 'date': rec['date'], 'emp_id': rec['emp_id'],
                'input_name': rec['emp_raw'], 'sheet_name': s_name, 'type': 'วันลา',
                'leave_type': clean_name, 'status': 'ℹ️ ข้าม: มีข้อมูลเดิมอยู่แล้ว', 'status_type': 'skip_dup',
                'scans': scans_str, 'target_cells': f"{leave_type_col}{row}", 'action': f"ข้าม: ช่อง {leave_type_col}{row} มี '{display_leave}' และเวลาตรงอยู่แล้ว",
                'detail': f"ข้อมูลเดิมในช่อง: {cell_g_val}", 'row': row, 'cell_updates': {},
                'audit_entry': {"sheet": sheet_name, "emp_id": rec['emp_id'], "input_name": rec['emp_raw'], "sheet_name": s_name, "type": "วันลา", "status": "ℹ️ ข้าม: ข้อมูลเดิมซ้ำ", "scans": scans_str, "detail": f"มีข้อมูล '{display_leave}' ครบถ้วนแล้ว"}
            })
            continue

        counts['update'] += 1
        counts['leave_updates'] += 1
        action_str = " | ".join(mod_actions)
        evaluated_items.append({
            'sheet': sheet_name, 'date': rec['date'], 'emp_id': rec['emp_id'],
            'input_name': rec['emp_raw'], 'sheet_name': s_name, 'type': 'วันลา',
            'leave_type': clean_name, 'status': '✅ จะอัปเดต', 'status_type': 'update',
            'scans': scans_str, 'target_cells': ", ".join(target_cells_list), 'action': action_str,
            'detail': f"{'ลาไม่เต็มวัน' if is_partial else 'ลาเต็มวัน'} ({clean_name})" + (f" [{m_msg}]" if m_msg else ""),
            'row': row, 'cell_updates': cell_updates,
            'audit_entry': {"sheet": sheet_name, "emp_id": rec['emp_id'], "input_name": rec['emp_raw'], "sheet_name": s_name, "type": "วันลา", "status": "✅ อัปเดตสำเร็จ", "scans": scans_str, "detail": action_str}
        })

    # 2. ประเมินใบแจ้งการมาทำงาน
    for rec in attendance_records:
        sheet_name = sheet_date_map.get(rec['date'])
        if not sheet_name:
            continue
        ws = wb[sheet_name]
        attend_type = 'เช้า (07:30)' if rec['target_col'] == 'L' else 'เย็น (ปัดเศษ 30 นาที)'

        row, s_name, m_type, m_msg = find_emp_row(sheet_name, rec['emp_id'], rec['emp_raw'])

        if m_type == 'not_found':
            counts['not_found'] += 1
            counts['attend_skips'] += 1
            evaluated_items.append({
                'sheet': sheet_name, 'date': rec['date'], 'emp_id': rec['emp_id'],
                'input_name': rec['emp_raw'], 'sheet_name': '-', 'type': 'แจ้งเวลา',
                'attend_type': attend_type, 'status': '⚠️ ไม่พบพนักงานในชีท', 'status_type': 'not_found',
                'scans': '-', 'target_cells': '-', 'action': 'ข้าม: ไม่พบรหัสและชื่อพนักงานในชีทนี้',
                'detail': m_msg, 'row': None, 'cell_updates': {},
                'audit_entry': {"sheet": sheet_name, "emp_id": rec['emp_id'], "input_name": rec['emp_raw'], "sheet_name": "-", "type": "แจ้งเวลา", "status": "⚠️ ไม่พบพนักงานในชีท", "scans": "-", "detail": m_msg}
            })
            continue


        if (rec['emp_id'], rec['date']) in full_day_leaves:
            l_rec = full_day_leaves[(rec['emp_id'], rec['date'])]
            l_cln = clean_leave_text(l_rec['leave_type_raw'])
            counts['conflict'] += 1
            counts['attend_skips'] += 1
            evaluated_items.append({
                'sheet': sheet_name, 'date': rec['date'], 'emp_id': rec['emp_id'],
                'input_name': rec['emp_raw'], 'sheet_name': s_name, 'type': 'แจ้งเวลา',
                'attend_type': attend_type, 'status': '⚡ ข้อมูลขัดแย้ง: ลาชนแจ้งเวลา', 'status_type': 'conflict',
                'scans': '-', 'target_cells': '-', 'action': f"ข้าม: พบวันลาเต็มวัน '{l_cln}' ในวันเดียวกัน ยึดวันลาเป็นหลัก",
                'detail': f"ชนกับวันลาเต็มวัน '{l_cln}'", 'row': row, 'cell_updates': {},
                'audit_entry': {"sheet": sheet_name, "emp_id": rec['emp_id'], "input_name": rec['emp_raw'], "sheet_name": s_name, "type": "แจ้งเวลา", "status": "⚡ ข้อมูลขัดแย้ง: ลาชนแจ้งเวลา", "scans": "-", "detail": f"พบวันลาเต็มวัน '{l_cln}' ชนกับใบแจ้งเวลาในวันเดียวกัน ยึดวันลาเป็นหลัก"}
            })
            continue

        target_col = rec['target_col']
        target_time_obj = rec['target_time']
        actual_target_col = checkin_col if target_col == 'L' else checkout_col

        scan_val = ws[f"{scan_col}{row}"].value
        h_in, h_out, scans_list, e_min, e_max = check_fingerprint_scans(scan_val)
        scans_str = ", ".join(scans_list) if scans_list else "-"

        if h_in and h_out:
            counts['skip_scan'] += 1
            counts['attend_skips'] += 1
            evaluated_items.append({
                'sheet': sheet_name, 'date': rec['date'], 'emp_id': rec['emp_id'],
                'input_name': rec['emp_raw'], 'sheet_name': s_name, 'type': 'แจ้งเวลา',
                'attend_type': attend_type, 'status': '⚠️ ข้าม: สแกนนิ้วครบแล้ว', 'status_type': 'skip_scan',
                'scans': scans_str, 'target_cells': f"{actual_target_col}{row}", 'action': f"ข้าม: มีสแกนนิ้วครบทั้งเข้า-ออก ({scans_str}) อาจกรอกแจ้งผิดวัน",
                'detail': "มีสแกนครบทั้งเข้าและออก", 'row': row, 'cell_updates': {},
                'audit_entry': {"sheet": sheet_name, "emp_id": rec['emp_id'], "input_name": rec['emp_raw'], "sheet_name": s_name, "type": "แจ้งเวลา", "status": "⚠️ ข้าม: สแกนนิ้วครบแล้ว", "scans": scans_str, "detail": f"สแกนนิ้วครบทั้งเข้า-ออกแล้ว ({scans_str}) อาจกรอกแจ้งผิดวัน"}
            })
            continue

        if target_col == 'L' and h_in:
            counts['skip_scan'] += 1
            counts['attend_skips'] += 1
            evaluated_items.append({
                'sheet': sheet_name, 'date': rec['date'], 'emp_id': rec['emp_id'],
                'input_name': rec['emp_raw'], 'sheet_name': s_name, 'type': 'แจ้งเวลา',
                'attend_type': attend_type, 'status': 'ℹ️ ข้าม: มีสแกนเข้าอยู่แล้ว', 'status_type': 'skip_scan',
                'scans': scans_str, 'target_cells': f"{actual_target_col}{row}", 'action': f"ข้าม: มีสแกนเข้างานช่วงเช้าแล้ว ({scans_str}) ไม่ต้องใช้ใบแจ้งฯ",
                'detail': "พบสแกนเข้างานเช้า", 'row': row, 'cell_updates': {},
                'audit_entry': {"sheet": sheet_name, "emp_id": rec['emp_id'], "input_name": rec['emp_raw'], "sheet_name": s_name, "type": "แจ้งเวลา", "status": "ℹ️ ข้าม: มีสแกนเข้าอยู่แล้ว", "scans": scans_str, "detail": f"พบสแกนเข้างาน {scans_str} อยู่แล้ว ไม่จำเป็นต้องใช้ใบแจ้งเวลา"}
            })
            continue

        if target_col == 'M' and h_out and e_max is not None and e_max >= 16.5:
            counts['skip_scan'] += 1
            counts['attend_skips'] += 1
            evaluated_items.append({
                'sheet': sheet_name, 'date': rec['date'], 'emp_id': rec['emp_id'],
                'input_name': rec['emp_raw'], 'sheet_name': s_name, 'type': 'แจ้งเวลา',
                'attend_type': attend_type, 'status': 'ℹ️ ข้าม: มีสแกนออกอยู่แล้ว', 'status_type': 'skip_scan',
                'scans': scans_str, 'target_cells': f"{actual_target_col}{row}", 'action': f"ข้าม: มีสแกนออกงานแล้ว ({scans_str}) ไม่ต้องใช้ใบแจ้งฯ",
                'detail': "พบสแกนออกงานเย็น >= 16:30", 'row': row, 'cell_updates': {},
                'audit_entry': {"sheet": sheet_name, "emp_id": rec['emp_id'], "input_name": rec['emp_raw'], "sheet_name": s_name, "type": "แจ้งเวลา", "status": "ℹ️ ข้าม: มีสแกนออกอยู่แล้ว", "scans": scans_str, "detail": f"พบสแกนออกงาน {scans_str} อยู่แล้ว ไม่จำเป็นต้องใช้ใบแจ้งเวลา"}
            })
            continue

        time_cell_val = ws[f"{actual_target_col}{row}"].value
        curr_dec = to_decimal_hours(time_cell_val)
        targ_dec = target_time_obj.hour + target_time_obj.minute / 60.0

        if curr_dec is not None and abs(curr_dec - targ_dec) < 0.01:
            counts['skip_dup'] += 1
            counts['attend_skips'] += 1
            evaluated_items.append({
                'sheet': sheet_name, 'date': rec['date'], 'emp_id': rec['emp_id'],
                'input_name': rec['emp_raw'], 'sheet_name': s_name, 'type': 'แจ้งเวลา',
                'attend_type': attend_type, 'status': 'ℹ️ ข้าม: เวลาตรงกันอยู่แล้ว', 'status_type': 'skip_dup',
                'scans': scans_str, 'target_cells': f"{actual_target_col}{row}", 'action': f"ข้าม: ช่อง {actual_target_col}{row} เป็นเวลา {target_time_obj.hour}:{target_time_obj.minute:02d} อยู่แล้ว",
                'detail': "เวลาเดิมตรงกับเป้าหมาย", 'row': row, 'cell_updates': {},
                'audit_entry': {"sheet": sheet_name, "emp_id": rec['emp_id'], "input_name": rec['emp_raw'], "sheet_name": s_name, "type": "แจ้งเวลา", "status": "ℹ️ ข้าม: เวลาตรงกันอยู่แล้ว", "scans": scans_str, "detail": f"เวลาในช่อง {actual_target_col} ตรงกับ {target_time_obj.hour}:{target_time_obj.minute:02d} อยู่แล้ว"}
            })
            continue

        time_label = f"{target_time_obj.hour}:{target_time_obj.minute:02d}"
        mod_actions = [f"บันทึกเวลาช่อง {actual_target_col}{row} ➔ {time_label} (แจ้งเดิม: {rec['orig_time']})"]
        cell_updates = {f"{actual_target_col}{row}": target_time_obj}

        cell_g_val = str(ws[f"{leave_type_col}{row}"].value or '').strip()
        if "มีใบแจ้งฯ" not in cell_g_val:
            new_g = f"{cell_g_val} / มีใบแจ้งฯ" if (cell_g_val and cell_g_val != "None") else "มีใบแจ้งฯ"
            mod_actions.append(f"เติมช่อง {leave_type_col}{row} ➔ '{new_g}'")
            cell_updates[f"{leave_type_col}{row}"] = new_g

        counts['update'] += 1
        counts['attend_updates'] += 1
        action_str = " | ".join(mod_actions)
        evaluated_items.append({
            'sheet': sheet_name, 'date': rec['date'], 'emp_id': rec['emp_id'],
            'input_name': rec['emp_raw'], 'sheet_name': s_name, 'type': 'แจ้งเวลา',
            'attend_type': attend_type, 'status': '✅ จะอัปเดต', 'status_type': 'update',
            'scans': scans_str, 'target_cells': f"{actual_target_col}{row}, {leave_type_col}{row}",
            'action': action_str, 'detail': f"แจ้งเวลา {rec['orig_time']} ➔ {time_label}" + (f" [{m_msg}]" if m_msg else ""),
            'row': row, 'cell_updates': cell_updates,
            'audit_entry': {"sheet": sheet_name, "emp_id": rec['emp_id'], "input_name": rec['emp_raw'], "sheet_name": s_name, "type": "แจ้งเวลา", "status": "✅ อัปเดตสำเร็จ", "scans": scans_str, "detail": action_str}
        })

    target_dates = set(sheet_date_map.keys())
    d_min = min(target_dates) if target_dates else date.today()
    d_max = max(target_dates) if target_dates else date.today()

    summary = {
        'total_items': len(evaluated_items),
        'updates_count': counts['update'],
        'skips_count': len(evaluated_items) - counts['update'],
        'leave_updates': counts['leave_updates'],
        'leave_skips': counts['leave_skips'],
        'attend_updates': counts['attend_updates'],
        'attend_skips': counts['attend_skips'],
        'skip_scan_count': counts['skip_scan'],
        'skip_dup_count': counts['skip_dup'],
        'conflict_count': counts['conflict'],
        'mismatch_count': counts['mismatch'],
        'not_found_count': counts['not_found'],
        'sheets_count': len(sheet_date_map),
        'total_employees': len(all_employees),
        'date_range': f"{d_min.strftime('%d/%m/%Y')} ถึง {d_max.strftime('%d/%m/%Y')}"
    }

    return summary, evaluated_items


# -------------------------------------------------------------
# ระบบจำลองการแก้ไขล่วงหน้า & สถิติโครงการ (Preview & Analytics)
# -------------------------------------------------------------

def simulate_all_changes(main_filepath, approval_filepath=None, attendance_filepath=None, col_config=None, log_callback=None, progress_callback=None):
    """
    จำลองการประมวลผลจริงทั้งหมดโดยไม่แก้ไขไฟล์หลัก (Dry-run Simulation)
    คืนค่า (summary, preview_items) ผ่าน evaluate_all_records ร่วมกับระบบประมวลผลจริง
    """
    log_callback = log_callback or (lambda msg: None)
    progress_callback = progress_callback or (lambda p: None)

    cfg = col_config or get_default_column_config()
    progress_callback(10)
    log_callback("เริ่มต้นจำลองการประมวลผล (Pre-Check Preview)...")

    wb = openpyxl.load_workbook(main_filepath, data_only=True)
    sheet_date_map = map_workbook_sheets(wb)
    if not sheet_date_map:
        wb.close()
        raise ValueError("ไม่พบชีทวันที่ในไฟล์ STAFF TIME")

    progress_callback(35)
    log_callback(f"พบชีทรายวัน {len(sheet_date_map)} ชีทในไฟล์ STAFF TIME")

    target_dates = set(sheet_date_map.keys())
    leave_records = []
    if approval_filepath and os.path.exists(approval_filepath):
        df_appr = read_excel_safely(approval_filepath)
        leave_records = parse_approval_records(df_appr, target_dates, log_callback)

    attendance_records = []
    if attendance_filepath and os.path.exists(attendance_filepath):
        df_att = read_excel_safely(attendance_filepath)
        attendance_records = parse_attendance_records(df_att, target_dates, log_callback)

    progress_callback(65)
    log_callback("⚙️ กำลังประเมินเงื่อนไข Safeguard และจับคู่ชื่อ-รหัสพนักงาน...")
    summary, items = evaluate_all_records(wb, sheet_date_map, leave_records, attendance_records, cfg)
    wb.close()

    progress_callback(95)
    log_callback(f"✅ จำลองเสร็จสิ้น: พบทั้งหมด {len(items)} รายการ (จะอัปเดต {summary['updates_count']} / ข้าม {summary['skips_count']})")
    return summary, items


def get_project_analytics(main_filepath, approval_filepath=None, attendance_filepath=None, col_config=None):
    """
    สรุปสถิติภาพรวมประจำงวดของโครงการ ทั้งวันลา การแจ้งเวลา และความถูกต้อง
    อ่าน leave_type และ attend_type ตรงจากผลการจำลอง ปราศจากบั๊กตัดข้อความ
    """
    summary, items = simulate_all_changes(main_filepath, approval_filepath, attendance_filepath, col_config)
    
    leave_types_count = {}
    attend_types_count = {'เช้า (07:30)': 0, 'เย็น (ปัดเศษ 30 นาที)': 0}
    
    for item in items:
        if item['type'] == 'วันลา':
            l_type = item.get('leave_type') or 'ไม่ระบุประเภท'
            leave_types_count[l_type] = leave_types_count.get(l_type, 0) + 1
        elif item['type'] == 'แจ้งเวลา':
            a_type = item.get('attend_type') or 'เช้า (07:30)'
            attend_types_count[a_type] = attend_types_count.get(a_type, 0) + 1
            
    return {
        'summary': summary,
        'leave_breakdown': leave_types_count,
        'attend_breakdown': attend_types_count,
        'total_items': len(items)
    }


def get_employee_details(emp_query, main_filepath, approval_filepath=None, attendance_filepath=None, col_config=None):
    """
    ค้นหาข้อมูลรายบุคคลสำหรับหน้าจอ Employee Quick Inspector:
    ดึงข้อมูลประวัติในชีท STAFF TIME ทุกวัน, รายการลาที่ได้รับอนุมัติ, และใบแจ้งเวลา
    """
    cfg = col_config or get_default_column_config()
    emp_id_col = cfg.get('emp_id_col', 'A').upper().strip() or 'A'
    emp_name_col = cfg.get('emp_name_col', 'B').upper().strip() or 'B'
    scan_col = cfg.get('scan_col', 'E').upper().strip() or 'E'
    leave_type_col = cfg.get('leave_type_col', 'G').upper().strip() or 'G'
    checkin_col = cfg.get('checkin_col', 'L').upper().strip() or 'L'
    checkout_col = cfg.get('checkout_col', 'M').upper().strip() or 'M'
    leave_code_col = cfg.get('leave_code_col', 'Q').upper().strip() or 'Q'
    try:
        start_row = int(cfg.get('start_row', 2))
    except Exception:
        start_row = 2

    query_str = str(emp_query).strip()
    query_id = extract_emp_id(query_str)
    query_name_clean = clean_thai_name(query_str)

    wb = openpyxl.load_workbook(main_filepath, data_only=True)
    sheet_date_map = map_workbook_sheets(wb)
    target_dates = set(sheet_date_map.keys())

    matched_id = None
    matched_name = None

    sorted_dates = sorted(sheet_date_map.keys())
    for dt in sorted_dates:
        sheet_name = sheet_date_map[dt]
        ws = wb[sheet_name]
        for r in range(start_row, ws.max_row + 1):
            val_id = ws[f"{emp_id_col}{r}"].value
            val_name = ws[f"{emp_name_col}{r}"].value
            eid = extract_emp_id(val_id) if val_id is not None else ""
            raw_n = str(val_name).strip() if val_name is not None and str(val_name) != 'nan' else ""
            
            is_match = False
            if query_id and eid and query_id == eid:
                is_match = True
            elif query_str and raw_n and (query_str.lower() in raw_n.lower() or (query_name_clean and query_name_clean.lower() in raw_n.lower())):
                is_match = True

            if is_match:
                matched_id = eid
                matched_name = raw_n
                break
        if matched_id:
            break

    if not matched_id and not matched_name:
        wb.close()
        return None

    daily_records = []
    scans_total_days = 0
    leave_total_days = 0

    for dt in sorted_dates:
        sheet_name = sheet_date_map[dt]
        ws = wb[sheet_name]
        row_found = None
        
        for r in range(start_row, ws.max_row + 1):
            val_id = ws[f"{emp_id_col}{r}"].value
            val_name = ws[f"{emp_name_col}{r}"].value
            eid = extract_emp_id(val_id) if val_id is not None else ""
            raw_n = str(val_name).strip() if val_name is not None and str(val_name) != 'nan' else ""
            if eid == matched_id:
                row_found = r
                break

        if row_found:
            scan_val = ws[f"{scan_col}{row_found}"].value
            h_in, h_out, scans_list, _, _ = check_fingerprint_scans(scan_val)
            scans_str = ", ".join(scans_list) if scans_list else "-"
            if h_in or h_out:
                scans_total_days += 1

            leave_val = str(ws[f"{leave_type_col}{row_found}"].value or '').strip()
            if leave_val and leave_val != "None":
                leave_total_days += 1
            code_val = str(ws[f"{leave_code_col}{row_found}"].value or '').strip()
            in_val = ws[f"{checkin_col}{row_found}"].value
            out_val = ws[f"{checkout_col}{row_found}"].value

            daily_records.append({
                'date': dt,
                'sheet': sheet_name,
                'row': row_found,
                'scans': scans_str,
                'has_in': h_in,
                'has_out': h_out,
                'leave': leave_val if leave_val != "None" else "-",
                'code': code_val if code_val != "None" else "-",
                'in_time': str(in_val) if in_val is not None else "-",
                'out_time': str(out_val) if out_val is not None else "-"
            })
        else:
            daily_records.append({
                'date': dt,
                'sheet': sheet_name,
                'row': None,
                'scans': "-",
                'has_in': False,
                'has_out': False,
                'leave': "-",
                'code': "-",
                'in_time': "-",
                'out_time': "-"
            })

    wb.close()

    appr_leaves = []
    if approval_filepath and os.path.exists(approval_filepath):
        df_appr = read_excel_safely(approval_filepath)
        all_l = parse_approval_records(df_appr, target_dates, lambda msg: None)
        for l in all_l:
            if l['emp_id'] == matched_id:
                appr_leaves.append(l)

    att_records = []
    if attendance_filepath and os.path.exists(attendance_filepath):
        df_att = read_excel_safely(attendance_filepath)
        all_att = parse_attendance_records(df_att, target_dates, lambda msg: None)
        for a in all_att:
            if a['emp_id'] == matched_id:
                att_records.append(a)

    return {
        'emp_id': matched_id,
        'emp_name': matched_name,
        'total_days_in_period': len(sorted_dates),
        'days_with_scans': scans_total_days,
        'days_with_leaves': leave_total_days,
        'approved_leaves_count': len(appr_leaves),
        'attendance_notifications_count': len(att_records),
        'daily_records': daily_records,
        'approved_leaves': appr_leaves,
        'attendance_records': att_records
    }


# -------------------------------------------------------------
# ฟังก์ชันประมวลผลและบันทึกไฟล์หลักจริง (Main Execution Engine)
# เรียกใช้ evaluate_all_records ร่วมกับ Preview เพื่อความสม่ำเสมอ 100%
# -------------------------------------------------------------

def process_all_files(main_filepath, approval_filepath=None, attendance_filepath=None, 
                      save_mode='new_file', col_config=None, log_callback=None, progress_callback=None):
    """
    ประมวลผลไฟล์ STAFF TIME ด้วยการลาที่ได้รับการอนุมัติ และ/หรือ ใบแจ้งการมาทำงาน
    ใช้ evaluate_all_records รับประกันว่ารายการที่บันทึกลง Excel ตรงกับรายการที่ Preview 100%
    """
    log_callback = log_callback or print
    progress_callback = progress_callback or (lambda p: None)
    cfg = col_config or get_default_column_config()

    progress_callback(5)
    log_callback("=" * 72)
    log_callback("ระบบบันทึกเวลาทำงานและประมวลผลการลาอัตโนมัติ (CNT STAFF TIME AUTOMATION)")
    log_callback("=" * 72)
    log_callback(f"ไฟล์รายงาน STAFF TIME: {os.path.basename(main_filepath)}")
    if approval_filepath:
        log_callback(f"ไฟล์การลาที่ได้รับอนุมัติ: {os.path.basename(approval_filepath)}")
    if attendance_filepath:
        log_callback(f"ไฟล์ใบแจ้งการมาทำงาน: {os.path.basename(attendance_filepath)}")

    wb = openpyxl.load_workbook(main_filepath)
    sheet_date_map = map_workbook_sheets(wb)
    if not sheet_date_map:
        wb.close()
        raise ValueError("ไม่พบชีทวันที่ในไฟล์ STAFF TIME กรุณาตรวจสอบว่าเป็นไฟล์รายงานเวลาที่ถูกต้อง")

    target_dates = set(sheet_date_map.keys())
    d_min = min(target_dates); d_max = max(target_dates)
    log_callback(f"พบชีทรายวันทั้งหมด {len(sheet_date_map)} ชีท (ช่วง {d_min.strftime('%d/%m/%Y')} ถึง {d_max.strftime('%d/%m/%Y')})")

    leave_records = []
    if approval_filepath and os.path.exists(approval_filepath):
        df_appr = read_excel_safely(approval_filepath)
        leave_records = parse_approval_records(df_appr, target_dates, log_callback)
        log_callback(f"✅ คัดกรองข้อมูลวันลาที่ตรงกับรอบไฟล์ STAFF TIME ได้ {len(leave_records)} รายการ")

    attendance_records = []
    if attendance_filepath and os.path.exists(attendance_filepath):
        df_att = read_excel_safely(attendance_filepath)
        attendance_records = parse_attendance_records(df_att, target_dates, log_callback)
        log_callback(f"✅ คัดกรองข้อมูลแจ้งเวลาที่ตรงกับรอบไฟล์ STAFF TIME ได้ {len(attendance_records)} รายการ")

    if not leave_records and not attendance_records:
        wb.close()
        log_callback("\n⚠️ ไม่พบข้อมูลที่ตรงกับรอบวันที่ของไฟล์ STAFF TIME เลย กรุณาตรวจสอบวันที่ในไฟล์นำเข้า")
        return None, None, 0, 0

    progress_callback(30)
    summary, evaluated_items = evaluate_all_records(wb, sheet_date_map, leave_records, attendance_records, cfg)

    audit_rows = []
    log_callback("\n⚙️ กำลังบันทึกข้อมูลลงในแต่ละชีท (พร้อม Double-Check ชื่อ-รหัส)...")

    for item in evaluated_items:
        ws = wb[item['sheet']] if item['sheet'] in wb.sheetnames else None
        if item['status_type'] == 'update' and ws:
            for coord, val in item.get('cell_updates', {}).items():
                cell = ws[coord]
                if isinstance(val, time):
                    set_cell_time(cell, val)
                else:
                    cell.value = val

        if item.get('action'):
            if item['status_type'] == 'update':
                log_callback(f"  ชีท {item['sheet']} | รหัส {item['emp_id']} ({item['sheet_name'] or item['input_name']}) -> {item['action']}")
            elif item['status_type'] in ['mismatch', 'conflict', 'skip_scan', 'not_found']:
                log_callback(f"  ⚠️ [{item['status']}] ชีท {item['sheet']} | รหัส {item['emp_id']} ({item['input_name']}): {item['action']}")

        if item.get('audit_entry'):
            audit_rows.append(item['audit_entry'])

    progress_callback(80)

    # จัดการฟอร์แมตเวลาให้ตรงกับต้นฉบับ 100% และตรวจสอบวันหยุดบริษัท (Clean-up Sweep)
    checkin_col = cfg.get('checkin_col', 'L').upper().strip() or 'L'
    checkout_col = cfg.get('checkout_col', 'M').upper().strip() or 'M'
    leave_type_col = cfg.get('leave_type_col', 'G').upper().strip() or 'G'
    leave_code_col = cfg.get('leave_code_col', 'Q').upper().strip() or 'Q'
    scan_col = cfg.get('scan_col', 'E').upper().strip() or 'E'
    start_row = int(cfg.get('start_row', 2))
    for dt, sheet_name in sheet_date_map.items():
        ws = wb[sheet_name]
        is_holiday, holiday_name = is_company_holiday(dt, ws)
        for r in range(start_row, ws.max_row + 1):
            for col_letter in [checkin_col, checkout_col]:
                c = ws[f"{col_letter}{r}"]
                if isinstance(c.value, str):
                    m = re.match(r'^(\d{1,2}):(\d{2})(?::\d{2})?$', c.value.strip())
                    if m:
                        set_cell_time(c, time(int(m.group(1)), int(m.group(2))))
                elif isinstance(c.value, (datetime, time)):
                    c.number_format = 'h:mm'
            # คืนค่าสูตรคำนวณในช่อง Q หากเคยมีตัวย่อติดไปกับสูตร
            q_cell = ws[f"{leave_code_col}{r}"]
            q_val = str(q_cell.value or '').strip()
            if q_val.startswith('=+P') and '/' in q_val:
                q_cell.value = re.sub(r'/[ก-๙A-Za-z]+$', '', q_val)

            # ตรวจสอบวันหยุดบริษัทสำหรับทุกแถว:
            # - คนไม่มีสแกน บังคับ Q เป็น 'น' และช่อง G ต้องว่าง
            # - คนมีสแกน ปล่อยตามสูตร (ห้ามเป็น 'น') และช่อง G ว่าง
            if is_holiday and ws[f"B{r}"].value:
                scan_val = ws[f"{scan_col}{r}"].value
                h_in, h_out, _, _, _ = check_fingerprint_scans(scan_val)
                g_cell = ws[f"{leave_type_col}{r}"]
                g_val = str(g_cell.value or '').strip()
                if not h_in and not h_out:
                    if q_val != 'น':
                        q_cell.value = 'น'
                    if g_val and g_val != 'None':
                        g_cell.value = None
                else:
                    if q_val == 'น':
                        q_cell.value = f"=+P{r}-1"
                    if g_val and g_val != 'None':
                        g_cell.value = None

    # อัปเดตข้อมูลพนักงานที่ไม่มีการแจ้งออนไลน์ลงในชีท 'ไม่มีแจ้งในระบบ' อัตโนมัติ (หากมีชีทนี้อยู่ในไฟล์)
    if 'ไม่มีแจ้งในระบบ' in wb.sheetnames:
        try:
            populate_unrecorded_sheet_in_workbook(wb, filename=os.path.basename(main_filepath), col_config=cfg, create_sheet=False)
            log_callback("📝 อัปเดตข้อมูลชีท 'ไม่มีแจ้งในระบบ' (จัดกลุ่มตามชื่อพนักงาน) เรียบร้อย")
        except Exception as e:
            log_callback(f"⚠️ ไม่สามารถอัปเดตชีท 'ไม่มีแจ้งในระบบ' ได้: {e}")

    # บันทึกไฟล์ผลลัพธ์
    dir_name = os.path.dirname(main_filepath)
    base_name = os.path.basename(main_filepath)
    name_only, ext_only = os.path.splitext(base_name)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if save_mode == 'overwrite':
        backup_path = os.path.join(dir_name, f"{name_only}_backup_{timestamp}{ext_only}")
        shutil.copy2(main_filepath, backup_path)
        log_callback(f"\nสร้างไฟล์สำรองเรียบร้อย: {backup_path}")
        output_filepath = main_filepath
    else:
        output_filepath = os.path.join(dir_name, f"{name_only}_updated_{timestamp}{ext_only}")

    log_callback(f"กำลังบันทึกไฟล์ไปที่: {output_filepath} ...")
    wb.save(output_filepath)
    wb.close()

    # สร้างรายงาน Audit Report
    report_filename = f"Audit_Report_{timestamp}.xlsx"
    report_filepath = os.path.join(dir_name, report_filename)
    try:
        create_audit_report(report_filepath, audit_rows)
        log_callback(f"📊 สร้างไฟล์รายงานตรวจสอบ (Audit Report) เรียบร้อย: {report_filename}")
    except Exception as e:
        log_callback(f"⚠️ ไม่สามารถสร้างไฟล์รายงานตรวจสอบได้: {e}")
        report_filepath = None

    progress_callback(100)
    total_up = summary['updates_count']
    total_sk = summary['skips_count']
    log_callback("=" * 68)
    log_callback("ประมวลผลเสร็จสิ้นสมบูรณ์แบบ 100%")
    log_callback(f"  • การลาที่ได้รับการอนุมัติ: อัปเดต {summary['leave_updates']} จุด, ข้าม {summary['leave_skips']} จุด")
    log_callback(f"  • ใบแจ้งการมาทำงาน: อัปเดต {summary['attend_updates']} จุด, ข้าม {summary['attend_skips']} จุด")
    log_callback(f"บันทึกไฟล์ที่: {output_filepath}")
    if report_filepath:
        log_callback(f"ไฟล์รายงานตรวจสอบ: {report_filepath}")
    log_callback("=" * 68)

    return output_filepath, report_filepath, total_up, total_sk


# -------------------------------------------------------------
# หน้าต่างโมดอล: คู่มือการใช้งานระบบ (User Manual Dialog)
# ออกแบบให้อ่านง่าย ชัดเจน ตัวหนังสือใหญ่ สบายตาสำหรับผู้ใหญ่ (Senior-Friendly)
# -------------------------------------------------------------

class ManualDialog(ctk.CTkToplevel):
    def __init__(self, parent, icon_path=None):
        super().__init__(parent)
        self.title("คู่มือการใช้งานระบบ CNT STAFF TIME AUTOMATION - BY WERAWAT THAMRONGDULLAPARK")
        center_window(self, 1080, 800, parent)
        self.minsize(960, 680)
        self.font_family = "Tahoma"

        if icon_path and os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        self.setup_ui()
        self.lift()
        self.attributes("-topmost", True)
        self.after_idle(lambda: self.attributes("-topmost", False))
        self.grab_set()
        self.focus_force()

    def setup_ui(self):
        header = ctk.CTkFrame(self, fg_color="#0f172a", corner_radius=0, height=88)
        header.pack(fill="x")
        
        lbl_t = ctk.CTkLabel(
            header,
            text="📖 คู่มือการใช้งานระบบ CNT STAFF TIME AUTOMATION (Universal Edition)",
            font=ctk.CTkFont(family=self.font_family, size=21, weight="bold"),
            text_color="#ffffff"
        )
        lbl_t.pack(pady=(12, 3))

        lbl_sub = ctk.CTkLabel(
            header,
            text="ขั้นตอนการทำงาน • กฎคำนวณเวลาเข้า-ออก • ระบบกันพลาด 5 ชั้น • การตั้งค่าตำแหน่งคอลัมน์\nพัฒนาโดย: คุณวีรวัช ธำรงดุลภาค (WERAWAT THAMRONGDULLAPARK) | บมจ.คริสเตียนีและนีลเส็น (ไทย)",
            font=ctk.CTkFont(family=self.font_family, size=13),
            text_color="#93c5fd"
        )
        lbl_sub.pack(pady=(0, 12))

        tabview = ctk.CTkTabview(self, corner_radius=12)
        tabview.pack(fill="both", expand=True, padx=18, pady=(10, 8))

        try:
            tabview._segmented_button.configure(
                font=ctk.CTkFont(family=self.font_family, size=14, weight="bold"),
                height=42
            )
        except Exception:
            pass

        tabs = [
            ("ขั้นตอนใช้งาน 3 ขั้นตอน", self._build_tab_steps),
            ("กฎเวลาและรหัสวันลา", self._build_tab_rules),
            ("ระบบกันพลาด 5 ชั้น", self._build_tab_safeguards),
            ("ตำแหน่งคอลัมน์มาตรฐาน", self._build_tab_columns),
            ("คำถามที่พบบ่อย (FAQ)", self._build_tab_faq)
        ]

        for tab_name, builder in tabs:
            tabview.add(tab_name)
            tab_frame = tabview.tab(tab_name)
            builder(tab_frame)

        btn_close = ctk.CTkButton(
            self,
            text="ปิดหน้าต่างคู่มือ",
            font=ctk.CTkFont(family=self.font_family, size=15, weight="bold"),
            fg_color="#334155",
            hover_color="#1e293b",
            height=42,
            width=140,
            corner_radius=8,
            command=self.destroy
        )
        btn_close.pack(pady=(4, 12))

    def _create_card(self, parent, badge_text, badge_color, title_text):
        """สร้างการ์ดข้อมูลมาตรฐาน ตัวหนังสือใหญ่ อ่านง่าย สบายตา"""
        card = ctk.CTkFrame(parent, fg_color="#ffffff", corner_radius=10, border_width=1.5, border_color="#cbd5e1")
        card.pack(fill="x", pady=8, padx=6)

        hdr = ctk.CTkFrame(card, fg_color="transparent")
        hdr.pack(fill="x", padx=16, pady=(12, 6))

        badge = ctk.CTkLabel(
            hdr,
            text=f" {badge_text} ",
            font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"),
            fg_color=badge_color,
            text_color="#ffffff",
            corner_radius=6
        )
        badge.pack(side="left", padx=(0, 10))

        title = ctk.CTkLabel(
            hdr,
            text=title_text,
            font=ctk.CTkFont(family=self.font_family, size=16, weight="bold"),
            text_color="#0f172a"
        )
        title.pack(side="left")

        return card

    def _build_tab_steps(self, parent):
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=4, pady=4)

        # ขั้นตอนที่ 1
        c1 = self._create_card(scroll, "ขั้นตอนที่ 1", "#1d4ed8", "เลือก 'ไฟล์ STAFF TIME (.xlsx)' ประจำงวด")
        desc1 = (
            "• คลิกปุ่ม 'เลือกไฟล์ STAFF TIME...' หรือสามารถ ลากไฟล์มาวาง (Drag & Drop) ลงในกรอบที่ 1 ได้ทันที\n"
            "• เลือกไฟล์หลักของโครงการประจำงวด เช่น '1856 3-4-5 DD.111026-311026.xlsx'\n"
            "• ระบบจะตรวจสอบโครงสร้างและค้นหาชีทรายวันทั้งหมดให้โดยอัตโนมัติ (เช่น ชีท 11.10.26 ถึง 31.10.26)"
        )
        lbl1 = ctk.CTkLabel(c1, text=desc1, font=ctk.CTkFont(family=self.font_family, size=14), text_color="#1e293b", justify="left")
        lbl1.pack(anchor="w", padx=18, pady=(0, 14))

        # ขั้นตอนที่ 2
        c2 = self._create_card(scroll, "ขั้นตอนที่ 2", "#4f46e5", "เลือกไฟล์ข้อมูลนำเข้า (วันลา หรือ ใบแจ้งเวลา)")
        desc2 = (
            "• ช่อง 2.1 (การลาที่ได้รับการอนุมัติ): เลือกไฟล์ 'อนุมัติ.xlsx' ที่ Export ออกมาจากระบบ CNTIWEB\n"
            "• ช่อง 2.2 (ใบแจ้งการมาทำงาน): เลือกไฟล์ 'ใบแจ้งการมาทำงาน.xlsx' ที่ Export จากระบบ CNTIWEB\n"
            "• 💡 สะดวกยิ่งขึ้น: ท่านสามารถลากไฟล์มาวาง (Drag & Drop) ได้ทั้งสองช่อง หรือลากทั้ง 2 ไฟล์พร้อมกันลงโปรแกรมได้ทันที\n"
            "• สามารถเลือกทำพร้อมกันทั้ง 2 ไฟล์ หรือทำทีละไฟล์ก็ได้ตามต้องการ"
        )
        lbl2 = ctk.CTkLabel(c2, text=desc2, font=ctk.CTkFont(family=self.font_family, size=14), text_color="#1e293b", justify="left")
        lbl2.pack(anchor="w", padx=18, pady=(0, 14))

        # ขั้นตอนที่ 3
        c3 = self._create_card(scroll, "ขั้นตอนที่ 3", "#059669", "ตรวจสอบรายการแก้ไข หรือ เริ่มประมวลผลทันที")
        desc3 = (
            "• ปุ่ม '🔍 ตรวจสอบรายการแก้ไขก่อนบันทึกจริง': แนะนำให้คลิกก่อน เพื่อเปิดดูทุกจุดที่โปรแกรมจะทำการบันทึก มีตัวกรองแยกประเภทและช่องค้นหา\n"
            "• ปุ่ม '⚡ เริ่มประมวลผลและบันทึกข้อมูล': บันทึกข้อมูลและสร้างไฟล์ใหม่ทันที พร้อมออกรายงานผล Audit Report .xlsx ให้ครบถ้วน\n"
            "• ปุ่ม '📊 วิเคราะห์เวลาทำงานผิดปกติ': ตรวจจับพนักงานที่ลืมสแกนนิ้ว สแกนรอบเดียว หรือยื่นใบแจ้งเวลาซ้ำ พร้อมส่งออก Excel สรุปผลได้ทันที"
        )
        lbl3 = ctk.CTkLabel(c3, text=desc3, font=ctk.CTkFont(family=self.font_family, size=14), text_color="#1e293b", justify="left")
        lbl3.pack(anchor="w", padx=18, pady=(0, 14))

    def _build_tab_rules(self, parent):
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=4, pady=4)

        # กฎเวลาเข้า-ออก
        c1 = self._create_card(scroll, "กฎเวลาเข้า-ออก", "#d97706", "⏰ กฎเวลาการเข้า-ออกงาน (Attendance Rules)")
        desc1 = (
            "1. เข้างานช่วงเช้า (Check-in ลงช่อง L):\n"
            "   • หากมีใบแจ้งเวลาในช่วง 04:00 - 11:30 น. ➔ บันทึกเวลาเข้างานมาตรฐานเป็น 07:30 อัตโนมัติ\n\n"
            "2. เลิกงานช่วงเย็น (Check-out ลงช่อง M):\n"
            "   • ระบบพิจารณาเวลาตั้งแต่ 16:30 เป็นต้นไป โดยปัดเศษนาทีลงเป็นช่วงละ 30 นาที (:00 หรือ :30) เสมอ\n"
            "     - 16:55  ➔  ปัดลงเป็น  16:30\n"
            "     - 17:00  ➔  คงเดิมที่   17:00\n"
            "     - 17:10  ➔  ปัดลงเป็น  17:00\n"
            "     - 17:33  ➔  ปัดลงเป็น  17:30\n"
            "     - 18:05  ➔  ปัดลงเป็น  18:00\n"
            "     - 18:31  ➔  ปัดลงเป็น  18:30\n\n"
            "3. เติมหมายเหตุอัตโนมัติ:\n"
            "   • ระบบจะเติมคำว่า 'มีใบแจ้งฯ' ลงในช่องประเภทลา (ช่อง G) ให้โดยอัตโนมัติ เพื่อให้ผู้ดูแลตรวจสอบได้ง่าย"
        )
        lbl1 = ctk.CTkLabel(c1, text=desc1, font=ctk.CTkFont(family=self.font_family, size=14), text_color="#1e293b", justify="left")
        lbl1.pack(anchor="w", padx=18, pady=(0, 14))

        # กฎรหัสย่อวันลา
        c2 = self._create_card(scroll, "กฎรหัสวันลา", "#7c3aed", "📋 กฎรหัสย่อวันลา (ช่อง Q) และการตัดคำภาษาไทย")
        desc2 = (
            "1. การตัดคำภาษาไทย:\n"
            "   • ระบบจะตัดข้อความภาษาอังกฤษออก เหลือเพียงชื่อภาษาไทยที่กระชับ เช่น 'ลาป่วย - Sick Leave' ➔ เหลือเพียง 'ลาป่วย'\n\n"
            "2. รหัสตัวย่อลงช่อง Q:\n"
            "   • พ  =  ลาพักร้อน (Annual Leave)\n"
            "   • ป  =  ลาป่วย (Sick Leave)\n"
            "   • อ  =  ลากิจ, ลาไม่ได้รับค่าจ้าง, หรือการลาประเภทอื่นๆ ทั้งหมด\n\n"
            "3. กรณีการลาไม่เต็มวัน (ระบุชั่วโมง):\n"
            "   • ระบบจะระบุจำนวนชั่วโมงชัดเจน เช่น 'ลาป่วย 3 ชม.' และปรับเวลาเข้า-ออกงานให้สอดคล้องกับชั่วโมงที่ลาโดยอัตโนมัติ"
        )
        lbl2 = ctk.CTkLabel(c2, text=desc2, font=ctk.CTkFont(family=self.font_family, size=14), text_color="#1e293b", justify="left")
        lbl2.pack(anchor="w", padx=18, pady=(0, 14))

    def _build_tab_safeguards(self, parent):
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=4, pady=4)

        safeguards = [
            ("ชั้นที่ 1", "#059669", "🛡️ ชั้นที่ 1: ตรวจสอบสแกนนิ้วดิบ (Biometric Scan Protection)",
             "ระบบตรวจช่องสแกนนิ้วดิบ (ช่อง E) หากพบว่าพนักงานมีการสแกนนิ้วครบทั้งเช้า (ตั้งแต่ 04:00 น.) และเลิกงาน (เย็น) แล้ว ระบบจะ 'ข้าม' ข้อมูลวันลาหรือใบแจ้งเวลาในวันนั้นทันที เพื่อป้องกันกรณีพนักงานจำวันผิด กรอกลาผิดวัน หรือแจ้งเวลาซ้ำซ้อน"),
            ("ชั้นที่ 2", "#059669", "🛡️ ชั้นที่ 2: ดับเบิลเช็คชื่อพนักงาน (Double-Check Name Similarity)",
             "ระบบเปรียบเทียบชื่อพนักงานในไฟล์นำเข้ากับชื่อที่ปรากฏในชีท STAFF TIME หากรหัสพนักงานตรงกันแต่ชื่อไม่ตรง (เช่น มีการคีย์รหัสผิดในไฟล์หลัก) ระบบจะข้ามเพื่อความปลอดภัย พร้อมแจ้งเตือนใน Log และ Audit Report ทันที"),
            ("ชั้นที่ 3", "#059669", "🛡️ ชั้นที่ 3: ตรวจจับข้อมูลขัดแย้ง (Conflict Detection)",
             "หากพนักงานคนเดียวกันในวันเดียวกัน มีทั้ง 'วันลาเต็มวัน' และ 'ใบแจ้งการมาทำงาน' ชนกัน ระบบจะยึด 'วันลาที่ได้รับการอนุมัติ' เป็นหลัก และข้ามใบแจ้งเวลา พร้อมแจ้งเตือนในรายงาน"),
            ("ชั้นที่ 4", "#059669", "🛡️ ชั้นที่ 4: ตรวจจับสแกนเดี่ยวซ้ำซ้อน (Redundant Single Scan)",
             "หากพนักงานมีสแกนเช้าจริงอยู่แล้ว แต่ไปยื่นใบแจ้งเวลาเข้างาน ระบบจะข้ามให้ เพราะมีเวลาสแกนจริงอยู่แล้ว เช่นเดียวกันหากมีสแกนเย็น >= 16:30 อยู่แล้ว จะข้ามใบแจ้งเวลาเลิกงาน"),
            ("ชั้นที่ 5", "#059669", "🛡️ ชั้นที่ 5: รายงานผลการตรวจสอบ Audit & Discrepancy Report (.xlsx)",
             "ทุกครั้งที่ประมวลผลเสร็จสิ้น ระบบจะสร้างไฟล์ 'Audit_Report_วันที่.xlsx' บันทึกรายละเอียดครบ 100% ทุกบรรทัด ทั้งรายการที่อัปเดตสำเร็จ และรายการที่ติดเงื่อนไขพร้อมระบุเหตุผลอย่างชัดเจน")
        ]

        for badge, color, title, desc in safeguards:
            c = self._create_card(scroll, badge, color, title)
            lbl = ctk.CTkLabel(c, text=desc, font=ctk.CTkFont(family=self.font_family, size=14), text_color="#1e293b", justify="left")
            lbl.pack(anchor="w", padx=18, pady=(0, 14))

    def _build_tab_columns(self, parent):
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=4, pady=4)

        c = self._create_card(scroll, "ตั้งค่าคอลัมน์", "#0284c7", "📌 ตำแหน่งคอลัมน์มาตรฐาน (รองรับปรับเปลี่ยนได้ทุกโครงการ)")
        desc = (
            "โปรแกรมรองรับการปรับแต่งตำแหน่งคอลัมน์เป้าหมายใน 'ข้อ 3. ตั้งค่าตำแหน่งคอลัมน์' บนหน้าต่างหลัก\n"
            "เผื่อกรณีโครงการอื่นของ CNT ใช้เทมเพลตที่คอลัมน์แตกต่างกัน โดยระบบจะจดจำค่าไว้ถาวรตลอดไป:\n\n"
            "• คอลัมน์ A (รหัสพนักงาน) : รหัสพนักงานประจำแถวนั้น\n"
            "• คอลัมน์ B (ชื่อพนักงาน) : ใช้ทำ Double-Check ตรวจสอบความถูกต้องของชื่อตัวกับไฟล์นำเข้า\n"
            "• คอลัมน์ E (สแกนดิบ) : เวลาสแกนนิ้วดิบ (Time) เช่น '07:28 17:02'\n"
            "• คอลัมน์ G (ประเภทลา) : ลงชื่อวันลาภาษาไทย หรือเติมข้อความ 'มีใบแจ้งฯ'\n"
            "• คอลัมน์ L (เวลาเข้า) : เวลาเข้างาน Check-in (ลง 07:30)\n"
            "• คอลัมน์ M (เวลาออก) : เวลาเลิกงาน Check-out (ลงเวลาปัดเศษลง 30 นาที)\n"
            "• คอลัมน์ Q (รหัสย่อ) : ลงรหัสตัวย่อ พ / ป / อ\n"
            "• เริ่มแถวที่ 2 : แถวแรกที่เริ่มมีข้อมูลพนักงาน (ข้ามหัวตารางแถวที่ 1)\n\n"
            "💡 หากต้องการกลับไปใช้ค่าเริ่มต้น สามารถคลิกปุ่ม 'คืนค่ามาตรฐาน' บนหน้าต่างหลักได้ตลอดเวลา"
        )
        lbl = ctk.CTkLabel(c, text=desc, font=ctk.CTkFont(family=self.font_family, size=14), text_color="#1e293b", justify="left")
        lbl.pack(anchor="w", padx=18, pady=(0, 16))

    def _build_tab_faq(self, parent):
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=4, pady=4)

        faqs = [
            ("คำถามที่ 1", "#0284c7", "Q: ทำไมข้อมูลของพนักงานบางคนถึงไม่ถูกนำไปกรอกลงในชีท?",
             "ตอบ: มีสาเหตุที่เป็นไปได้ตามระบบกันพลาด ดังนี้ครับ:\n"
             "1. พนักงานมีสแกนนิ้วครบทั้งเช้าและเย็นแล้ว (ระบบป้องกันการยื่นลาผิดวัน)\n"
             "2. มีข้อมูลเดิมหรือเวลาเดิมตรงกับเป้าหมายอยู่แล้วในช่องนั้น (ระบบข้ามเพื่อไม่เขียนทับซ้ำ)\n"
             "3. ชื่อพนักงานในชีท STAFF TIME ไม่ตรงกับไฟล์นำเข้า (ระบบข้ามเพื่อป้องกันคีย์รหัสผิด)\n"
             "4. วันที่ในไฟล์นำเข้าไม่ได้อยู่ในช่วงงวดของไฟล์ STAFF TIME\n\n"
             "💡 ท่านสามารถกดปุ่ม 'ตรวจสอบรายการแก้ไข' หรือเปิดไฟล์ 'Audit Report' เพื่อดูเหตุผลรายคนได้ทันที"),

            ("คำถามที่ 2", "#0284c7", "Q: ไฟล์ผลลัพธ์ถูกบันทึกไว้ที่ไหน และไฟล์ต้นฉบับจะเสียหายหรือไม่?",
             "ตอบ: ไฟล์ต้นฉบับจะไม่ถูกแก้ไขหรือทับเลยครับ!\n"
             "โปรแกรมจะสร้างไฟล์ใหม่ชื่อลงท้ายด้วย '_updated_วันเวลา.xlsx' ไว้ในโฟลเดอร์เดียวกับไฟล์ STAFF TIME เดิมเสมอ\n"
             "ท่านสามารถคลิกปุ่ม 'เปิดโฟลเดอร์ผลลัพธ์' บนหน้าต่างหลักเพื่อเปิดโฟลเดอร์ดูไฟล์ได้ทันที"),

            ("คำถามที่ 3", "#0284c7", "Q: หากต้องการดูข้อมูลของพนักงานรายคนแบบเจาะลึกทำอย่างไร?",
             "ตอบ: คลิกปุ่ม 'ตรวจสอบรายคน' บนแถบด้านบนของหน้าต่างหลัก แล้วพิมพ์รหัสหรือชื่อพนักงานเพื่อค้นหาประวัติการทำงานได้ทันที"),

            ("คำถามที่ 4", "#0284c7", "Q: ต้องการส่งโปรแกรมให้เพื่อนร่วมงานที่โครงการอื่นใช้งาน ต้องทำอย่างไร?",
             "ตอบ: สามารถส่งไฟล์ 'LeaveAutomation_1856.exe' หรือ 'LeaveAutomation_1856.zip' ให้เพื่อนร่วมงานได้ทันที\n"
             "โปรแกรมเป็นแบบ Standalone สามารถเปิดใช้งานได้ทันทีโดยไม่ต้องติดตั้ง Python หรือโปรแกรมเสริมใดๆ ทั้งสิ้น")
        ]

        for badge, color, q, a in faqs:
            c = self._create_card(scroll, badge, color, q)
            lbl = ctk.CTkLabel(c, text=a, font=ctk.CTkFont(family=self.font_family, size=14), text_color="#1e293b", justify="left")
            lbl.pack(anchor="w", padx=18, pady=(0, 14))


# -------------------------------------------------------------
# หน้าต่างโมดอล: แสดงรายการที่จะแก้ไขทั้งหมด (Preview All Changes)
# รองรับ Pagination ทรงประสิทธิภาพ ป้องกันกระตุกกรณีไฟล์ใหญ่
# -------------------------------------------------------------

class PreviewDialog(ctk.CTkToplevel):
    def __init__(self, parent, summary, preview_items, on_proceed_callback=None, icon_path=None):
        super().__init__(parent)
        self.parent = parent
        self.summary = summary
        self.all_items = preview_items
        self.filtered_items = preview_items
        self.on_proceed_callback = on_proceed_callback
        self.current_filter = "all"
        self.font_family = "Tahoma"

        # ตัวแปรแบ่งหน้าเพื่อความลื่นไหล (โหลดทันที ไม่กระตุก)
        self.display_limit = 25
        self.page_size = 25

        self.title("ตรวจสอบรายการแก้ไขทั้งหมดในไฟล์ STAFF TIME (Pre-Check Preview)")
        center_window(self, 1120, 780, parent)
        self.minsize(980, 680)

        if icon_path and os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        self.setup_ui()
        self.lift()
        self.attributes("-topmost", True)
        self.after_idle(lambda: self.attributes("-topmost", False))
        self.grab_set()
        self.focus_force()

    def setup_ui(self):
        header = ctk.CTkFrame(self, fg_color="#1a73e8", corner_radius=0, height=80)
        header.pack(fill="x")

        lbl_t = ctk.CTkLabel(
            header,
            text="รายการที่จะแก้ไขทั้งหมดในไฟล์ STAFF TIME (Pre-Check Preview)",
            font=ctk.CTkFont(family=self.font_family, size=20, weight="bold"),
            text_color="#ffffff"
        )
        lbl_t.pack(pady=(12, 2))

        lbl_sub = ctk.CTkLabel(
            header,
            text=f"งวดวันที่: {self.summary.get('date_range', '-')} | ตรวจสอบทุกจุดก่อนบันทึกลงไฟล์จริง ปลอดภัย 100%",
            font=ctk.CTkFont(family=self.font_family, size=13),
            text_color="#fef08a"
        )
        lbl_sub.pack(pady=(0, 12))

        # KPI Summary Bar
        kpi_frame = ctk.CTkFrame(self, fg_color="#f8fafc", corner_radius=10, border_width=1, border_color="#e2e8f0")
        kpi_frame.pack(fill="x", padx=16, pady=(10, 6))
        for col_i in range(5):
            kpi_frame.grid_columnconfigure(col_i, weight=1)

        kpis = [
            ("รายการทั้งหมด", str(self.summary.get('total_items', 0)), "#0f172a", "#f1f5f9"),
            ("จะอัปเดต", str(self.summary.get('updates_count', 0)), "#15803d", "#dcfce7"),
            ("ข้าม (ติดเงื่อนไข)", str(self.summary.get('skips_count', 0)), "#b45309", "#fef3c7"),
            ("ข้อมูลขัดแย้ง", str(self.summary.get('conflict_count', 0)), "#b91c1c", "#fee2e2"),
            ("จำนวนชีทรายวัน", f"{self.summary.get('sheets_count', 0)} ชีท", "#0369a1", "#e0f2fe")
        ]

        for idx, (label, val, fg_txt, bg_pill) in enumerate(kpis):
            card = ctk.CTkFrame(kpi_frame, fg_color=bg_pill, corner_radius=8, border_width=1, border_color="#cbd5e1")
            card.grid(row=0, column=idx, padx=6, pady=8, sticky="ew")
            lbl_l = ctk.CTkLabel(card, text=label, font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"), text_color=fg_txt)
            lbl_l.pack(pady=(6, 1))
            lbl_v = ctk.CTkLabel(card, text=val, font=ctk.CTkFont(family=self.font_family, size=18, weight="bold"), text_color=fg_txt)
            lbl_v.pack(pady=(0, 6))

        # Filter & Search Controls
        filter_bar = ctk.CTkFrame(self, fg_color="transparent")
        filter_bar.pack(fill="x", padx=16, pady=(4, 8))
        filter_bar.grid_columnconfigure(0, weight=1)

        self.search_entry = ctk.CTkEntry(
            filter_bar,
            placeholder_text="ค้นหารหัสพนักงาน, ชื่อพนักงาน, หรือชีทวันที่...",
            font=ctk.CTkFont(family=self.font_family, size=13),
            height=38,
            corner_radius=8
        )
        self.search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.search_entry.bind("<KeyRelease>", lambda e: self.apply_filters())

        self.btn_f_all = ctk.CTkButton(
            filter_bar, text=f"ทั้งหมด ({len(self.all_items)})", font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"),
            fg_color="#1a73e8", width=95, height=38, corner_radius=8,
            command=lambda: self.set_filter("all")
        )
        self.btn_f_all.grid(row=0, column=1, padx=3)

        self.btn_f_upd = ctk.CTkButton(
            filter_bar, text=f"✅ จะอัปเดต ({self.summary.get('updates_count', 0)})", font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"),
            fg_color="#e2e8f0", text_color="#15803d", hover_color="#bbf7d0", width=120, height=38, corner_radius=8,
            command=lambda: self.set_filter("update")
        )
        self.btn_f_upd.grid(row=0, column=2, padx=3)

        self.btn_f_skip = ctk.CTkButton(
            filter_bar, text=f"⚠️ ข้าม ({self.summary.get('skips_count', 0)})", font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"),
            fg_color="#e2e8f0", text_color="#b45309", hover_color="#fde68a", width=105, height=38, corner_radius=8,
            command=lambda: self.set_filter("skips")
        )
        self.btn_f_skip.grid(row=0, column=3, padx=3)

        self.btn_f_conf = ctk.CTkButton(
            filter_bar, text=f"⚡ ข้อมูลขัดแย้ง ({self.summary.get('conflict_count', 0)})", font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"),
            fg_color="#e2e8f0", text_color="#b91c1c", hover_color="#fecaca", width=130, height=38, corner_radius=8,
            command=lambda: self.set_filter("conflict")
        )
        self.btn_f_conf.grid(row=0, column=4, padx=(3, 0))

        self.lbl_counter = ctk.CTkLabel(
            self,
            text=f"แสดง {min(self.display_limit, len(self.all_items))} จากทั้งหมด {len(self.all_items)} รายการ",
            font=ctk.CTkFont(family=self.font_family, size=12),
            text_color="#64748b"
        )
        self.lbl_counter.pack(anchor="w", padx=20, pady=(0, 4))

        # Scrollable Container
        self.scroll_items = ctk.CTkScrollableFrame(self, fg_color="#f1f5f9", corner_radius=10, border_width=1, border_color="#cbd5e1")
        self.scroll_items.pack(fill="both", expand=True, padx=16, pady=(0, 10))
        self.scroll_items.grid_columnconfigure(0, weight=1)

        # Bottom Action Bar
        bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_frame.pack(fill="x", padx=16, pady=(0, 12))
        bottom_frame.grid_columnconfigure(0, weight=1)

        btn_cancel = ctk.CTkButton(
            bottom_frame,
            text="✕ ปิดหน้าต่างตรวจสอบ",
            font=ctk.CTkFont(family=self.font_family, size=14),
            fg_color="#94a3b8",
            hover_color="#64748b",
            height=46,
            width=180,
            corner_radius=8,
            command=self.destroy
        )
        btn_cancel.grid(row=0, column=1, padx=(0, 10))

        if self.on_proceed_callback:
            btn_proceed = ctk.CTkButton(
                bottom_frame,
                text="ยืนยันและเริ่มประมวลผลบันทึกไฟล์ทันที",
                font=ctk.CTkFont(family=self.font_family, size=15, weight="bold"),
                fg_color="#f59e0b",
                hover_color="#d97706",
                height=46,
                corner_radius=8,
                command=self._proceed
            )
            btn_proceed.grid(row=0, column=2)

        self.render_items()

    def set_filter(self, filter_name):
        self.current_filter = filter_name
        self.display_limit = self.page_size
        btns = [
            ("all", self.btn_f_all, "#1a73e8", "#ffffff"),
            ("update", self.btn_f_upd, "#15803d", "#ffffff"),
            ("skips", self.btn_f_skip, "#b45309", "#ffffff"),
            ("conflict", self.btn_f_conf, "#b91c1c", "#ffffff")
        ]
        for f_key, btn, active_bg, active_fg in btns:
            if f_key == filter_name:
                btn.configure(fg_color=active_bg, text_color=active_fg)
            else:
                btn.configure(fg_color="#e2e8f0", text_color="#334155")
        self.apply_filters()

    def apply_filters(self):
        query = self.search_entry.get().strip().lower()
        items = self.all_items

        if self.current_filter == "update":
            items = [x for x in items if x['status_type'] == 'update']
        elif self.current_filter == "skips":
            items = [x for x in items if x['status_type'] != 'update']
        elif self.current_filter == "conflict":
            items = [x for x in items if x['status_type'] == 'conflict']

        if query:
            filtered = []
            for item in items:
                text_corpus = f"{item['sheet']} {item['emp_id']} {item['input_name']} {item['sheet_name']} {item['action']} {item['status']}".lower()
                if query in text_corpus:
                    filtered.append(item)
            items = filtered

        self.filtered_items = items
        self.update_counter()
        self.render_items()

    def update_counter(self):
        showing = min(self.display_limit, len(self.filtered_items))
        self.lbl_counter.configure(text=f"แสดง {showing} จากที่กรองพบ {len(self.filtered_items)} รายการ (ทั้งหมด {len(self.all_items)})")

    def _load_more(self):
        self.display_limit += self.page_size
        self.update_counter()
        self.render_items()

    def _load_all(self):
        self.display_limit = len(self.filtered_items)
        self.update_counter()
        self.render_items()

    def render_items(self):
        for widget in self.scroll_items.winfo_children():
            widget.destroy()

        if not self.filtered_items:
            empty_lbl = ctk.CTkLabel(
                self.scroll_items,
                text="ไม่พบรายการที่ตรงกับเงื่อนไขการค้นหา",
                font=ctk.CTkFont(family=self.font_family, size=15),
                text_color="#94a3b8"
            )
            empty_lbl.pack(pady=40)
            return

        items_to_render = self.filtered_items[:self.display_limit]

        for idx, item in enumerate(items_to_render, 1):
            st = item['status_type']
            if st == 'update':
                card_bg = "#ffffff"
                border_col = "#86efac"
                badge_bg = "#dcfce7"
                badge_fg = "#15803d"
            elif st == 'conflict' or st == 'mismatch':
                card_bg = "#ffffff"
                border_col = "#fca5a5"
                badge_bg = "#fee2e2"
                badge_fg = "#b91c1c"
            else:
                card_bg = "#ffffff"
                border_col = "#cbd5e1"
                badge_bg = "#fef3c7"
                badge_fg = "#b45309"

            card = ctk.CTkFrame(self.scroll_items, fg_color=card_bg, corner_radius=10, border_width=1.5, border_color=border_col)
            card.pack(fill="x", pady=4, padx=2)
            card.grid_columnconfigure(1, weight=1)

            top_line = ctk.CTkFrame(card, fg_color="transparent")
            top_line.pack(fill="x", padx=14, pady=(10, 4))
            top_line.grid_columnconfigure(0, weight=1)

            emp_disp = f"รหัส {item['emp_id']} : {item['input_name']}"
            if item['sheet_name'] and item['sheet_name'] != '-':
                emp_disp += f" (ในชีท: {item['sheet_name']})"

            lbl_emp = ctk.CTkLabel(
                top_line,
                text=f"[{idx:03d}] ชีท {item['sheet']}  |  {emp_disp}  |  ประเภท: {item['type']}",
                font=ctk.CTkFont(family=self.font_family, size=13, weight="bold"),
                text_color="#1e293b"
            )
            lbl_emp.grid(row=0, column=0, sticky="w")

            badge = ctk.CTkLabel(
                top_line,
                text=f" {item['status']} ",
                font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"),
                text_color=badge_fg,
                fg_color=badge_bg,
                corner_radius=6
            )
            badge.grid(row=0, column=1, sticky="e")

            mid_line = ctk.CTkFrame(card, fg_color="transparent")
            mid_line.pack(fill="x", padx=14, pady=(0, 4))
            
            action_text = f"การดำเนินการ: {item['action']}"
            lbl_action = ctk.CTkLabel(
                mid_line,
                text=action_text,
                font=ctk.CTkFont(family=self.font_family, size=13, weight="bold" if st == 'update' else "normal"),
                text_color="#15803d" if st == 'update' else "#334155"
            )
            lbl_action.pack(anchor="w")

            bot_line = ctk.CTkFrame(card, fg_color="transparent")
            bot_line.pack(fill="x", padx=14, pady=(0, 10))
            
            sub_info = f"ช่องเป้าหมาย: [{item.get('target_cells', '-')}]  |  สแกนจริงในชีท (ช่อง E): {item.get('scans', '-')}"
            if item.get('detail'):
                sub_info += f"  |  รายละเอียด: {item['detail']}"

            lbl_sub = ctk.CTkLabel(
                bot_line,
                text=sub_info,
                font=ctk.CTkFont(family=self.font_family, size=12),
                text_color="#64748b"
            )
            lbl_sub.pack(anchor="w")

        # ปุ่มโหลดเพิ่มเติมหากรายการเกิน limit
        if len(self.filtered_items) > self.display_limit:
            more_frame = ctk.CTkFrame(self.scroll_items, fg_color="transparent")
            more_frame.pack(fill="x", pady=12)
            more_frame.grid_columnconfigure(0, weight=1)
            more_frame.grid_columnconfigure(1, weight=1)

            btn_more = ctk.CTkButton(
                more_frame,
                text=f"➕ แสดงเพิ่มอีก {self.page_size} รายการ (กำลังแสดง {self.display_limit} จาก {len(self.filtered_items)})",
                font=ctk.CTkFont(family=self.font_family, size=13, weight="bold"),
                fg_color="#e0f2fe",
                text_color="#0369a1",
                hover_color="#bae6fd",
                height=40,
                corner_radius=8,
                command=self._load_more
            )
            btn_more.grid(row=0, column=0, padx=6, sticky="ew")

            btn_all = ctk.CTkButton(
                more_frame,
                text=f"⚡ แสดงทั้งหมด ({len(self.filtered_items)} รายการ)",
                font=ctk.CTkFont(family=self.font_family, size=13),
                fg_color="#f1f5f9",
                text_color="#475569",
                hover_color="#e2e8f0",
                height=40,
                corner_radius=8,
                command=self._load_all
            )
            btn_all.grid(row=0, column=1, padx=6, sticky="ew")

    def _proceed(self):
        self.destroy()
        if self.on_proceed_callback:
            self.on_proceed_callback()


# -------------------------------------------------------------
# หน้าต่างโมดอล: ตรวจสอบรายคน (Employee Quick Inspector Dialog)
# -------------------------------------------------------------

class InspectorDialog(ctk.CTkToplevel):
    def __init__(self, parent, main_file, appr_file, att_file, col_config=None, icon_path=None):
        super().__init__(parent)
        self.parent = parent
        self.main_file = main_file
        self.appr_file = appr_file
        self.att_file = att_file
        self.col_config = col_config
        self.font_family = "Tahoma"

        self.title("ระบบตรวจสอบข้อมูลรายบุคคล (Employee Quick Inspector)")
        center_window(self, 1000, 740, parent)
        self.minsize(880, 620)

        if icon_path and os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        self.setup_ui()
        self.lift()
        self.attributes("-topmost", True)
        self.after_idle(lambda: self.attributes("-topmost", False))
        self.grab_set()
        self.focus_force()

    def setup_ui(self):
        header = ctk.CTkFrame(self, fg_color="#0f172a", corner_radius=0, height=80)
        header.pack(fill="x")

        lbl_t = ctk.CTkLabel(
            header,
            text="ระบบตรวจสอบข้อมูลรายบุคคล (Employee Quick Inspector)",
            font=ctk.CTkFont(family=self.font_family, size=20, weight="bold"),
            text_color="#ffffff"
        )
        lbl_t.pack(pady=(12, 2))

        lbl_sub = ctk.CTkLabel(
            header,
            text="ตรวจสอบประวัติรายวันในไฟล์ STAFF TIME, วันลาที่ได้รับอนุมัติ และใบแจ้งเวลาทำงานของพนักงานรายบุคคล",
            font=ctk.CTkFont(family=self.font_family, size=13),
            text_color="#fef08a"
        )
        lbl_sub.pack(pady=(0, 12))

        search_card = ctk.CTkFrame(self, fg_color="#ffffff", corner_radius=10, border_width=1, border_color="#cbd5e1")
        search_card.pack(fill="x", padx=16, pady=10)
        search_card.grid_columnconfigure(0, weight=1)

        self.txt_query = ctk.CTkEntry(
            search_card,
            placeholder_text="พิมพ์รหัสพนักงาน หรือ ชื่อพนักงาน (เช่น 5539 หรือ นริศร)...",
            font=ctk.CTkFont(family=self.font_family, size=14),
            height=42,
            corner_radius=8
        )
        self.txt_query.grid(row=0, column=0, padx=(14, 10), pady=12, sticky="ew")
        self.txt_query.bind("<Return>", lambda e: self.do_search())

        btn_search = ctk.CTkButton(
            search_card,
            text="ค้นหาข้อมูล",
            font=ctk.CTkFont(family=self.font_family, size=14, weight="bold"),
            fg_color="#1d4ed8",
            hover_color="#1557b0",
            width=110,
            height=42,
            corner_radius=8,
            command=self.do_search
        )
        btn_search.grid(row=0, column=1, padx=(0, 8), pady=12)

        btn_clear = ctk.CTkButton(
            search_card,
            text="✕ ล้าง",
            font=ctk.CTkFont(family=self.font_family, size=14),
            fg_color="#e2e8f0",
            hover_color="#cbd5e1",
            text_color="#334155",
            width=80,
            height=42,
            corner_radius=8,
            command=self.clear_search
        )
        btn_clear.grid(row=0, column=2, padx=(0, 14), pady=12)

        self.results_container = ctk.CTkFrame(self, fg_color="transparent")
        self.results_container.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        self.show_initial_state()

    def show_initial_state(self):
        for w in self.results_container.winfo_children():
            w.destroy()

        hint_card = ctk.CTkFrame(self.results_container, fg_color="#ffffff", corner_radius=10, border_width=1, border_color="#cbd5e1")
        hint_card.pack(fill="both", expand=True)

        lbl_hint = ctk.CTkLabel(
            hint_card,
            text="กรุณาพิมพ์รหัสพนักงาน หรือ ชื่อพนักงานในช่องค้นหาด้านบน แล้วกดปุ่ม 'ค้นหาข้อมูล'\nระบบจะรวบรวมประวัติการสแกนนิ้ว, วันลาที่อนุมัติ และใบแจ้งเวลาทั้งหมดของพนักงานท่านนั้นมาแสดงทันที",
            font=ctk.CTkFont(family=self.font_family, size=14),
            text_color="#64748b",
            justify="center"
        )
        lbl_hint.pack(expand=True, pady=40)

    def clear_search(self):
        self.txt_query.delete(0, 'end')
        self.show_initial_state()

    def do_search(self):
        query = self.txt_query.get().strip()
        if not query:
            return

        if not self.main_file or not os.path.exists(self.main_file):
            messagebox.showerror("แจ้งเตือน", "กรุณาเลือกไฟล์ STAFF TIME (.xlsx) ให้ถูกต้องก่อนค้นหา")
            return

        for w in self.results_container.winfo_children():
            w.destroy()

        lbl_loading = ctk.CTkLabel(
            self.results_container,
            text=f"⏳ กำลังค้นหาข้อมูลของ '{query}' ในไฟล์ STAFF TIME...",
            font=ctk.CTkFont(family=self.font_family, size=15),
            text_color="#1a73e8"
        )
        lbl_loading.pack(pady=40)
        self.update_idletasks()

        info = get_employee_details(query, self.main_file, self.appr_file, self.att_file, self.col_config)

        for w in self.results_container.winfo_children():
            w.destroy()

        if not info:
            not_found_card = ctk.CTkFrame(self.results_container, fg_color="#ffffff", corner_radius=10, border_width=1, border_color="#fca5a5")
            not_found_card.pack(fill="both", expand=True)
            lbl_nf = ctk.CTkLabel(
                not_found_card,
                text=f"❌ ไม่พบข้อมูลพนักงานที่ตรงกับคำค้นหา: '{query}'\nกรุณาตรวจสอบรหัสหรือชื่อพนักงานอีกครั้ง",
                font=ctk.CTkFont(family=self.font_family, size=15, weight="bold"),
                text_color="#b91c1c"
            )
            lbl_nf.pack(expand=True, pady=40)
            return

        profile_card = ctk.CTkFrame(self.results_container, fg_color="#ffffff", corner_radius=10, border_width=1.5, border_color="#1a73e8")
        profile_card.pack(fill="x", pady=(0, 8))
        profile_card.grid_columnconfigure(0, weight=1)

        top_p = ctk.CTkFrame(profile_card, fg_color="#eff6ff", corner_radius=8)
        top_p.pack(fill="x", padx=10, pady=10)
        top_p.grid_columnconfigure(0, weight=1)

        lbl_name = ctk.CTkLabel(
            top_p,
            text=f"รหัสพนักงาน: {info['emp_id']}  |  ชื่อ-นามสกุล: {info['emp_name']}",
            font=ctk.CTkFont(family=self.font_family, size=17, weight="bold"),
            text_color="#1e40af"
        )
        lbl_name.grid(row=0, column=0, sticky="w", padx=14, pady=8)

        stats_line = (
            f"จำนวนวันในงวด: {info['total_days_in_period']} วัน  |  "
            f"⏱️ วันที่มีสแกนนิ้ว: {info['days_with_scans']} วัน   |   "
            f"วันลาที่บันทึกแล้ว: {info['days_with_leaves']} วัน  |  "
            f"วันลาที่อนุมัติ: {info['approved_leaves_count']} รายการ  |  "
            f"⏰ ใบแจ้งเวลา: {info['attendance_notifications_count']} ฉบับ"
        )
        lbl_stats = ctk.CTkLabel(
            profile_card,
            text=stats_line,
            font=ctk.CTkFont(family=self.font_family, size=13),
            text_color="#475569"
        )
        lbl_stats.pack(anchor="w", padx=14, pady=(0, 10))

        tabview = ctk.CTkTabview(self.results_container, corner_radius=10)
        tabview.pack(fill="both", expand=True)

        tabview.add("ตารางรายวันใน STAFF TIME")
        tabview.add(f"การลาที่ได้รับอนุมัติ ({info['approved_leaves_count']})")
        tabview.add(f"ใบแจ้งการมาทำงาน ({info['attendance_notifications_count']})")

        # 1. Daily Tab
        t1 = tabview.tab("ตารางรายวันใน STAFF TIME")
        s1 = ctk.CTkScrollableFrame(t1, fg_color="transparent")
        s1.pack(fill="both", expand=True)

        for d in info['daily_records']:
            row_card = ctk.CTkFrame(s1, fg_color="#ffffff", corner_radius=8, border_width=1, border_color="#cbd5e1")
            row_card.pack(fill="x", pady=3, padx=2)
            row_card.grid_columnconfigure(0, weight=1)

            d_str = d['date'].strftime('%d/%m/%Y')
            status_text = []
            if d['has_in'] and d['has_out']:
                status_text.append("✅ สแกนครบ")
            elif d['has_in']:
                status_text.append("⚠️ สแกนเฉพาะเข้า")
            elif d['has_out']:
                status_text.append("⚠️ สแกนเฉพาะออก")
            else:
                status_text.append("❌ ไม่มีสแกน")

            if d['leave'] != '-':
                status_text.append(f"วันลา: {d['leave']}")

            top_t = f"{d_str} (ชีท {d['sheet']})  |  {' '.join(status_text)}"
            lbl_top = ctk.CTkLabel(row_card, text=top_t, font=ctk.CTkFont(family=self.font_family, size=13, weight="bold"), text_color="#0f172a")
            lbl_top.pack(anchor="w", padx=12, pady=(6, 2))

            det_t = f"สแกนจริง (E): {d['scans']}   |   เข้า (L): {d['in_time']}   |   ออก (M): {d['out_time']}   |   รหัสย่อ (Q): {d['code']}"
            lbl_det = ctk.CTkLabel(row_card, text=det_t, font=ctk.CTkFont(family=self.font_family, size=12), text_color="#64748b")
            lbl_det.pack(anchor="w", padx=12, pady=(0, 6))

        # 2. Approved Leaves Tab
        t2 = tabview.tab(f"การลาที่ได้รับอนุมัติ ({info['approved_leaves_count']})")
        s2 = ctk.CTkScrollableFrame(t2, fg_color="transparent")
        s2.pack(fill="both", expand=True)

        if not info['approved_leaves']:
            lbl_no = ctk.CTkLabel(s2, text="ไม่พบข้อมูลวันลาในไฟล์การลาที่ได้รับการอนุมัติ", font=ctk.CTkFont(family=self.font_family, size=13), text_color="#94a3b8")
            lbl_no.pack(pady=20)
        else:
            for idx, l in enumerate(info['approved_leaves'], 1):
                c = ctk.CTkFrame(s2, fg_color="#ffffff", corner_radius=8, border_width=1, border_color="#cbd5e1")
                c.pack(fill="x", pady=3, padx=2)
                d_str = l['date'].strftime('%d/%m/%Y')
                lbl_l1 = ctk.CTkLabel(c, text=f"[{idx}] วันที่: {d_str}   |   ประเภท: {l['leave_type_raw']}", font=ctk.CTkFont(family=self.font_family, size=13, weight="bold"), text_color="#b45309")
                lbl_l1.pack(anchor="w", padx=12, pady=(6, 2))
                hrs_display = f"{l['hours']:g} ชม." if l.get('hours') is not None else ""
                lbl_l2 = ctk.CTkLabel(c, text=f"จำนวน: {'ลาไม่เต็มวัน ' + hrs_display if l['is_partial'] else 'ลาเต็มวัน'}   |   ชื่อที่ระบุ: {l['emp_raw']}", font=ctk.CTkFont(family=self.font_family, size=12), text_color="#64748b")
                lbl_l2.pack(anchor="w", padx=12, pady=(0, 6))

        # 3. Attendance Notifications Tab
        t3 = tabview.tab(f"ใบแจ้งการมาทำงาน ({info['attendance_notifications_count']})")
        s3 = ctk.CTkScrollableFrame(t3, fg_color="transparent")
        s3.pack(fill="both", expand=True)

        if not info['attendance_records']:
            lbl_no = ctk.CTkLabel(s3, text="ไม่พบข้อมูลในใบแจ้งการมาทำงาน", font=ctk.CTkFont(family=self.font_family, size=13), text_color="#94a3b8")
            lbl_no.pack(pady=20)
        else:
            for idx, a in enumerate(info['attendance_records'], 1):
                c = ctk.CTkFrame(s3, fg_color="#ffffff", corner_radius=8, border_width=1, border_color="#cbd5e1")
                c.pack(fill="x", pady=3, padx=2)
                d_str = a['date'].strftime('%d/%m/%Y')
                t_lbl = f"{a['target_time'].hour}:{a['target_time'].minute:02d}"
                side_lbl = "ช่วงเช้า (Check-in 07:30)" if a['target_col'] == 'L' else "ช่วงเย็น (Check-out ปัดเศษ 30 นาที)"
                lbl_a1 = ctk.CTkLabel(c, text=f"[{idx}] วันที่: {d_str}   |   {side_lbl}", font=ctk.CTkFont(family=self.font_family, size=13, weight="bold"), text_color="#15803d")
                lbl_a1.pack(anchor="w", padx=12, pady=(6, 2))
                lbl_a2 = ctk.CTkLabel(c, text=f"เวลาที่แจ้งมา: {a['orig_time']}   ➔   ปรับเป็น: {t_lbl} ลงช่อง {a['target_col']}", font=ctk.CTkFont(family=self.font_family, size=12), text_color="#64748b")
                lbl_a2.pack(anchor="w", padx=12, pady=(0, 6))


# -------------------------------------------------------------
# หน้าต่างโมดอล: สถิติภาพรวมโครงการ (STAFF TIME Analytics Dialog)
# -------------------------------------------------------------

class AnalyticsDialog(ctk.CTkToplevel):
    def __init__(self, parent, main_file, appr_file, att_file, col_config=None, icon_path=None):
        super().__init__(parent)
        self.parent = parent
        self.main_file = main_file
        self.appr_file = appr_file
        self.att_file = att_file
        self.col_config = col_config
        self.font_family = "Tahoma"

        self.title("สถิติภาพรวมเวลาทำงานและวันลา (STAFF TIME Analytics)")
        center_window(self, 960, 720, parent)
        self.minsize(860, 620)

        if icon_path and os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        self.setup_ui()
        self.lift()
        self.attributes("-topmost", True)
        self.after_idle(lambda: self.attributes("-topmost", False))
        self.grab_set()
        self.focus_force()

    def setup_ui(self):
        header = ctk.CTkFrame(self, fg_color="#0f172a", corner_radius=0, height=80)
        header.pack(fill="x")

        lbl_t = ctk.CTkLabel(
            header,
            text="สถิติภาพรวมเวลาทำงานและวันลา (STAFF TIME Analytics)",
            font=ctk.CTkFont(family=self.font_family, size=20, weight="bold"),
            text_color="#ffffff"
        )
        lbl_t.pack(pady=(12, 2))

        lbl_sub = ctk.CTkLabel(
            header,
            text="สรุปสถิติภาพรวมการลงเวลาทำงาน วันลาที่อนุมัติ และใบแจ้งเวลาประจำงวดของหน่วยงาน",
            font=ctk.CTkFont(family=self.font_family, size=13),
            text_color="#cbd5e1"
        )
        lbl_sub.pack(pady=(0, 12))

        scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=16, pady=10)

        lbl_load = ctk.CTkLabel(scroll, text="⏳ กำลังประมวลผลข้อมูลสถิติภาพรวม...", font=ctk.CTkFont(family=self.font_family, size=15), text_color="#1a73e8")
        lbl_load.pack(pady=40)
        self.update_idletasks()

        data = get_project_analytics(self.main_file, self.appr_file, self.att_file, self.col_config)
        lbl_load.destroy()

        sm = data['summary']

        kpi_frame = ctk.CTkFrame(scroll, fg_color="#ffffff", corner_radius=10, border_width=1, border_color="#cbd5e1")
        kpi_frame.pack(fill="x", pady=(0, 14))
        for c_i in range(3):
            kpi_frame.grid_columnconfigure(c_i, weight=1)

        kpi_list = [
            ("จำนวนพนักงานทั้งหมด", f"{sm.get('total_employees', 0)} คน", "#1d4ed8", "#eff6ff"),
            ("งวดเวลาทำงาน", f"{sm.get('sheets_count', 0)} วัน", "#0f766e", "#f0fdfa"),
            ("ช่วงวันที่", sm.get('date_range', '-'), "#0369a1", "#f0f9ff"),
            ("ข้อมูลวันลาทั้งหมด", f"{sm.get('leave_updates', 0) + sm.get('leave_skips', 0)} รายการ", "#b45309", "#fffbeb"),
            ("⏰ ข้อมูลใบแจ้งเวลาทั้งหมด", f"{sm.get('attend_updates', 0) + sm.get('attend_skips', 0)} ฉบับ", "#15803d", "#f0fdf4"),
            ("รายการที่ระบบกันพลาดทำงาน", f"{sm.get('skip_scan_count', 0) + sm.get('conflict_count', 0)} จุด", "#b91c1c", "#fef2f2")
        ]

        for idx, (title, val, fg, bg) in enumerate(kpi_list):
            row_idx = idx // 3
            col_idx = idx % 3
            p = ctk.CTkFrame(kpi_frame, fg_color=bg, corner_radius=8, border_width=1, border_color="#e2e8f0")
            p.grid(row=row_idx, column=col_idx, padx=8, pady=8, sticky="ew")
            l1 = ctk.CTkLabel(p, text=title, font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"), text_color=fg)
            l1.pack(pady=(8, 2))
            l2 = ctk.CTkLabel(p, text=val, font=ctk.CTkFont(family=self.font_family, size=16, weight="bold"), text_color=fg)
            l2.pack(pady=(0, 8))

        # Section 1: Leave Breakdown (อ่านตรงจาก leave_type)
        sec1 = ctk.CTkFrame(scroll, fg_color="#ffffff", corner_radius=10, border_width=1, border_color="#cbd5e1")
        sec1.pack(fill="x", pady=(0, 14))
        lbl_s1_t = ctk.CTkLabel(sec1, text="สถิติแยกตามประเภทการลา (Leave Breakdown)", font=ctk.CTkFont(family=self.font_family, size=15, weight="bold"), text_color="#b45309")
        lbl_s1_t.pack(anchor="w", padx=16, pady=(12, 6))

        if data['leave_breakdown']:
            for l_type, cnt in sorted(data['leave_breakdown'].items(), key=lambda x: -x[1]):
                r_box = ctk.CTkFrame(sec1, fg_color="#fffbeb", corner_radius=6)
                r_box.pack(fill="x", padx=16, pady=3)
                r_box.grid_columnconfigure(0, weight=1)
                l_lbl = ctk.CTkLabel(r_box, text=f"• {l_type}", font=ctk.CTkFont(family=self.font_family, size=13), text_color="#92400e")
                l_lbl.grid(row=0, column=0, sticky="w", padx=10, pady=6)
                c_lbl = ctk.CTkLabel(r_box, text=f"{cnt} รายการ", font=ctk.CTkFont(family=self.font_family, size=13, weight="bold"), text_color="#b45309")
                c_lbl.grid(row=0, column=1, sticky="e", padx=10, pady=6)
        else:
            lbl_no_l = ctk.CTkLabel(sec1, text="ไม่มีข้อมูลวันลาในงวดนี้", font=ctk.CTkFont(family=self.font_family, size=13), text_color="#94a3b8")
            lbl_no_l.pack(anchor="w", padx=16, pady=8)

        # Section 2: Attendance Breakdown
        sec2 = ctk.CTkFrame(scroll, fg_color="#ffffff", corner_radius=10, border_width=1, border_color="#cbd5e1")
        sec2.pack(fill="x", pady=(0, 14))
        lbl_s2_t = ctk.CTkLabel(sec2, text="⏰ สถิติใบแจ้งการมาทำงาน (Attendance Breakdown)", font=ctk.CTkFont(family=self.font_family, size=15, weight="bold"), text_color="#15803d")
        lbl_s2_t.pack(anchor="w", padx=16, pady=(12, 6))

        for a_type, cnt in data['attend_breakdown'].items():
            r_box = ctk.CTkFrame(sec2, fg_color="#f0fdf4", corner_radius=6)
            r_box.pack(fill="x", padx=16, pady=3)
            r_box.grid_columnconfigure(0, weight=1)
            l_lbl = ctk.CTkLabel(r_box, text=f"• {a_type}", font=ctk.CTkFont(family=self.font_family, size=13), text_color="#166534")
            l_lbl.grid(row=0, column=0, sticky="w", padx=10, pady=6)
            c_lbl = ctk.CTkLabel(r_box, text=f"{cnt} รายการ", font=ctk.CTkFont(family=self.font_family, size=13, weight="bold"), text_color="#15803d")
            c_lbl.grid(row=0, column=1, sticky="e", padx=10, pady=6)

        # Section 3: Safeguards
        sec3 = ctk.CTkFrame(scroll, fg_color="#ffffff", corner_radius=10, border_width=1, border_color="#cbd5e1")
        sec3.pack(fill="x", pady=(0, 14))
        lbl_s3_t = ctk.CTkLabel(sec3, text="การตรวจสอบความถูกต้องและความปลอดภัย (Safeguard Status)", font=ctk.CTkFont(family=self.font_family, size=15, weight="bold"), text_color="#1e293b")
        lbl_s3_t.pack(anchor="w", padx=16, pady=(12, 6))

        sg_items = [
            ("• สแกนนิ้วครบแล้ว (ข้ามวันลา/แจ้งเวลา เพื่อกันพลาด)", sm.get('skip_scan_count', 0)),
            ("• ข้อมูลเดิมตรงกันอยู่แล้วในตาราง", sm.get('skip_dup_count', 0)),
            ("• ข้อมูลขัดแย้ง (พบทั้งวันลาและใบแจ้งเวลาในวันเดียวกัน)", sm.get('conflict_count', 0)),
            ("• ไม่พบรหัสพนักงานในชีทของวันนั้น", sm.get('not_found_count', 0))
        ]

        for s_title, s_cnt in sg_items:
            r_box = ctk.CTkFrame(sec3, fg_color="#f8fafc", corner_radius=6)
            r_box.pack(fill="x", padx=16, pady=3)
            r_box.grid_columnconfigure(0, weight=1)
            l_lbl = ctk.CTkLabel(r_box, text=s_title, font=ctk.CTkFont(family=self.font_family, size=13), text_color="#334155")
            l_lbl.grid(row=0, column=0, sticky="w", padx=10, pady=6)
            c_lbl = ctk.CTkLabel(r_box, text=f"{s_cnt} จุด", font=ctk.CTkFont(family=self.font_family, size=13, weight="bold"), text_color="#0f172a")
            c_lbl.grid(row=0, column=1, sticky="e", padx=10, pady=6)

        sec3_pad = ctk.CTkFrame(sec3, fg_color="transparent", height=10)
        sec3_pad.pack()


# -------------------------------------------------------------
# -------------------------------------------------------------
# ระบบวิเคราะห์เวลาทำงานและสแกนผิดปกติอัจฉริยะ (Smart Anomaly Analyzer)
# -------------------------------------------------------------

def analyze_attendance_anomalies(main_source, col_config=None):
    """
    วิเคราะห์ข้อมูลสแกนดิบและเวลาทำงานในทุกชีท เพื่อหาความผิดปกติ:
    1. missing_punch: สแกนขาเดียว (มีเฉพาะเข้า หรือมีเฉพาะออก)
    2. unrecorded_absence: ขาดงานโดยไม่แจ้ง (วันทำงานปกติ ไม่มีสแกน และไม่มีข้อมูลการลา)

    รองรับทั้ง main_source เป็น filepath (str) หรือ openpyxl.Workbook
    จัดกลุ่มตามชื่อพนักงาน และเรียงตามลำดับวันที่โดยอัตโนมัติ (หาก 1 คนมีหลายวัน จะเรียงติดกัน)
    """
    cfg = col_config or get_default_column_config()
    if isinstance(main_source, str):
        if not main_source or not os.path.exists(main_source):
            return {'success': False, 'error': 'ไม่พบไฟล์รายงาน STAFF TIME'}
        try:
            wb = openpyxl.load_workbook(main_source, data_only=True)
        except Exception as e:
            return {'success': False, 'error': f'ไม่สามารถเปิดไฟล์ Excel ได้: {str(e)}'}
    elif isinstance(main_source, openpyxl.Workbook):
        wb = main_source
    else:
        return {'success': False, 'error': 'ข้อมูลไฟล์ไม่ถูกต้อง'}

    sheets = map_workbook_sheets(wb)
    if not sheets:
        return {'success': False, 'error': 'ไม่พบชีทวันที่ประจำงวดในไฟล์ STAFF TIME'}

    anomalies = []
    counts = {
        'total': 0,
        'missing_punch': 0,
        'unrecorded_absence': 0,
        'sheets_checked': len(sheets)
    }

    start_row = int(cfg.get('start_row', 2))
    emp_id_col = cfg.get('emp_id_col', 'A').upper().strip() or 'A'
    emp_name_col = cfg.get('emp_name_col', 'B').upper().strip() or 'B'
    scan_col = cfg.get('scan_col', 'E').upper().strip() or 'E'
    leave_type_col = cfg.get('leave_type_col', 'G').upper().strip() or 'G'
    time_in_col = cfg.get('time_in_col', 'J').upper().strip() or 'J'
    time_out_col = cfg.get('time_out_col', 'K').upper().strip() or 'K'
    checkin_col = cfg.get('checkin_col', 'L').upper().strip() or 'L'
    checkout_col = cfg.get('checkout_col', 'M').upper().strip() or 'M'
    leave_code_col = cfg.get('leave_code_col', 'Q').upper().strip() or 'Q'

    for dt, sheet_name in sheets.items():
        ws = wb[sheet_name]
        is_holiday, h_name = is_company_holiday(dt, ws)
        is_sunday = (dt.weekday() == 6)

        for r in range(start_row, ws.max_row + 1):
            emp_id = str(ws[f"{emp_id_col}{r}"].value or '').strip()
            emp_name = str(ws[f"{emp_name_col}{r}"].value or '').strip()
            if not emp_id or not emp_name:
                continue

            # 1. ตรวจสอบข้อมูลการลาและรหัสวันลา (Col G และ Col Q)
            leave_val = str(ws[f"{leave_type_col}{r}"].value or '').strip()
            has_leave = bool(leave_val and leave_val.lower() != 'none')
            
            code_val = str(ws[f"{leave_code_col}{r}"].value or '').strip()
            has_q_code = any(c in code_val for c in ['พ', 'ลป', 'ลก', 'น', 'บ', 'ค', 'ส', 'อ', 'ป']) and not code_val.startswith('=')

            # หากมีรายการแจ้งในระบบแล้ว (มีใบแจ้งฯ หรือมีบันทึกการลาในช่อง G หรือรหัสวันลาในช่อง Q)
            has_system_notification = has_leave or has_q_code

            # 2. เวลาค่าจริงจากสแกนนิ้ว: ช่อง J (Time-In) และ K (Time-Out)
            j_val = ws[f"{time_in_col}{r}"].value
            k_val = ws[f"{time_out_col}{r}"].value
            has_j = is_valid_time(j_val)
            has_k = is_valid_time(k_val)

            # 3. เวลาค่าที่ปรับตามการลาแล้ว: ช่อง L (Check-in) และ M (Check-Out)
            l_val = ws[f"{checkin_col}{r}"].value
            m_val = ws[f"{checkout_col}{r}"].value
            has_l = is_valid_time(l_val)
            has_m = is_valid_time(m_val)

            # 4. สแกนดิบในช่อง E
            scan_val = ws[f"{scan_col}{r}"].value
            has_in_raw, has_out_raw, scans_list, e_dec, l_dec = check_fingerprint_scans(scan_val)

            # === ดับเบิลเช็คทั้งค่าจริงและค่าที่ปรับตามการลาแล้ว ===
            # ยืนยันเวลาเข้า (IN) หาก:
            # - ช่อง L (ปรับตามการลาแล้ว) มีเวลาที่ถูกต้อง หรือ
            # - ช่อง J (ค่าจริงจากสแกน) มีเวลาที่ถูกต้อง หรือ
            # - สแกนดิบมีเวลาเข้า
            in_confirmed = has_l or has_j or has_in_raw

            # ยืนยันเวลาออก (OUT) หาก:
            # - ช่อง M (ปรับตามการลาแล้ว) มีเวลาที่ถูกต้อง หรือ
            # - ช่อง K (ค่าจริงจากสแกน) มีเวลาที่ถูกต้อง หรือ
            # - สแกนดิบมีเวลาออก
            out_confirmed = has_m or has_k or has_out_raw

            # กรณีพิเศษ: มีสแกนตั้งแต่ 2 เวลาขึ้นไปและห่างกัน >= 1 ชม.
            if len(scans_list) >= 2 and e_dec is not None and l_dec is not None and (l_dec - e_dec) >= 1.0:
                in_confirmed = True
                out_confirmed = True

            # ถ้าทั้งเวลาเข้าและเวลาออกได้รับการยืนยันครบถ้วน -> ข้อมูลสมบูรณ์ ไม่มีความผิดปกติ
            if in_confirmed and out_confirmed:
                continue

            # หากมีรายการแจ้งในระบบแล้ว (มีใบลา หรือใบแจ้งฯ ในช่อง G หรือรหัสวันลาในช่อง Q) -> ข้อมูลได้รับอนุมัติ/บันทึกแล้ว ไม่ถือว่าไม่มีแจ้ง
            if has_system_notification:
                continue

            # --- ตรวจจับความผิดปกติที่แท้จริง ---
            # 1. สแกนขาเดียว (Missing Punch) - มีเข้าแต่ขาดออก หรือ มีออกแต่ขาดเข้า
            if in_confirmed and not out_confirmed:
                time_in_str = format_time_val(l_val if has_l else (j_val if has_j else (scans_list[0] if scans_list else '')))
                note = f"มีเวลาเข้า ({time_in_str})" if time_in_str else "มีเวลาเข้า"
                detail = f"ลืมสแกนออก ({note})"
                action = 'ติดตามขอใบแจ้งเวลาทำงานเพื่อยืนยันเวลาออก'
                anomalies.append({
                    'category': 'missing_punch',
                    'category_desc': 'ลืมสแกน (ขาเดียว)',
                    'sheet': sheet_name,
                    'date': dt.strftime('%d/%m/%Y'),
                    'date_obj': dt,
                    'emp_id': emp_id,
                    'emp_name': emp_name,
                    'scan_raw': str(scan_val or '-'),
                    'scans_list': scans_list,
                    'scans_str': time_in_str or str(scan_val or '-'),
                    'item': 'ลืมสแกนออก',
                    'note': note,
                    'detail': detail,
                    'action_needed': action
                })
                counts['missing_punch'] += 1
                counts['total'] += 1

            elif not in_confirmed and out_confirmed:
                time_out_str = format_time_val(m_val if has_m else (k_val if has_k else (scans_list[-1] if scans_list else '')))
                note = f"มีเวลาออก ({time_out_str})" if time_out_str else "มีเวลาออก"
                detail = f"ลืมสแกนเข้า ({note})"
                action = 'ติดตามขอใบแจ้งเวลาทำงานเพื่อยืนยันเวลาเข้า'
                anomalies.append({
                    'category': 'missing_punch',
                    'category_desc': 'ลืมสแกน (ขาเดียว)',
                    'sheet': sheet_name,
                    'date': dt.strftime('%d/%m/%Y'),
                    'date_obj': dt,
                    'emp_id': emp_id,
                    'emp_name': emp_name,
                    'scan_raw': str(scan_val or '-'),
                    'scans_list': scans_list,
                    'scans_str': time_out_str or str(scan_val or '-'),
                    'item': 'ลืมสแกนเข้า',
                    'note': note,
                    'detail': detail,
                    'action_needed': action
                })
                counts['missing_punch'] += 1
                counts['total'] += 1

            # 2. ไม่พบสแกนและไม่ลา (Unrecorded Absence) - ไม่มีสแกน และไม่มีการแจ้งในระบบ
            elif not in_confirmed and not out_confirmed and not is_holiday and not is_sunday:
                anomalies.append({
                    'category': 'unrecorded_absence',
                    'category_desc': 'ไม่พบสแกนและไม่ลา',
                    'sheet': sheet_name,
                    'date': dt.strftime('%d/%m/%Y'),
                    'date_obj': dt,
                    'emp_id': emp_id,
                    'emp_name': emp_name,
                    'scan_raw': '-',
                    'scans_list': [],
                    'scans_str': '-',
                    'item': 'ไม่พบสแกนและไม่ลา',
                    'note': 'ไม่พบสแกนนิ้วและไม่มีบันทึกการลา',
                    'detail': 'ไม่พบสแกนนิ้วในระบบ และไม่มีบันทึกข้อมูลการลาในตาราง',
                    'action_needed': 'ตรวจสอบว่าพนักงานขาดงาน หรือยังไม่ได้บันทึกใบลาที่ได้รับอนุมัติ'
                })
                counts['unrecorded_absence'] += 1
                counts['total'] += 1

    # จัดกลุ่มตามชื่อพนักงาน และเรียงตามลำดับวันที่โดยอัตโนมัติ
    def anomaly_sort_key(a):
        dt_o = a.get('date_obj')
        d_tup = (dt_o.year, dt_o.month, dt_o.day) if dt_o else (9999, 12, 31)
        return (clean_thai_name(a.get('emp_name', '')), str(a.get('emp_id', '')), d_tup)

    anomalies.sort(key=anomaly_sort_key)

    return {
        'success': True,
        'counts': counts,
        'anomalies': anomalies
    }


def populate_unrecorded_sheet_in_workbook(wb, anomalies=None, filename="", col_config=None, create_sheet=True):
    """
    กรอกข้อมูลพนักงานที่ไม่มีการแจ้งออนไลน์ลงในชีท 'ไม่มีแจ้งในระบบ' ของ openpyxl.Workbook
    จัดกลุ่มตามชื่อพนักงาน (หาก 1 คนมีหลายวัน จะเรียงติดกันตามลำดับวันที่)
    พร้อมฟอร์แมตตามแบบฟอร์มมาตรฐานของบริษัท (Yellow Banner A1, Blue Header A2:G2, Hidden Col C)
    """
    sheet_name = 'ไม่มีแจ้งในระบบ'
    if sheet_name not in wb.sheetnames:
        if not create_sheet:
            return False
        ws = wb.create_sheet(title=sheet_name)
    else:
        ws = wb[sheet_name]

    if anomalies is None:
        res = analyze_attendance_anomalies(wb, col_config=col_config)
        anomalies = res.get('anomalies', [])

    # 1. ปลดการ merge เซลล์เก่าในแถว 3 ขึ้นไป คอลัมน์ 1-7 เพื่อไม่ให้ติด MergedCell read-only
    ranges_to_unmerge = []
    for mrange in list(ws.merged_cells.ranges):
        if mrange.min_row >= 3 and mrange.min_col <= 7:
            ranges_to_unmerge.append(mrange)
    for r in ranges_to_unmerge:
        ws.unmerge_cells(str(r))

    # 2. หัวตารางแถว 1 (Yellow Banner)
    yellow_fill = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
    if not ws['A1'].value:
        m_proj = re.search(r'(?:CN\.?|cn\.?)?(\d{4,5})', str(filename))
        proj_code = f"CN.{m_proj.group(1)}" if m_proj else ""
        ws['A1'].value = f"รายชื่อที่ไม่มีการแจ้งออนไลน์ {proj_code}".strip()
    ws['A1'].font = Font(name="Arial", size=12, bold=True)

    for c_idx in range(1, 8):
        ws.cell(1, c_idx).fill = yellow_fill

    # F1:G1 merge & formula NOW()
    ws.cell(1, 6).value = "=NOW()"
    ws.cell(1, 6).font = Font(name="Arial", size=10, bold=False)
    ws.cell(1, 6).alignment = Alignment(horizontal="right", vertical="center")
    ws.cell(1, 6).number_format = 'm/d/yy h:mm'
    if 'F1:G1' not in [str(m) for m in ws.merged_cells.ranges]:
        ws.merge_cells('F1:G1')

    # 3. หัวตารางแถว 2 (Blue Header)
    blue_fill = PatternFill(start_color="FF0070C0", end_color="FF0070C0", fill_type="solid")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    headers = {
        1: ("รหัส", "center"),
        2: ("ชื่อ-สกุล", "left"),
        4: ("วันที่", "center"),
        5: ("รายการ", "center"),
        6: ("หมายเหตุ", "left"),
        7: ("ติดตามครั้งที่", "center")
    }
    for c_idx, (h_text, h_align) in headers.items():
        cell = ws.cell(2, c_idx)
        cell.value = h_text
        cell.fill = blue_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal=h_align, vertical="center")

    # ขนาดคอลัมน์
    ws.column_dimensions['A'].width = 8.5
    ws.column_dimensions['B'].width = 24.0
    ws.column_dimensions['C'].hidden = True
    ws.column_dimensions['C'].width = 0
    ws.column_dimensions['D'].width = 13.0
    ws.column_dimensions['E'].width = 22.0
    ws.column_dimensions['F'].width = 32.0
    ws.column_dimensions['G'].width = 15.0

    # 4. ล้างข้อมูลเก่าแถว 3 ขึ้นไป เฉพาะคอลัมน์ A ถึง G (ไม่แตะคอลัมน์โน้ต I เป็นต้นไป)
    max_clear = max(ws.max_row, len(anomalies) + 5)
    empty_border = Border()
    empty_fill = PatternFill(fill_type=None)
    for r in range(3, max_clear + 1):
        for c in range(1, 8):
            cell = ws.cell(r, c)
            cell.value = None
            cell.border = empty_border
            cell.fill = empty_fill

    # 5. เขียนรายการความผิดปกติ
    data_font = Font(name="Arial", size=10)
    data_border = Border(
        left=Side(style='thin', color='A6A6A6'),
        right=Side(style='thin', color='A6A6A6'),
        top=Side(style='hair', color='D9D9D9'),
        bottom=Side(style='hair', color='D9D9D9')
    )

    for idx, a in enumerate(anomalies, start=3):
        cat = a.get('category', '')
        detail = a.get('detail', '')
        scans_str = a.get('scans_str', '')

        item_name = a.get('item')
        note = a.get('note')

        if not item_name or not note:
            if cat == 'missing_punch':
                if 'ลืมสแกนเข้า' in detail:
                    item_name = item_name or 'ลืมสแกนเข้า'
                    note = note or (f"มีเวลาออก ({scans_str})" if scans_str and scans_str != '-' else "มีเวลาออก")
                else:
                    item_name = item_name or 'ลืมสแกนออก'
                    note = note or (f"มีเวลาเข้า ({scans_str})" if scans_str and scans_str != '-' else "มีเวลาเข้า")
            else:
                item_name = item_name or 'ไม่พบสแกนและไม่ลา'
                note = note or 'ไม่พบสแกนนิ้วและไม่มีบันทึกการลา'

        emp_id_val = a.get('emp_id', '')
        val_id = int(emp_id_val) if str(emp_id_val).isdigit() else str(emp_id_val)

        row_vals = {
            1: (val_id, "center"),
            2: (str(a.get('emp_name', '')), "left"),
            3: ("", "center"),
            4: (str(a.get('date', '')), "center"),
            5: (item_name, "center"),
            6: (note, "left"),
            7: ("", "center")
        }

        for c_idx, (val, align) in row_vals.items():
            cell = ws.cell(idx, c_idx)
            cell.value = val
            cell.font = data_font
            cell.alignment = Alignment(horizontal=align, vertical="center")
            cell.border = data_border

    # Auto Filter
    if anomalies:
        ws.auto_filter.ref = f"A2:G{2 + len(anomalies)}"

    return True


def fill_unrecorded_sheet(target, anomalies=None, col_config=None, save_path=None, create_sheet=True):
    """
    กรอกข้อมูลพนักงานที่ไม่มีการแจ้งออนไลน์ลงในชีท 'ไม่มีแจ้งในระบบ'
    target: สามารถเป็น filepath (str) หรือ openpyxl.Workbook
    """
    if isinstance(target, str):
        if not os.path.exists(target):
            return {'success': False, 'error': 'ไม่พบไฟล์รายงาน STAFF TIME'}
        wb = openpyxl.load_workbook(target)
        populate_unrecorded_sheet_in_workbook(
            wb,
            anomalies=anomalies,
            filename=os.path.basename(target),
            col_config=col_config,
            create_sheet=create_sheet
        )
        out_file = save_path or target
        wb.save(out_file)
        wb.close()
        return {'success': True, 'path': out_file, 'count': len(anomalies) if anomalies is not None else 0}
    elif isinstance(target, openpyxl.Workbook):
        populate_unrecorded_sheet_in_workbook(
            target,
            anomalies=anomalies,
            col_config=col_config,
            create_sheet=create_sheet
        )
        return {'success': True, 'workbook': target, 'count': len(anomalies) if anomalies is not None else 0}
    return {'success': False, 'error': 'target ไม่ถูกต้อง'}


class AnomalyDialog(ctk.CTkToplevel):
    """หน้าต่างแสดงผลการวิเคราะห์เวลาทำงานและสแกนผิดปกติอัจฉริยะ"""
    def __init__(self, parent, main_file, col_config=None, icon_path=None):
        super().__init__(parent)
        self.parent = parent
        self.main_file = main_file
        self.col_config = col_config
        self.font_family = "Tahoma"

        self.title("ระบบวิเคราะห์เวลาทำงานและสแกนผิดปกติอัจฉริยะ (Smart Attendance Anomaly Analyzer)")
        center_window(self, 1020, 740, parent)
        self.minsize(900, 620)

        if icon_path and os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        self.active_filter = "all"
        self.search_query = ""
        self.anomalies_data = []

        self.setup_ui()
        self.load_data()

        self.grab_set()
        self.focus_force()

    def setup_ui(self):
        # 1. Header
        header = ctk.CTkFrame(self, fg_color="#0f172a", corner_radius=0, height=80)
        header.pack(fill="x")

        lbl_t = ctk.CTkLabel(
            header,
            text="ระบบวิเคราะห์เวลาทำงานและสแกนผิดปกติอัจฉริยะ (Smart Anomaly Analyzer)",
            font=ctk.CTkFont(family=self.font_family, size=20, weight="bold"),
            text_color="#ffffff"
        )
        lbl_t.pack(pady=(12, 2))

        lbl_sub = ctk.CTkLabel(
            header,
            text="ตรวจจับสแกนขาเดียว, ลืมสแกนเข้า/ออก, และไม่พบการลงเวลาในวันทำงาน สำหรับทุกโครงการของ CNT",
            font=ctk.CTkFont(family=self.font_family, size=13),
            text_color="#cbd5e1"
        )
        lbl_sub.pack(pady=(0, 12))

        # 2. Summary KPI Cards
        self.kpi_frame = ctk.CTkFrame(self, fg_color="#f8fafc", corner_radius=0)
        self.kpi_frame.pack(fill="x", padx=16, pady=(12, 4))
        for i in range(4):
            self.kpi_frame.grid_columnconfigure(i, weight=1)

        self.lbl_kpi_total = self._create_kpi_card(self.kpi_frame, 0, "รายการผิดปกติทั้งหมด", "-", "#b45309", "#fffbeb")
        self.lbl_kpi_punch = self._create_kpi_card(self.kpi_frame, 1, "ลืมสแกน (ขาเดียว)", "-", "#b91c1c", "#fef2f2")
        self.lbl_kpi_absent = self._create_kpi_card(self.kpi_frame, 2, "ไม่พบสแกนและไม่ลา", "-", "#c2410c", "#fff7ed")
        self.lbl_kpi_sheets = self._create_kpi_card(self.kpi_frame, 3, "จำนวนวันที่ตรวจพบ", "-", "#0369a1", "#f0f9ff")

        # 2.5 Visual Attendance Distribution Gauge (UI/UX Pro Max)
        self.gauge_frame = ctk.CTkFrame(self, fg_color="#ffffff", corner_radius=10, border_width=1, border_color="#cbd5e1")
        self.gauge_frame.pack(fill="x", padx=16, pady=(4, 6))

        gf_top = ctk.CTkFrame(self.gauge_frame, fg_color="transparent")
        gf_top.pack(fill="x", padx=12, pady=(8, 4))
        lbl_g_title = ctk.CTkLabel(gf_top, text="📊 สัดส่วนความสมบูรณ์ของเวลาทำงาน (Attendance Health Gauge)", font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"), text_color="#0f172a")
        lbl_g_title.pack(side="left")
        self.lbl_gauge_pct = ctk.CTkLabel(gf_top, text="กำลังวิเคราะห์ข้อมูล...", font=ctk.CTkFont(family=self.font_family, size=11, weight="bold"), text_color="#475569")
        self.lbl_gauge_pct.pack(side="right")

        self.bar_bg = ctk.CTkFrame(self.gauge_frame, fg_color="#e2e8f0", height=14, corner_radius=7)
        self.bar_bg.pack(fill="x", padx=12, pady=(0, 8))
        self.bar_bg.pack_propagate(False)

        # 3. Filter & Controls Bar
        ctrl_frame = ctk.CTkFrame(self, fg_color="transparent")
        ctrl_frame.pack(fill="x", padx=16, pady=8)
        ctrl_frame.grid_columnconfigure(1, weight=1)

        filter_box = ctk.CTkFrame(ctrl_frame, fg_color="transparent")
        filter_box.grid(row=0, column=0, sticky="w")

        self.btn_f_all = ctk.CTkButton(
            filter_box, text="ทั้งหมด", width=80, height=34, corner_radius=8,
            font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"),
            fg_color="#1d4ed8", hover_color="#1e40af", command=lambda: self.set_filter("all")
        )
        self.btn_f_all.pack(side="left", padx=(0, 4))

        self.btn_f_punch = ctk.CTkButton(
            filter_box, text="ลืมสแกน (ขาเดียว)", width=125, height=34, corner_radius=8,
            font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"),
            fg_color="#f1f5f9", hover_color="#e2e8f0", text_color="#334155", command=lambda: self.set_filter("missing_punch")
        )
        self.btn_f_punch.pack(side="left", padx=4)

        self.btn_f_absent = ctk.CTkButton(
            filter_box, text="ไม่พบสแกน/ไม่ลา", width=125, height=34, corner_radius=8,
            font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"),
            fg_color="#f1f5f9", hover_color="#e2e8f0", text_color="#334155", command=lambda: self.set_filter("unrecorded_absence")
        )
        self.btn_f_absent.pack(side="left", padx=4)

        # Search box
        search_box = ctk.CTkFrame(ctrl_frame, fg_color="transparent")
        search_box.grid(row=0, column=1, sticky="ew", padx=10)
        search_box.grid_columnconfigure(0, weight=1)

        self.txt_search = ctk.CTkEntry(
            search_box,
            placeholder_text="พิมพ์รหัสพนักงาน หรือชื่อพนักงานเพื่อกรอง...",
            font=ctk.CTkFont(family=self.font_family, size=13),
            height=36,
            corner_radius=8
        )
        self.txt_search.grid(row=0, column=0, sticky="ew")
        self.txt_search.bind("<KeyRelease>", lambda e: self.on_search_changed())

        # Action Buttons Container (Right side)
        action_box = ctk.CTkFrame(ctrl_frame, fg_color="transparent")
        action_box.grid(row=0, column=2, sticky="e")

        self.btn_fill_sheet = ctk.CTkButton(
            action_box,
            text="📝 กรอกลงชีท \"ไม่มีแจ้งในระบบ\"",
            font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"),
            fg_color="#1d4ed8",
            hover_color="#1e40af",
            height=36,
            corner_radius=8,
            command=self.fill_unrecorded_direct
        )
        self.btn_fill_sheet.pack(side="left", padx=(0, 6))

        self.btn_export = ctk.CTkButton(
            action_box,
            text="📥 ส่งออก Excel ฟอร์มบริษัท",
            font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"),
            fg_color="#059669",
            hover_color="#047857",
            height=36,
            corner_radius=8,
            command=self.export_to_excel
        )
        self.btn_export.pack(side="left")

        # 4. Results List Container
        self.list_scroll = ctk.CTkScrollableFrame(self, fg_color="#f1f5f9", corner_radius=10)
        self.list_scroll.pack(fill="both", expand=True, padx=16, pady=(4, 10))
        self.list_scroll.grid_columnconfigure(0, weight=1)

        # 5. Bottom Status Bar
        bot_bar = ctk.CTkFrame(self, fg_color="transparent")
        bot_bar.pack(fill="x", padx=16, pady=(0, 10))

        self.lbl_status = ctk.CTkLabel(
            bot_bar,
            text="กำลังโหลดข้อมูลการวิเคราะห์...",
            font=ctk.CTkFont(family=self.font_family, size=13),
            text_color="#475569"
        )
        self.lbl_status.pack(side="left")

        btn_close = ctk.CTkButton(
            bot_bar,
            text="ปิดหน้าต่าง",
            font=ctk.CTkFont(family=self.font_family, size=13, weight="bold"),
            fg_color="#334155",
            hover_color="#1e293b",
            height=36,
            width=110,
            corner_radius=8,
            command=self.destroy
        )
        btn_close.pack(side="right")

        # ผูก Tooltip (UI/UX Pro Max)
        add_tooltip(self.btn_f_all, "แสดงรายการความผิดปกติทั้งหมดที่ตรวจพบ")
        add_tooltip(self.btn_f_punch, "กรองเฉพาะพนักงานที่มีสแกนนิ้วเพียงขาเดียว (ขาดเช้าหรือขาดเย็น)")
        add_tooltip(self.btn_f_absent, "กรองเฉพาะพนักงานที่ไม่มีสแกนนิ้วเลย และไม่ได้ลงวันลา")
        add_tooltip(self.btn_fill_sheet, "กรอกรายชื่อผู้ไม่มีการแจ้งออนไลน์ลงในชีท 'ไม่มีแจ้งในระบบ' ของไฟล์ STAFF TIME ทันที")
        add_tooltip(self.btn_export, "ส่งออกรายงานเป็นไฟล์ Excel ใหม่ตามแบบฟอร์มมาตรฐานของบริษัท")
        add_tooltip(btn_close, "ปิดหน้าต่างวิเคราะห์เวลาผิดปกติ")

    def _create_kpi_card(self, parent, col, title, initial_val, text_col, bg_col):
        card = ctk.CTkFrame(parent, fg_color=bg_col, corner_radius=10, border_width=1, border_color="#cbd5e1")
        card.grid(row=0, column=col, padx=5, pady=4, sticky="nsew")
        lbl_v = ctk.CTkLabel(card, text=initial_val, font=ctk.CTkFont(family=self.font_family, size=22, weight="bold"), text_color=text_col)
        lbl_v.pack(pady=(8, 2))
        lbl_t = ctk.CTkLabel(card, text=title, font=ctk.CTkFont(family=self.font_family, size=12), text_color="#334155")
        lbl_t.pack(pady=(0, 8))
        return lbl_v

    def load_data(self):
        res = analyze_attendance_anomalies(self.main_file, self.col_config)
        if not res.get('success'):
            messagebox.showerror("เกิดข้อผิดพลาด", res.get('error', 'ไม่สามารถวิเคราะห์ข้อมูลได้'))
            self.destroy()
            return

        counts = res.get('counts', {})
        self.anomalies_data = res.get('anomalies', [])

        total = counts.get('total', 0)
        missing_punch = counts.get('missing_punch', 0)
        unrecorded = counts.get('unrecorded_absence', 0)

        self.lbl_kpi_total.configure(text=f"{total:,} จุด")
        self.lbl_kpi_punch.configure(text=f"{missing_punch:,} คน")
        self.lbl_kpi_absent.configure(text=f"{unrecorded:,} จุด")

        # อัปเดตแถบสัดส่วนความสมบูรณ์ของเวลาทำงาน (Visual Attendance Distribution Gauge)
        for c in self.bar_bg.winfo_children():
            try:
                c.destroy()
            except Exception:
                pass
        if total > 0:
            pct_punch = (missing_punch / total) * 100
            pct_absent = (unrecorded / total) * 100
            self.lbl_gauge_pct.configure(
                text=f"ลืมสแกนขาเดียว: {missing_punch} จุด ({pct_punch:.1f}%)  |  ไม่พบสแกน/ไม่ลา: {unrecorded} จุด ({pct_absent:.1f}%)",
                text_color="#0f172a"
            )
            w_punch = missing_punch / total
            w_absent = unrecorded / total
            if w_punch > 0:
                f_p = ctk.CTkFrame(self.bar_bg, fg_color="#dc2626", corner_radius=0, height=14)
                f_p.place(relx=0, rely=0, relwidth=w_punch, relheight=1.0)
            if w_absent > 0:
                f_a = ctk.CTkFrame(self.bar_bg, fg_color="#ea580c", corner_radius=0, height=14)
                f_a.place(relx=w_punch, rely=0, relwidth=w_absent, relheight=1.0)
        else:
            self.lbl_gauge_pct.configure(
                text="ข้อมูลเวลาทำงานถูกต้องสมบูรณ์ 100% (ไม่พบจุดผิดปกติ)",
                text_color="#15803d"
            )
            f_ok = ctk.CTkFrame(self.bar_bg, fg_color="#16a34a", corner_radius=7, height=14)
            f_ok.place(relx=0, rely=0, relwidth=1.0, relheight=1.0)

        # Distinct sheets
        distinct_sheets = len(set(a['sheet'] for a in self.anomalies_data))
        self.lbl_kpi_sheets.configure(text=f"{distinct_sheets} ชีท")

        self.render_list()

    def set_filter(self, filter_name):
        self.active_filter = filter_name
        self.btn_f_all.configure(fg_color="#1d4ed8" if filter_name == "all" else "#f1f5f9", text_color="#ffffff" if filter_name == "all" else "#334155")
        self.btn_f_punch.configure(fg_color="#1d4ed8" if filter_name == "missing_punch" else "#f1f5f9", text_color="#ffffff" if filter_name == "missing_punch" else "#334155")
        self.btn_f_absent.configure(fg_color="#1d4ed8" if filter_name == "unrecorded_absence" else "#f1f5f9", text_color="#ffffff" if filter_name == "unrecorded_absence" else "#334155")
        self.render_list()

    def on_search_changed(self):
        self.search_query = self.txt_search.get().strip().lower()
        self.render_list()

    def render_list(self):
        for child in self.list_scroll.winfo_children():
            child.destroy()

        filtered = []
        for a in self.anomalies_data:
            if self.active_filter != "all" and a['category'] != self.active_filter:
                continue
            if self.search_query:
                q = self.search_query
                match = (q in a['emp_id'].lower() or q in a['emp_name'].lower() or q in a['sheet'].lower())
                if not match:
                    continue
            filtered.append(a)

        self.lbl_status.configure(text=f"แสดง {len(filtered):,} รายการ จากทั้งหมด {len(self.anomalies_data):,} จุดผิดปกติ")

        if not filtered:
            empty_card = ctk.CTkFrame(self.list_scroll, fg_color="#ffffff", corner_radius=10, border_width=1, border_color="#cbd5e1")
            empty_card.pack(fill="x", pady=20, padx=6)
            lbl_empty = ctk.CTkLabel(
                empty_card,
                text="ไม่พบข้อมูลความผิดปกติ หรือข้อมูลการลงเวลาถูกต้องครบถ้วนสมบูรณ์ 100%",
                font=ctk.CTkFont(family=self.font_family, size=14, weight="bold"),
                text_color="#16a34a"
            )
            lbl_empty.pack(pady=30)
            return

        for idx, item in enumerate(filtered):
            card = ctk.CTkFrame(self.list_scroll, fg_color="#ffffff", corner_radius=10, border_width=1, border_color="#cbd5e1")
            card.pack(fill="x", pady=4, padx=4)

            top_row = ctk.CTkFrame(card, fg_color="transparent")
            top_row.pack(fill="x", padx=14, pady=(10, 4))
            top_row.grid_columnconfigure(0, weight=1)

            lbl_emp = ctk.CTkLabel(
                top_row,
                text=f"[{idx+1:03d}] วันที่ {item['date']} (ชีท {item['sheet']})  |  รหัส: {item['emp_id']}  -  {item['emp_name']}",
                font=ctk.CTkFont(family=self.font_family, size=13, weight="bold"),
                text_color="#0f172a"
            )
            lbl_emp.grid(row=0, column=0, sticky="w")

            is_punch = (item['category'] == 'missing_punch')
            badge_bg = "#fef2f2" if is_punch else "#fff7ed"
            badge_fg = "#b91c1c" if is_punch else "#c2410c"
            badge = ctk.CTkLabel(
                top_row,
                text=f" {item['category_desc']} ",
                font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"),
                fg_color=badge_bg,
                text_color=badge_fg,
                corner_radius=6
            )
            badge.grid(row=0, column=1, sticky="e")

            mid_row = ctk.CTkFrame(card, fg_color="transparent")
            mid_row.pack(fill="x", padx=14, pady=(0, 4))
            lbl_detail = ctk.CTkLabel(
                mid_row,
                text=f"• สแกนดิบในชีท: [{item['scan_raw']}]  |  รายละเอียด: {item['detail']}",
                font=ctk.CTkFont(family=self.font_family, size=12),
                text_color="#334155"
            )
            lbl_detail.pack(anchor="w")

            bot_row = ctk.CTkFrame(card, fg_color="transparent")
            bot_row.pack(fill="x", padx=14, pady=(0, 10))
            lbl_act = ctk.CTkLabel(
                bot_row,
                text=f"• ข้อเสนอแนะ: {item['action_needed']}",
                font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"),
                text_color="#0369a1"
            )
            lbl_act.pack(anchor="w")

    def fill_unrecorded_direct(self):
        """กรอกข้อมูลลงในชีท 'ไม่มีแจ้งในระบบ' ของไฟล์ STAFF TIME ทันที"""
        if not self.main_file or not os.path.exists(self.main_file):
            messagebox.showerror("ข้อผิดพลาด", "ไม่พบไฟล์รายงาน STAFF TIME สำหรับดำเนินการ", parent=self)
            return

        if not self.anomalies_data:
            messagebox.showinfo("ไม่มีข้อมูล", "ไม่พบข้อมูลความผิดปกติ หรือข้อมูลการลงเวลาถูกต้องครบถ้วนสมบูรณ์ 100%", parent=self)
            return

        confirm = messagebox.askyesno(
            "ยืนยันการกรอกข้อมูล",
            f"ระบบจะกรอกรายชื่อผู้ไม่มีการแจ้งออนไลน์ ({len(self.anomalies_data):,} รายการ) ลงในชีท 'ไม่มีแจ้งในระบบ' ของไฟล์:\n{os.path.basename(self.main_file)}\n\n(หาก 1 คนมีหลายวัน ระบบจัดเรียงชื่อให้อยู่ติดกันตามลำดับวันที่อัตโนมัติ)\n\nต้องการดำเนินการต่อหรือไม่?",
            parent=self
        )
        if not confirm:
            return

        try:
            res = fill_unrecorded_sheet(self.main_file, anomalies=self.anomalies_data, col_config=self.col_config)
            if res.get('success'):
                messagebox.showinfo(
                    "ดำเนินการสำเร็จ",
                    f"กรอกข้อมูลลงชีท 'ไม่มีแจ้งในระบบ' เรียบร้อยแล้ว!\n\n• จำนวนที่บันทึก: {len(self.anomalies_data):,} รายการ\n• บันทึกลงไฟล์: {os.path.basename(self.main_file)}\n• พร้อมให้ฝ่ายบุคคล/ผู้เกี่ยวข้องติดตามผลได้ทันที",
                    parent=self
                )
            else:
                messagebox.showerror("เกิดข้อผิดพลาด", res.get('error', 'ไม่สามารถกรอกข้อมูลได้'), parent=self)
        except PermissionError:
            messagebox.showerror(
                "ไฟล์กำลังถูกเปิดใช้งาน (Permission Denied)",
                f"ไฟล์ {os.path.basename(self.main_file)} กำลังถูกเปิดอยู่ในโปรแกรม Excel\n\nกรุณาปิดโปรแกรม Excel ก่อน แล้วลองกดปุ่มนี้ใหม่อีกครั้ง",
                parent=self
            )
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกข้อมูลลงไฟล์ได้: {str(e)}", parent=self)

    def export_to_excel(self):
        if not self.anomalies_data:
            messagebox.showinfo("ไม่มีข้อมูล", "ไม่พบข้อมูลความผิดปกติที่จะส่งออก", parent=self)
            return

        default_name = f"รายชื่อที่ไม่มีการแจ้งออนไลน์_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        save_path = filedialog.asksaveasfilename(
            parent=self,
            title="บันทึกรายงานความผิดปกติเป็นไฟล์ Excel (ฟอร์มบริษัท)",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not save_path:
            return

        try:
            wb = openpyxl.Workbook()
            populate_unrecorded_sheet_in_workbook(
                wb,
                anomalies=self.anomalies_data,
                filename=os.path.basename(self.main_file or ''),
                create_sheet=True
            )
            # ลบ Sheet เริ่มต้นหากมี sheet 'ไม่มีแจ้งในระบบ' อยู่
            if 'Sheet' in wb.sheetnames and len(wb.sheetnames) > 1 and 'ไม่มีแจ้งในระบบ' in wb.sheetnames:
                del wb['Sheet']

            wb.save(save_path)
            wb.close()
            messagebox.showinfo(
                "ส่งออกสำเร็จ",
                f"สร้างไฟล์รายงานความผิดปกติตามแบบฟอร์มบริษัทเรียบร้อยแล้ว:\n{save_path}\n\n• จัดกลุ่มพนักงานที่มีหลายวันเรียงต่อกันตามลำดับวันที่",
                parent=self
            )
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกไฟล์ Excel ได้: {str(e)}", parent=self)

# -------------------------------------------------------------
# ระบบอัปเดตโปรแกรมอัตโนมัติผ่าน GitHub Releases (Auto-Updater)
# -------------------------------------------------------------

def parse_version_tuple(v_str):
    """แปลงข้อความเวอร์ชัน เช่น 'v1.2.3' เป็น tuple ตัวเลข (1, 2, 3)"""
    nums = re.findall(r'\d+', str(v_str))
    if not nums:
        return (0, 0, 0)
    return tuple(int(x) for x in nums)


def is_newer_version(remote_ver, local_ver):
    """ส่งคืน True หาก remote_ver ใหม่กว่า local_ver"""
    r_parts = parse_version_tuple(remote_ver)
    l_parts = parse_version_tuple(local_ver)
    max_len = max(len(r_parts), len(l_parts))
    r_padded = r_parts + (0,) * (max_len - len(r_parts))
    l_padded = l_parts + (0,) * (max_len - len(l_parts))
    return r_padded > l_padded


def check_github_update(repo_name, current_version=APP_VERSION):
    """
    ตรวจสอบเวอร์ชันล่าสุดจาก GitHub Releases API
    Returns dict สรุปข้อมูลเวอร์ชัน, ลิงก์ดาวน์โหลด และ changelog
    """
    repo = (repo_name or "").strip().strip("/")
    if not repo or "/" not in repo:
        return {
            'success': False,
            'error_msg': f"รูปแบบชื่อ GitHub Repository ไม่ถูกต้อง ('{repo}') ต้องเป็น 'owner/repo' เช่น 'DietrichGebert/LeaveAutomation-1856'"
        }

    url = f"https://api.github.com/repos/{repo}/releases/latest"
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": f"LeaveAutomation-Updater/{current_version}",
            "Accept": "application/vnd.github.v3+json"
        }
    )
    try:
        with urllib.request.urlopen(req, timeout=12) as response:
            raw_data = response.read().decode("utf-8")
            data = json.loads(raw_data)
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return {
                'success': False,
                'error_msg': f"ไม่พบ Release บน GitHub สำหรับ '{repo}' (404 Not Found)\nโปรดตรวจสอบว่าได้สร้าง Release และแนบไฟล์บน GitHub แล้ว หรือระบุชื่อ Repo ถูกต้อง"
            }
        elif e.code == 403:
            return {
                'success': False,
                'error_msg': "GitHub API ติดการจำกัดจำนวนครั้ง (Rate Limit) กรุณารอสักครู่แล้วลองใหม่"
            }
        else:
            return {
                'success': False,
                'error_msg': f"GitHub API ส่งรหัสข้อผิดพลาด HTTP {e.code}: {e.reason}"
            }
    except Exception as e:
        return {
            'success': False,
            'error_msg': f"ไม่สามารถเชื่อมต่ออินเทอร์เน็ตหรือ GitHub ได้ ({str(e)})"
        }

    tag_name = data.get("tag_name", "").strip()
    release_name = data.get("name", "") or tag_name
    body = data.get("body", "") or "ไม่มีบันทึกการเปลี่ยนแปลง (No changelog provided)"
    published_at = data.get("published_at", "")

    # ตรวจสอบว่าเวอร์ชันใหม่กว่าหรือไม่
    update_avail = is_newer_version(tag_name, current_version)

    # ค้นหาไฟล์ Asset (.exe ก่อน, ตามด้วย .zip)
    assets = data.get("assets", [])
    target_asset = None
    for a in assets:
        name = a.get("name", "").lower()
        if name.endswith(".exe"):
            target_asset = a
            break

    if not target_asset:
        for a in assets:
            name = a.get("name", "").lower()
            if name.endswith(".zip"):
                target_asset = a
                break

    asset_url = target_asset.get("browser_download_url", "") if target_asset else ""
    asset_name = target_asset.get("name", "") if target_asset else ""
    asset_size = target_asset.get("size", 0) if target_asset else 0

    return {
        'success': True,
        'tag_name': tag_name,
        'release_name': release_name,
        'published_at': published_at,
        'body': body,
        'update_available': update_avail,
        'asset_url': asset_url,
        'asset_name': asset_name,
        'asset_size': asset_size,
        'error_msg': ""
    }


class UpdateDialog(ctk.CTkToplevel):
    def __init__(self, parent, current_version=APP_VERSION, github_repo=DEFAULT_GITHUB_REPO, icon_path=None, on_repo_save=None):
        super().__init__(parent)
        self.parent = parent
        self.current_version = current_version
        self.github_repo = github_repo
        self.icon_path = icon_path
        self.on_repo_save = on_repo_save
        self.font_family = "Tahoma"

        self.release_data = None
        self.is_downloading = False
        self.is_cancelled = False

        self.title("ระบบตรวจสอบและอัปเดตโปรแกรม (Auto-Updater)")
        center_window(self, 720, 640, parent)
        self.minsize(680, 580)

        if icon_path and os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        self.setup_ui()
        self.lift()
        self.attributes("-topmost", True)
        self.after_idle(lambda: self.attributes("-topmost", False))
        self.grab_set()
        self.focus_force()

        # เริ่มตรวจสอบเวอร์ชันอัตโนมัติเมื่อเปิดหน้าต่าง
        self.after(300, self.start_check_update)

    def setup_ui(self):
        # แถบหัวเรื่องสีฟ้า Google Blue
        header = ctk.CTkFrame(self, fg_color="#0f172a", corner_radius=0, height=80)
        header.pack(fill="x")

        lbl_t = ctk.CTkLabel(
            header,
            text="ระบบอัปเดตโปรแกรมอัตโนมัติ (Live Auto-Updater)",
            font=ctk.CTkFont(family=self.font_family, size=20, weight="bold"),
            text_color="#ffffff"
        )
        lbl_t.pack(pady=(12, 2))

        lbl_sub = ctk.CTkLabel(
            header,
            text="ตรวจสอบและติดตั้งเวอร์ชันใหม่ล่าสุดจาก GitHub Releases (OHM282/LeaveAutomation-CNT)",
            font=ctk.CTkFont(family=self.font_family, size=13),
            text_color="#cbd5e1"
        )
        lbl_sub.pack(pady=(0, 12))

        scroll = ctk.CTkScrollableFrame(self, fg_color="#f8fafc")
        scroll.pack(fill="both", expand=True, padx=16, pady=10)

        # -------------------------------------------------------------
        # CARD 1: ข้อมูลเวอร์ชันและคลังโค้ด GitHub
        # -------------------------------------------------------------
        c1 = ctk.CTkFrame(scroll, fg_color="#ffffff", corner_radius=10, border_width=1, border_color="#cbd5e1")
        c1.pack(fill="x", pady=(0, 12))

        lbl_c1_t = ctk.CTkLabel(
            c1,
            text="ข้อมูลเวอร์ชันและคลังโค้ด (GitHub Repository)",
            font=ctk.CTkFont(family=self.font_family, size=15, weight="bold"),
            text_color="#1e293b"
        )
        lbl_c1_t.pack(anchor="w", padx=16, pady=(12, 6))

        # แถวเวอร์ชันปัจจุบัน
        row_ver = ctk.CTkFrame(c1, fg_color="transparent")
        row_ver.pack(fill="x", padx=16, pady=3)
        lbl_v1 = ctk.CTkLabel(row_ver, text="เวอร์ชันปัจจุบันในเครื่อง:", font=ctk.CTkFont(family=self.font_family, size=13), text_color="#475569")
        lbl_v1.pack(side="left")
        badge_v1 = ctk.CTkLabel(
            row_ver,
            text=f" v{self.current_version} ",
            font=ctk.CTkFont(family=self.font_family, size=13, weight="bold"),
            fg_color="#e2e8f0",
            text_color="#0f172a",
            corner_radius=6
        )
        badge_v1.pack(side="left", padx=8)

        # แถว GitHub Repo
        row_repo = ctk.CTkFrame(c1, fg_color="transparent")
        row_repo.pack(fill="x", padx=16, pady=6)
        lbl_rp = ctk.CTkLabel(row_repo, text="GitHub Repo:", font=ctk.CTkFont(family=self.font_family, size=13, weight="bold"), text_color="#334155")
        lbl_rp.pack(side="left")
        self.txt_repo = ctk.CTkEntry(
            row_repo,
            width=280,
            height=32,
            font=ctk.CTkFont(family=self.font_family, size=13)
        )
        self.txt_repo.insert(0, self.github_repo)
        self.txt_repo.pack(side="left", padx=8)

        btn_save_rp = ctk.CTkButton(
            row_repo,
            text="บันทึกและตรวจสอบ",
            font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"),
            fg_color="#0284c7",
            hover_color="#0369a1",
            text_color="#ffffff",
            width=130,
            height=32,
            corner_radius=6,
            command=self.on_save_repo_clicked
        )
        btn_save_rp.pack(side="left")

        # แถวแสดงสถานะการตรวจสอบ
        row_st = ctk.CTkFrame(c1, fg_color="#f1f5f9", corner_radius=8)
        row_st.pack(fill="x", padx=16, pady=(6, 12))
        self.lbl_status = ctk.CTkLabel(
            row_st,
            text="⏳ กำลังเตรียมตรวจสอบเวอร์ชันล่าสุด...",
            font=ctk.CTkFont(family=self.font_family, size=13, weight="bold"),
            text_color="#0284c7",
            wraplength=620,
            justify="left"
        )
        self.lbl_status.pack(anchor="w", padx=12, pady=8)

        # -------------------------------------------------------------
        # CARD 2: ข้อมูลการอัปเดตและบันทึกการเปลี่ยนแปลง (Changelog)
        # -------------------------------------------------------------
        c2 = ctk.CTkFrame(scroll, fg_color="#ffffff", corner_radius=10, border_width=1, border_color="#cbd5e1")
        c2.pack(fill="x", pady=(0, 12))

        lbl_c2_t = ctk.CTkLabel(
            c2,
            text="รายละเอียดเวอร์ชันล่าสุดและบันทึกการเปลี่ยนแปลง (What's New)",
            font=ctk.CTkFont(family=self.font_family, size=15, weight="bold"),
            text_color="#1e293b"
        )
        lbl_c2_t.pack(anchor="w", padx=16, pady=(12, 4))

        self.lbl_release_title = ctk.CTkLabel(
            c2,
            text="ชื่อเวอร์ชัน: -",
            font=ctk.CTkFont(family=self.font_family, size=14, weight="bold"),
            text_color="#0f172a"
        )
        self.lbl_release_title.pack(anchor="w", padx=16, pady=2)

        self.lbl_release_meta = ctk.CTkLabel(
            c2,
            text="วันที่เผยแพร่: -  |  ขนาดไฟล์: -",
            font=ctk.CTkFont(family=self.font_family, size=12),
            text_color="#64748b"
        )
        self.lbl_release_meta.pack(anchor="w", padx=16, pady=(0, 6))

        lbl_ch_title = ctk.CTkLabel(c2, text="บันทึกการเปลี่ยนแปลง (Release Notes):", font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"), text_color="#334155")
        lbl_ch_title.pack(anchor="w", padx=16, pady=(4, 2))

        self.txt_changelog = ctk.CTkTextbox(
            c2,
            height=130,
            font=ctk.CTkFont(family=self.font_family, size=12),
            wrap="word",
            border_width=1,
            border_color="#cbd5e1",
            fg_color="#f8fafc"
        )
        self.txt_changelog.pack(fill="x", padx=16, pady=(0, 12))

        # -------------------------------------------------------------
        # CARD 3: ความคืบหน้าการดาวน์โหลดและติดตั้ง
        # -------------------------------------------------------------
        c3 = ctk.CTkFrame(scroll, fg_color="#ffffff", corner_radius=10, border_width=1, border_color="#cbd5e1")
        c3.pack(fill="x", pady=(0, 12))

        lbl_c3_t = ctk.CTkLabel(
            c3,
            text="ความคืบหน้าการดาวน์โหลดและติดตั้ง (Progress)",
            font=ctk.CTkFont(family=self.font_family, size=15, weight="bold"),
            text_color="#1e293b"
        )
        lbl_c3_t.pack(anchor="w", padx=16, pady=(12, 6))

        self.prog_bar = ctk.CTkProgressBar(c3, height=16, corner_radius=8, progress_color="#16a34a")
        self.prog_bar.pack(fill="x", padx=16, pady=4)
        self.prog_bar.set(0.0)

        self.lbl_progress = ctk.CTkLabel(
            c3,
            text="พร้อมดาวน์โหลด",
            font=ctk.CTkFont(family=self.font_family, size=13),
            text_color="#475569"
        )
        self.lbl_progress.pack(anchor="w", padx=16, pady=(2, 12))

        # -------------------------------------------------------------
        # แถบปุ่มกดด้านล่าง (Action Buttons)
        # -------------------------------------------------------------
        bar_bottom = ctk.CTkFrame(self, fg_color="#ffffff", height=65, border_width=1, border_color="#e2e8f0")
        bar_bottom.pack(fill="x", side="bottom")

        btn_box = ctk.CTkFrame(bar_bottom, fg_color="transparent")
        btn_box.pack(side="right", padx=16, pady=12)

        self.btn_download = ctk.CTkButton(
            btn_box,
            text="⚡ อัปเดตทันที (ดาวน์โหลด & รีสตาร์ท)",
            font=ctk.CTkFont(family=self.font_family, size=14, weight="bold"),
            fg_color="#16a34a",
            hover_color="#15803d",
            text_color="#ffffff",
            height=42,
            corner_radius=8,
            state="disabled",
            command=self.start_download_and_install
        )
        self.btn_download.pack(side="left", padx=6)

        self.btn_check = ctk.CTkButton(
            btn_box,
            text="ตรวจสอบอีกครั้ง",
            font=ctk.CTkFont(family=self.font_family, size=13, weight="bold"),
            fg_color="#0284c7",
            hover_color="#0369a1",
            text_color="#ffffff",
            height=42,
            corner_radius=8,
            command=self.start_check_update
        )
        self.btn_check.pack(side="left", padx=6)

        self.btn_close = ctk.CTkButton(
            btn_box,
            text="ปิดหน้าต่าง",
            font=ctk.CTkFont(family=self.font_family, size=13),
            fg_color="#64748b",
            hover_color="#475569",
            text_color="#ffffff",
            height=42,
            corner_radius=8,
            command=self.destroy
        )
        self.btn_close.pack(side="left", padx=6)

    def on_save_repo_clicked(self):
        new_repo = self.txt_repo.get().strip()
        if not new_repo or "/" not in new_repo:
            messagebox.showwarning("แจ้งเตือน", "กรุณาระบุชื่อ GitHub Repository ในรูปแบบ 'owner/repo' เช่น 'DietrichGebert/LeaveAutomation-1856'")
            return
        self.github_repo = new_repo
        if self.on_repo_save:
            self.on_repo_save(new_repo)
        self.start_check_update()

    def start_check_update(self):
        repo = self.txt_repo.get().strip() or self.github_repo
        self.lbl_status.configure(text="⏳ กำลังตรวจสอบข้อมูลเวอร์ชันล่าสุดจาก GitHub Releases...", text_color="#0284c7")
        self.btn_download.configure(state="disabled")
        self.btn_check.configure(state="disabled")
        self.txt_changelog.configure(state="normal")
        self.txt_changelog.delete("1.0", "end")
        self.txt_changelog.insert("1.0", "กำลังดึงข้อมูลบันทึกการเปลี่ยนแปลง...")
        self.txt_changelog.configure(state="disabled")

        def _worker():
            res = check_github_update(repo, self.current_version)
            self.after(0, lambda: self._on_check_completed(res))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_check_completed(self, res):
        self.btn_check.configure(state="normal")
        if not res.get("success"):
            err = res.get("error_msg", "ไม่สามารถตรวจสอบการอัปเดตได้")
            self.lbl_status.configure(text=f"⚠️ {err}", text_color="#ef4444")
            self.lbl_release_title.configure(text="ชื่อเวอร์ชัน: -")
            self.lbl_release_meta.configure(text="วันที่เผยแพร่: -  |  ขนาดไฟล์: -")
            self.txt_changelog.configure(state="normal")
            self.txt_changelog.delete("1.0", "end")
            self.txt_changelog.insert(
                "1.0",
                f"ไม่สามารถดึงข้อมูลได้:\n{err}\n\nคำแนะนำการแก้ปัญหา:\n"
                f"1. หากยังไม่ได้สร้าง GitHub Repository ให้สร้างบน https://github.com\n"
                f"2. ไปที่ Releases -> Create a new release\n"
                f"3. ตั้งชื่อแท็กเวอร์ชัน เช่น v1.0.1 แล้วอัปโหลดไฟล์ LeaveAutomation_1856.exe\n"
                f"4. ตรวจสอบชื่อ Repository ด้านบนให้ถูกต้อง แล้วกด 'บันทึก & ตรวจสอบ'"
            )
            self.txt_changelog.configure(state="disabled")
            self.btn_download.configure(state="disabled")
            return

        self.release_data = res
        tag = res.get("tag_name", "")
        name = res.get("release_name", "")
        body = res.get("body", "")
        pub = res.get("published_at", "")[:10]
        size_bytes = res.get("asset_size", 0)
        size_str = f"{size_bytes / (1024*1024):.1f} MB" if size_bytes > 0 else "ไม่ระบุ"
        asset_name = res.get("asset_name", "")

        self.lbl_release_title.configure(text=f"เวอร์ชัน {tag}: {name}")
        self.lbl_release_meta.configure(text=f"เผยแพร่เมื่อ: {pub}  |  ไฟล์: {asset_name or 'ไม่พบไฟล์'} ({size_str})")

        self.txt_changelog.configure(state="normal")
        self.txt_changelog.delete("1.0", "end")
        self.txt_changelog.insert("1.0", body)
        self.txt_changelog.configure(state="disabled")

        if res.get("update_available"):
            self.lbl_status.configure(text=f"ตรวจพบเวอร์ชันใหม่ {tag} (พร้อมให้อัปเดต)", text_color="#059669")
            if res.get("asset_url"):
                self.btn_download.configure(
                    text=f"⚡ ดาวน์โหลด & อัปเดตเป็น {tag} ทันที",
                    state="normal",
                    fg_color="#16a34a",
                    hover_color="#15803d"
                )
            else:
                self.lbl_status.configure(text=f"⚠️ พบเวอร์ชันใหม่ {tag} แต่ยังไม่มีไฟล์ .exe แนบใน Release", text_color="#eab308")
                self.btn_download.configure(state="disabled")
        else:
            self.lbl_status.configure(text=f"✅ ท่านกำลังใช้งานเวอร์ชันล่าสุดแล้ว ({self.current_version})", text_color="#0284c7")
            if res.get("asset_url"):
                self.btn_download.configure(
                    text=f"ดาวน์โหลดและติดตั้งเวอร์ชัน {tag} ซ้ำ",
                    state="normal",
                    fg_color="#0284c7",
                    hover_color="#0369a1"
                )
            else:
                self.btn_download.configure(state="disabled")

    def start_download_and_install(self):
        if not self.release_data or not self.release_data.get("asset_url"):
            messagebox.showwarning("แจ้งเตือน", "ไม่พบลิงก์ดาวน์โหลดสำหรับเวอร์ชันนี้")
            return

        tag = self.release_data.get("tag_name", "")
        confirm = messagebox.askyesno(
            "ยืนยันการอัปเดต",
            f"ระบบจะดาวน์โหลดเวอร์ชันใหม่ ({tag}) และรีสตาร์ทโปรแกรมโดยอัตโนมัติ\n\nต้องการดำเนินการต่อหรือไม่?"
        )
        if not confirm:
            return

        self.is_downloading = True
        self.is_cancelled = False
        self.btn_download.configure(state="disabled")
        self.btn_check.configure(state="disabled")
        self.btn_close.configure(state="disabled")
        self.prog_bar.set(0.0)
        self.lbl_progress.configure(text="⏳ กำลังเตรียมการเชื่อมต่อดาวน์โหลด...")

        asset_url = self.release_data.get("asset_url")
        asset_name = self.release_data.get("asset_name")
        total_size = self.release_data.get("asset_size", 0)

        threading.Thread(
            target=self._download_worker,
            args=(asset_url, asset_name, total_size),
            daemon=True
        ).start()

    def _download_worker(self, asset_url, asset_name, total_size):
        try:
            is_frozen = getattr(sys, 'frozen', False)
            if is_frozen:
                current_exe = os.path.abspath(sys.executable)
                exe_dir = os.path.dirname(current_exe)
                exe_name = os.path.basename(current_exe)
            else:
                exe_dir = os.path.dirname(os.path.abspath(__file__))
                exe_name = "LeaveAutomation_1856.exe"
                current_exe = os.path.join(exe_dir, exe_name)

            is_zip = asset_name.lower().endswith(".zip")
            if is_zip:
                temp_download = os.path.join(exe_dir, "temp_update_download.zip")
            else:
                temp_download = os.path.join(exe_dir, f"{exe_name}.new")

            req = urllib.request.Request(
                asset_url,
                headers={"User-Agent": f"LeaveAutomation-Updater/{APP_VERSION}"}
            )

            with urllib.request.urlopen(req, timeout=60) as resp:
                actual_size = int(resp.headers.get("Content-Length", total_size or 0))
                downloaded = 0
                chunk_size = 128 * 1024  # 128 KB
                with open(temp_download, "wb") as f:
                    while True:
                        if self.is_cancelled:
                            break
                        chunk = resp.read(chunk_size)
                        if not chunk:
                            break
                        f.write(chunk)
                        downloaded += len(chunk)

                        frac = (downloaded / actual_size) if actual_size > 0 else 0.5
                        mb_d = downloaded / (1024 * 1024)
                        mb_t = actual_size / (1024 * 1024) if actual_size > 0 else 0
                        percent = int(frac * 100) if actual_size > 0 else 50

                        def _up(p=percent, d=mb_d, t=mb_t, fr=frac):
                            self.prog_bar.set(fr)
                            self.lbl_progress.configure(
                                text=f"กำลังดาวน์โหลด: {d:.1f} MB / {t:.1f} MB ({p}%)"
                            )
                        self.after(0, _up)

            if self.is_cancelled:
                try:
                    if os.path.exists(temp_download):
                        os.remove(temp_download)
                except Exception:
                    pass
                return

            target_new_exe = os.path.join(exe_dir, f"{exe_name}.new")
            if is_zip:
                def _unzip_msg():
                    self.lbl_progress.configure(text="กำลังแตกไฟล์เวอร์ชันใหม่...")
                self.after(0, _unzip_msg)

                with zipfile.ZipFile(temp_download, 'r') as zf:
                    exe_inside = None
                    for name in zf.namelist():
                        if name.lower().endswith(".exe"):
                            exe_inside = name
                            break
                    if not exe_inside:
                        raise Exception("ไม่พบไฟล์ .exe ในไฟล์ .zip ที่ดาวน์โหลดมา")
                    with zf.open(exe_inside) as src, open(target_new_exe, "wb") as dst:
                        shutil.copyfileobj(src, dst)
                try:
                    os.remove(temp_download)
                except Exception:
                    pass

            # เตรียมการแทนที่และรันเวอร์ชันใหม่
            if is_frozen:
                def _prep_msg():
                    self.prog_bar.set(1.0)
                    self.lbl_progress.configure(text="กำลังติดตั้งและเริ่มโปรแกรมใหม่...")
                self.after(0, _prep_msg)

                bat_path = os.path.join(exe_dir, "updater.bat")
                bat_content = f"""@echo off
chcp 65001 >nul
set _MEIPASS2=
set _MEIPASS=
set PYI_PARENT_PID=
set PYTHONPATH=
set PYTHONHOME=
title อัปเดตโปรแกรม LeaveAutomation_1856
echo ========================================================
echo    กำลังดำเนินการอัปเดต LeaveAutomation_1856 เป็นเวอร์ชันใหม่...
echo    กรุณารอสักครู่ ห้ามปิดหน้าต่างนี้
echo ========================================================
timeout /t 2 /nobreak >nul
:retry
del "{current_exe}" 2>nul
if exist "{current_exe}" (
    timeout /t 1 /nobreak >nul
    goto retry
)
move /y "{target_new_exe}" "{current_exe}" >nul
if not exist "{current_exe}" (
    echo [ERROR] ไม่สามารถย้ายไฟล์เวอร์ชันใหม่ได้
    pause
    exit /b 1
)
echo ติดตั้งเสร็จสิ้น กำลังเปิดโปรแกรม...
set _MEIPASS2=
set _MEIPASS=
set PYI_PARENT_PID=
set PYTHONPATH=
set PYTHONHOME=
timeout /t 1 /nobreak >nul
start "" "{current_exe}"
del "%~f0"
"""
                with open(bat_path, "w", encoding="utf-8") as f:
                    f.write(bat_content)

                clean_env = os.environ.copy()
                for pyi_var in ['_MEIPASS2', '_MEIPASS', 'PYI_PARENT_PID', 'PYTHONPATH', 'PYTHONHOME']:
                    clean_env.pop(pyi_var, None)

                subprocess.Popen(
                    ["cmd.exe", "/c", bat_path],
                    cwd=exe_dir,
                    env=clean_env,
                    creationflags=subprocess.CREATE_NEW_CONSOLE if hasattr(subprocess, 'CREATE_NEW_CONSOLE') else 0
                )
                self.after(500, self._exit_app)
            else:
                # โหมดพัฒนา (Python dev)
                final_exe = os.path.join(exe_dir, "LeaveAutomation_1856.exe")
                if os.path.exists(final_exe):
                    try:
                        os.remove(final_exe)
                    except Exception:
                        pass
                shutil.move(target_new_exe, final_exe)

                def _done_dev():
                    self.prog_bar.set(1.0)
                    self.lbl_progress.configure(text="✅ ดาวน์โหลดเสร็จสมบูรณ์!")
                    self.btn_close.configure(state="normal")
                    messagebox.showinfo(
                        "อัปเดตสำเร็จ",
                        f"ดาวน์โหลดไฟล์เวอร์ชันใหม่เรียบร้อยแล้ว:\n{final_exe}\n\n(เนื่องจากรันในโหมด Python Dev จึงไม่แทนที่ python.exe)"
                    )
                self.after(0, _done_dev)

        except Exception as e:
            def _err(err_str=str(e)):
                self.lbl_progress.configure(text=f"❌ การดาวน์โหลดล้มเหลว: {err_str}", text_color="#ef4444")
                self.btn_download.configure(state="normal")
                self.btn_check.configure(state="normal")
                self.btn_close.configure(state="normal")
                messagebox.showerror("ข้อผิดพลาดในการอัปเดต", f"เกิดข้อผิดพลาดระหว่างดาวน์โหลด:\n{err_str}")
            self.after(0, _err)

    def _exit_app(self):
        try:
            self.parent.destroy()
        except Exception:
            pass
        os._exit(0)


# -------------------------------------------------------------
# หน้าต่างโปรแกรมหลัก โมเดิร์นพรีเมียม CustomTkinter (Senior-Friendly)
# -------------------------------------------------------------

class ModernCustomTkinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CNT STAFF TIME AUTOMATION - บริษัท คริสเตียนีและนีลเส็น (ไทย) จำกัด (มหาชน) [BY WERAWAT THAMRONGDULLAPARK]")
        self.root.geometry("1220x890")
        self.root.minsize(960, 640)
        
        self.font_family = "Tahoma"
        self.font_scale_var = ctk.StringVar(value="100%")
        self.font_scale = 1.0

        # ตัวแปรฟอนต์มาตรฐานสำหรับระบบ Senior-Friendly Dynamic Font Scaler (UI/UX Pro Max)
        self.font_title = ctk.CTkFont(family=self.font_family, size=20, weight="bold")
        self.font_h2 = ctk.CTkFont(family=self.font_family, size=14, weight="bold")
        self.font_body = ctk.CTkFont(family=self.font_family, size=13)
        self.font_btn = ctk.CTkFont(family=self.font_family, size=13, weight="bold")
        self.font_sm = ctk.CTkFont(family=self.font_family, size=11, weight="bold")
        self.font_cta = ctk.CTkFont(family=self.font_family, size=18, weight="bold")

        self.current_toast = None
        self.output_file = None
        self.report_file = None
        self.is_processing = False
        
        # โหลดไอคอนประจำโปรแกรม
        self.icon_path = ensure_clock_icon()
        if self.icon_path and os.path.exists(self.icon_path):
            try:
                self.root.iconbitmap(self.icon_path)
            except Exception:
                pass

        # โหลดโลโก้บริษัท Christiani & Nielsen
        self.logo_path = ensure_cnt_logo()
        self.cnt_logo_img = None
        if self.logo_path and os.path.exists(self.logo_path):
            try:
                from PIL import Image
                logo_pil = Image.open(self.logo_path)
                # รักษาอัตราส่วนของโลโก้ (107x78 -> height 40, width 55)
                self.cnt_logo_img = ctk.CTkImage(
                    light_image=logo_pil,
                    dark_image=logo_pil,
                    size=(55, 40)
                )
            except Exception:
                self.cnt_logo_img = None
                
        self.github_repo = DEFAULT_GITHUB_REPO
        self.setup_ui()
        self.load_saved_settings()

        # บันทึกการตั้งค่าอัตโนมัติเมื่อปิดโปรแกรม
        self.root.protocol("WM_DELETE_WINDOW", self.on_close_app)

        # ระบบ Drag & Drop ลากไฟล์มาวางบนโปรแกรมได้ทันที
        self.dnd_enabled = False
        try:
            import windnd
            windnd.hook_dropfiles(self.root, func=self.on_files_dropped)
            self.dnd_enabled = True
        except Exception:
            self.dnd_enabled = False

        # ตรวจสอบการอัปเดตแบบเบื้องหลัง (Background update check)
        self.root.after(2000, self.start_silent_update_check)

    def on_font_scale_changed(self, val):
        """จัดการเมื่อผู้ใช้เปลี่ยนสเกลตัวอักษร [100% | 115% | 125%]"""
        mapping = {"100%": 1.0, "115%": 1.15, "125%": 1.25}
        scale = mapping.get(val, 1.0)
        self._apply_font_scale(scale)
        self.save_current_settings()
        self.show_toast(f"ปรับขนาดตัวอักษรเป็น {val}", "info")

    def _apply_font_scale(self, scale):
        """คำนวณและอัปเดตขนาดตัวอักษรทุกตัวแปรฟอนต์แบบเรียลไทม์ (Senior-Friendly)"""
        self.font_scale = scale
        self.font_title.configure(size=int(20 * scale))
        self.font_h2.configure(size=int(14 * scale))
        self.font_body.configure(size=int(13 * scale))
        self.font_btn.configure(size=int(13 * scale))
        self.font_sm.configure(size=int(11 * scale))
        self.font_cta.configure(size=int(18 * scale))

    def show_toast(self, message, toast_type="info", duration_ms=2500):
        """ระบบการ์ดแจ้งเตือนแบบไม่ขัดจังหวะ (Non-blocking Toast Notification - UI/UX Pro Max)"""
        try:
            if self.current_toast:
                self.current_toast.destroy()
                self.current_toast = None
        except Exception:
            pass

        theme = {
            "success": ("#10b981", "#ffffff", "✅"),
            "info": ("#2563eb", "#ffffff", "ℹ️"),
            "warning": ("#f59e0b", "#ffffff", "⚠️"),
            "error": ("#ef4444", "#ffffff", "❌")
        }
        bg_col, fg_col, icon = theme.get(toast_type, ("#1e293b", "#ffffff", "🔔"))

        toast = ctk.CTkFrame(
            self.root,
            fg_color=bg_col,
            corner_radius=10,
            border_width=1,
            border_color="#ffffff"
        )
        toast.place(relx=0.98, rely=0.94, anchor="se")
        self.current_toast = toast

        content_box = ctk.CTkFrame(toast, fg_color="transparent")
        content_box.pack(padx=16, pady=9)

        lbl_ico = ctk.CTkLabel(content_box, text=icon, font=ctk.CTkFont(size=14))
        lbl_ico.pack(side="left", padx=(0, 8))

        lbl_txt = ctk.CTkLabel(
            content_box,
            text=message,
            font=ctk.CTkFont(family=self.font_family, size=13, weight="bold"),
            text_color=fg_col
        )
        lbl_txt.pack(side="left")

        def _fade_out():
            try:
                if self.current_toast == toast:
                    toast.destroy()
                    self.current_toast = None
            except Exception:
                pass

        self.root.after(duration_ms, _fade_out)

    def _flash_card(self, card):
        """สร้างเอฟเฟกต์ไฮไลต์ขอบการ์ดเมื่อลากไฟล์มาวาง (Visual Drag & Drop Feedback)"""
        if not card:
            return
        try:
            card.configure(border_color="#2563eb", border_width=2.5)
            def _reset():
                try:
                    card.configure(border_color=self.color_border, border_width=1.5)
                except Exception:
                    pass
            self.root.after(750, _reset)
        except Exception:
            pass

    def setup_ui(self):
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)
        
        # สีธีม Enterprise 2026 (Modern Slate Navy + High Contrast)
        self.color_navy = "#0f172a"          # Header Navy (Slate 900)
        self.color_primary = "#1d4ed8"       # CNT Corporate Blue (Blue 700)
        self.color_primary_hover = "#1e40af"
        self.color_accent = "#0284c7"        # Sky Blue for preview & secondary
        self.color_accent_hover = "#0369a1"
        self.color_emerald = "#059669"       # Emerald Green for Main Execution CTA
        self.color_emerald_hover = "#047857"
        self.color_pineapple = "#059669"     # Backward compatibility
        self.color_coconut = "#059669"       # Backward compatibility
        self.color_card_bg = "#ffffff"
        self.color_canvas = "#f1f5f9"        # Soft Slate 100
        self.color_border = "#cbd5e1"        # Slate 300
        self.color_border_subtle = "#e2e8f0" # Slate 200
        self.color_text_dark = "#0f172a"     # Slate 900 (High contrast)
        self.color_text_muted = "#475569"    # Slate 600

        # -------------------------------------------------------------
        # 1. แถบหัวเรื่องพรีเมียม (2026 Enterprise Responsive Header)
        # -------------------------------------------------------------
        header_frame = ctk.CTkFrame(
            self.root, 
            corner_radius=0, 
            fg_color=self.color_navy
        )
        header_frame.grid(row=0, column=0, sticky="ew")
        header_frame.grid_columnconfigure(0, weight=1)

        # แถว 1: อัตลักษณ์องค์กร (ซ้าย) + ตัวควบคุมระบบ & สเกลตัวอักษร (ขวา)
        header_top = ctk.CTkFrame(header_frame, fg_color="transparent")
        header_top.pack(fill="x", padx=20, pady=(10, 4))
        header_top.grid_columnconfigure(0, weight=1)

        brand_box = ctk.CTkFrame(header_top, fg_color="transparent")
        brand_box.grid(row=0, column=0, sticky="w")

        if self.cnt_logo_img:
            logo_badge = ctk.CTkFrame(
                brand_box,
                fg_color="#ffffff",
                corner_radius=8,
                border_width=1,
                border_color="#334155"
            )
            logo_badge.pack(side="left", padx=(0, 12))
            lbl_logo = ctk.CTkLabel(logo_badge, text="", image=self.cnt_logo_img)
            lbl_logo.pack(padx=6, pady=3)

        brand_text_box = ctk.CTkFrame(brand_box, fg_color="transparent")
        brand_text_box.pack(side="left")

        lbl_brand = ctk.CTkLabel(
            brand_text_box,
            text="CHRISTIANI & NIELSEN (THAI) PCL.",
            font=self.font_h2,
            text_color="#ffffff"
        )
        lbl_brand.pack(anchor="w")

        lbl_credit = ctk.CTkLabel(
            brand_text_box,
            text="STAFF TIME SUITE  |  BY WERAWAT THAMRONGDULLAPARK",
            font=self.font_sm,
            text_color="#93c5fd"
        )
        lbl_credit.pack(anchor="w", pady=(1, 0))

        # แถบควบคุม Utility ประจำระบบ (ขวาบน)
        utility_box = ctk.CTkFrame(header_top, fg_color="transparent")
        utility_box.grid(row=0, column=1, sticky="e")

        self.btn_update = ctk.CTkButton(
            utility_box,
            text=f"อัปเดต v{APP_VERSION}",
            font=self.font_btn,
            fg_color=self.color_primary,
            hover_color=self.color_primary_hover,
            text_color="#ffffff",
            height=32,
            width=110,
            corner_radius=6,
            command=self.open_updater
        )
        self.btn_update.pack(side="left", padx=(0, 8))

        # สเกลเลอร์ปรับขนาดตัวอักษรสำหรับผู้สูงวัย / สบายตา (Senior-Friendly Accessibility Scaler - UI/UX Pro Max)
        self.seg_font = ctk.CTkSegmentedButton(
            utility_box,
            values=["100%", "115%", "125%"],
            variable=self.font_scale_var,
            command=self.on_font_scale_changed,
            font=self.font_sm,
            height=32,
            width=135,
            selected_color=self.color_primary,
            selected_hover_color=self.color_primary_hover,
            unselected_color="#1e293b",
            unselected_hover_color="#334155",
            text_color="#ffffff",
            dynamic_resizing=False
        )
        self.seg_font.pack(side="left")

        # แถว 2: ชื่อโปรแกรมหลักเต็มความกว้าง (ไม่มีปุ่มมาเบียด ไม่ตกขอบ 100%)
        header_title_bar = ctk.CTkFrame(header_frame, fg_color="transparent")
        header_title_bar.pack(fill="x", padx=20, pady=(2, 4))

        lbl_title = ctk.CTkLabel(
            header_title_bar,
            text="ระบบบันทึกเวลาทำงานและประมวลผลการลาอัตโนมัติ (STAFF TIME)",
            font=self.font_title,
            text_color="#ffffff"
        )
        lbl_title.pack(anchor="w")

        lbl_subtitle = ctk.CTkLabel(
            header_title_bar,
            text="ระบบมาตรฐานส่วนกลางสำหรับทุกโครงการและทุกหน่วยงานของ CNT ทั่วประเทศ (Universal Edition)",
            font=self.font_body,
            text_color="#cbd5e1"
        )
        lbl_subtitle.pack(anchor="w", pady=(1, 0))

        # แถว 3: แถบเมนูเครื่องมือ 4 ระบบงาน (Navigation Ribbon - จัดวางอิสระ ไม่ล้นขอบ 100%)
        toolbar_box = ctk.CTkFrame(header_frame, fg_color="transparent")
        toolbar_box.pack(fill="x", padx=20, pady=(2, 10))

        btn_manual = ctk.CTkButton(
            toolbar_box,
            text="คู่มือการใช้งาน",
            font=self.font_btn,
            fg_color="#1e293b",
            hover_color="#334155",
            text_color="#ffffff",
            border_width=1,
            border_color="#475569",
            height=34,
            corner_radius=6,
            command=self.open_manual
        )
        btn_manual.pack(side="left", padx=(0, 6))

        btn_analytics = ctk.CTkButton(
            toolbar_box,
            text="สถิติภาพรวม",
            font=self.font_btn,
            fg_color="#1e293b",
            hover_color="#334155",
            text_color="#ffffff",
            border_width=1,
            border_color="#475569",
            height=34,
            corner_radius=6,
            command=self.open_analytics
        )
        btn_analytics.pack(side="left", padx=(0, 6))

        btn_anomaly = ctk.CTkButton(
            toolbar_box,
            text="วิเคราะห์เวลาผิดปกติ",
            font=self.font_btn,
            fg_color="#b45309",
            hover_color="#92400e",
            text_color="#ffffff",
            border_width=1,
            border_color="#d97706",
            height=34,
            corner_radius=6,
            command=self.open_anomaly_dialog
        )
        btn_anomaly.pack(side="left", padx=(0, 6))

        btn_inspector = ctk.CTkButton(
            toolbar_box,
            text="ตรวจสอบรายคน",
            font=self.font_btn,
            fg_color="#1e293b",
            hover_color="#334155",
            text_color="#ffffff",
            border_width=1,
            border_color="#475569",
            height=34,
            corner_radius=6,
            command=self.open_inspector
        )
        btn_inspector.pack(side="left", padx=(0, 6))

        # ผูก Tooltip บน Toolbar และ Utility
        add_tooltip(btn_manual, "เปิดคู่มือการใช้งานระบบ บันทึกเวลาและการประมวลผลการลาอย่างละเอียด")
        add_tooltip(btn_analytics, "เปิดหน้าต่างวิเคราะห์สถิติภาพรวม การมาทำงานและประเภทการลาทั้งหมด")
        add_tooltip(btn_anomaly, "เปิดระบบวิเคราะห์เวลาผิดปกติ (ขาดสแกนขาเดียว / ไม่พบสแกนและไม่ลา)")
        add_tooltip(btn_inspector, "เปิดระบบตรวจสอบประวัติเวลาทำงานและข้อมูลย้อนหลังรายบุคคล")
        add_tooltip(self.btn_update, "ตรวจสอบและอัปเดตโปรแกรมเป็นเวอร์ชันล่าสุดจาก GitHub")
        add_tooltip(self.seg_font, "ปรับขนาดตัวอักษรทั้งโปรแกรม (100% ปกติ | 115% กลาง | 125% ใหญ่พิเศษเพื่อความสบายตา)")

        # -------------------------------------------------------------
        # 2. พื้นที่เนื้อหาหลัก (Scrollable Card Container)
        # -------------------------------------------------------------
        main_scroll = ctk.CTkScrollableFrame(
            self.root,
            corner_radius=0,
            fg_color=self.color_canvas
        )
        main_scroll.grid(row=1, column=0, sticky="nsew", padx=16, pady=12)
        main_scroll.grid_columnconfigure(0, weight=1)

        # CARD 1: ไฟล์รายงาน STAFF TIME (.xlsx) - Compact
        self.card_main = ctk.CTkFrame(
            main_scroll,
            corner_radius=12,
            fg_color=self.color_card_bg,
            border_width=1.5,
            border_color=self.color_border
        )
        self.card_main.grid(row=0, column=0, sticky="ew", pady=(0, 6), padx=4)
        self.card_main.grid_columnconfigure(1, weight=1)

        c1_row = ctk.CTkFrame(self.card_main, fg_color="transparent")
        c1_row.pack(fill="x", padx=14, pady=8)
        c1_row.grid_columnconfigure(1, weight=1)

        badge_c1 = ctk.CTkLabel(
            c1_row,
            text="ขั้นตอนที่ 1",
            font=self.font_sm,
            fg_color=self.color_primary,
            text_color="#ffffff",
            corner_radius=6,
            height=28,
            width=80
        )
        badge_c1.grid(row=0, column=0, padx=(0, 10))

        self.txt_main = ctk.CTkEntry(
            c1_row,
            placeholder_text="คลิกปุ่มด้านขวา หรือ ลากไฟล์ STAFF TIME ประจำงวด (.xlsx) มาวางที่นี่...",
            font=self.font_body,
            height=38,
            corner_radius=8,
            border_color="#cbd5e1",
            fg_color="#f8fafc"
        )
        self.txt_main.grid(row=0, column=1, sticky="ew", padx=(0, 10))

        btn_browse_main = ctk.CTkButton(
            c1_row,
            text="เลือกไฟล์ STAFF TIME...",
            font=self.font_btn,
            fg_color=self.color_primary,
            hover_color=self.color_primary_hover,
            height=38,
            width=180,
            corner_radius=8,
            command=self.browse_main
        )
        btn_browse_main.grid(row=0, column=2)

        add_tooltip(self.txt_main, "ช่องระบุที่อยู่ไฟล์ STAFF TIME (.xlsx) - สามารถคลิกเลือกหรือลากไฟล์มาวางได้โดยตรง")
        add_tooltip(btn_browse_main, "คลิกเพื่อค้นหาและเลือกไฟล์รายงาน STAFF TIME ประจำงวด (.xlsx)")

        # CARD 2: ข้อมูลนำเข้า 2 แหล่ง (การลาที่ได้รับการอนุมัติ & ใบแจ้งการมาทำงาน) - Compact Single-Level
        self.card_sources = ctk.CTkFrame(
            main_scroll,
            corner_radius=12,
            fg_color=self.color_card_bg,
            border_width=1.5,
            border_color=self.color_border
        )
        self.card_sources.grid(row=1, column=0, sticky="ew", pady=(0, 6), padx=4)
        self.card_sources.grid_columnconfigure(0, weight=1)

        # แถบหัวข้อ Card 2 (กะทัดรัด)
        c2_head = ctk.CTkFrame(self.card_sources, fg_color="transparent")
        c2_head.pack(fill="x", padx=14, pady=(6, 2))

        badge_c2 = ctk.CTkLabel(
            c2_head,
            text="ขั้นตอนที่ 2",
            font=self.font_sm,
            fg_color=self.color_primary,
            text_color="#ffffff",
            corner_radius=6,
            height=26,
            width=80
        )
        badge_c2.pack(side="left", padx=(0, 10))

        lbl_c2_title = ctk.CTkLabel(
            c2_head,
            text="ไฟล์นำเข้าข้อมูล (วันลาที่ได้รับอนุมัติ และ/หรือ ใบแจ้งเวลาทำงานจาก CNTIWEB)",
            font=self.font_h2,
            text_color=self.color_text_dark
        )
        lbl_c2_title.pack(side="left")

        # 2.1 แถวข้อมูลวันลา (Row 1)
        r_leave = ctk.CTkFrame(self.card_sources, fg_color="transparent")
        r_leave.pack(fill="x", padx=14, pady=(2, 3))
        r_leave.grid_columnconfigure(1, weight=1)

        badge_leave = ctk.CTkLabel(
            r_leave,
            text="ข้อมูลวันลา",
            font=self.font_sm,
            fg_color=self.color_accent,
            text_color="#ffffff",
            corner_radius=6,
            height=28,
            width=80
        )
        badge_leave.grid(row=0, column=0, padx=(0, 10))

        self.txt_appr = ctk.CTkEntry(
            r_leave,
            placeholder_text="คลิกเลือก หรือลากไฟล์การลาที่ได้รับอนุมัติ (เช่น อนุมัติ.xlsx)...",
            font=self.font_body,
            height=36,
            corner_radius=8,
            border_color="#cbd5e1",
            fg_color="#f8fafc"
        )
        self.txt_appr.grid(row=0, column=1, sticky="ew", padx=(0, 8))

        btn_browse_appr = ctk.CTkButton(
            r_leave,
            text="เลือกไฟล์วันลา...",
            font=self.font_btn,
            fg_color=self.color_accent,
            hover_color=self.color_accent_hover,
            height=36,
            width=140,
            corner_radius=8,
            command=self.browse_appr
        )
        btn_browse_appr.grid(row=0, column=2, padx=(0, 6))

        btn_clear_appr = ctk.CTkButton(
            r_leave,
            text="ล้าง",
            font=self.font_sm,
            fg_color="#e2e8f0",
            hover_color="#cbd5e1",
            text_color="#334155",
            height=36,
            width=50,
            corner_radius=8,
            command=lambda: [self.txt_appr.delete(0, 'end'), self.save_current_settings()]
        )
        btn_clear_appr.grid(row=0, column=3)

        # 2.2 แถวใบแจ้งเวลาทำงาน (Row 2)
        r_att = ctk.CTkFrame(self.card_sources, fg_color="transparent")
        r_att.pack(fill="x", padx=14, pady=(2, 8))
        r_att.grid_columnconfigure(1, weight=1)

        badge_att = ctk.CTkLabel(
            r_att,
            text="ใบแจ้งเวลา",
            font=self.font_sm,
            fg_color=self.color_emerald,
            text_color="#ffffff",
            corner_radius=6,
            height=28,
            width=80
        )
        badge_att.grid(row=0, column=0, padx=(0, 10))

        self.txt_att = ctk.CTkEntry(
            r_att,
            placeholder_text="คลิกเลือก หรือลากไฟล์ใบแจ้งการมาทำงาน / ลืมสแกนนิ้ว (ถ้ามี)...",
            font=self.font_body,
            height=36,
            corner_radius=8,
            border_color="#cbd5e1",
            fg_color="#f8fafc"
        )
        self.txt_att.grid(row=0, column=1, sticky="ew", padx=(0, 8))

        btn_browse_att = ctk.CTkButton(
            r_att,
            text="เลือกใบแจ้งเวลา...",
            font=self.font_btn,
            fg_color=self.color_emerald,
            hover_color=self.color_emerald_hover,
            height=36,
            width=140,
            corner_radius=8,
            command=self.browse_att
        )
        btn_browse_att.grid(row=0, column=2, padx=(0, 6))

        btn_clear_att = ctk.CTkButton(
            r_att,
            text="ล้าง",
            font=self.font_sm,
            fg_color="#e2e8f0",
            hover_color="#cbd5e1",
            text_color="#334155",
            height=36,
            width=50,
            corner_radius=8,
            command=lambda: [self.txt_att.delete(0, 'end'), self.save_current_settings()]
        )
        btn_clear_att.grid(row=0, column=3)

        add_tooltip(self.txt_appr, "ช่องระบุไฟล์รายงานการลาที่ได้รับอนุมัติจาก CNTIWEB (.xlsx / .csv)")
        add_tooltip(btn_browse_appr, "เลือกไฟล์การลาที่ได้รับอนุมัติจากระบบ CNTIWEB")
        add_tooltip(btn_clear_appr, "ล้างไฟล์การลาที่เลือกออก")
        add_tooltip(self.txt_att, "ช่องระบุไฟล์ใบแจ้งการมาทำงาน / ลืมสแกนนิ้ว จาก CNTIWEB (.xlsx / .csv)")
        add_tooltip(btn_browse_att, "เลือกไฟล์ใบแจ้งเวลาทำงาน / ลืมสแกนนิ้ว")
        add_tooltip(btn_clear_att, "ล้างไฟล์ใบแจ้งเวลาทำงานที่เลือกออก")

        # CARD 3: กำหนดช่องคอลัมน์เป้าหมายในไฟล์ STAFF TIME (Universal)
        self.card_cols = ctk.CTkFrame(
            main_scroll,
            corner_radius=12,
            fg_color=self.color_card_bg,
            border_width=1.5,
            border_color=self.color_border
        )
        self.card_cols.grid(row=2, column=0, sticky="ew", pady=(0, 6), padx=4)

        self.cols_expanded = False

        self.btn_toggle_cols = ctk.CTkButton(
            self.card_cols,
            text="[+]  การกำหนดตำแหน่งคอลัมน์ในไฟล์ STAFF TIME (คลิกเพื่อเปิดแก้ไข  |  ค่ามาตรฐาน: A, B, E, G, L, M, Q, แถว 2)",
            font=self.font_btn,
            fg_color="#f8fafc",
            hover_color="#f1f5f9",
            text_color=self.color_primary,
            border_width=1,
            border_color="#cbd5e1",
            anchor="w",
            height=36,
            corner_radius=8,
            command=self.toggle_column_settings
        )
        self.btn_toggle_cols.pack(fill="x", padx=10, pady=6)

        self.frame_cols_body = ctk.CTkFrame(self.card_cols, fg_color="transparent")

        header_cols_box = ctk.CTkFrame(self.frame_cols_body, fg_color="transparent")
        header_cols_box.pack(fill="x", padx=10, pady=(6, 4))
        header_cols_box.grid_columnconfigure(0, weight=1)

        lbl_cols_title = ctk.CTkLabel(
            header_cols_box,
            text="ปรับแต่งตำแหน่งคอลัมน์เป้าหมายในไฟล์ STAFF TIME:",
            font=self.font_h2,
            text_color=self.color_text_dark
        )
        lbl_cols_title.grid(row=0, column=0, sticky="w")

        btn_box_cols = ctk.CTkFrame(header_cols_box, fg_color="transparent")
        btn_box_cols.grid(row=0, column=1, sticky="e")

        btn_save_cols = ctk.CTkButton(
            btn_box_cols,
            text="บันทึกตำแหน่งคอลัมน์",
            font=self.font_sm,
            fg_color=self.color_emerald,
            hover_color=self.color_emerald_hover,
            text_color="#ffffff",
            height=32,
            width=150,
            corner_radius=8,
            command=self.save_column_settings_explicit
        )
        btn_save_cols.pack(side="left", padx=(0, 8))

        btn_reset_cols = ctk.CTkButton(
            btn_box_cols,
            text="คืนค่ามาตรฐาน",
            font=self.font_sm,
            fg_color="#f1f5f9",
            hover_color="#e2e8f0",
            text_color=self.color_primary,
            border_width=1,
            border_color="#cbd5e1",
            height=32,
            width=120,
            corner_radius=8,
            command=self.reset_column_defaults
        )
        btn_reset_cols.pack(side="left")

        lbl_cols_hint = ctk.CTkLabel(
            self.frame_cols_body,
            text="* สามารถแก้ไขตัวอักษรช่องคอลัมน์ได้ทันที หากตารางข้อมูลของแต่ละโครงการหรือแต่ละหน่วยงานอยู่คนละช่องกัน (ช่อง B ใช้ทำ Double-Check ชื่อพนักงาน)",
            font=self.font_sm,
            text_color=self.color_text_muted
        )
        lbl_cols_hint.pack(anchor="w", padx=10, pady=(0, 10))

        col_inputs_frame = ctk.CTkFrame(self.frame_cols_body, fg_color="#f8fafc", corner_radius=10, border_width=1, border_color="#cbd5e1")
        col_inputs_frame.pack(fill="x", padx=6, pady=(0, 10))
        for col_i in range(8):
            col_inputs_frame.grid_columnconfigure(col_i, weight=1)

        # 1. รหัสพนักงาน
        lbl_c_emp = ctk.CTkLabel(col_inputs_frame, text="รหัสพนักงาน", font=self.font_sm, text_color=self.color_text_dark)
        lbl_c_emp.grid(row=0, column=0, pady=(10, 2))
        self.entry_col_emp = ctk.CTkEntry(col_inputs_frame, font=self.font_h2, width=65, height=38, justify="center")
        self.entry_col_emp.insert(0, "A")
        self.entry_col_emp.grid(row=1, column=0, pady=(0, 10))

        # 2. ชื่อพนักงาน
        lbl_c_name = ctk.CTkLabel(col_inputs_frame, text="ชื่อพนักงาน", font=self.font_sm, text_color=self.color_text_dark)
        lbl_c_name.grid(row=0, column=1, pady=(10, 2))
        self.entry_col_name = ctk.CTkEntry(col_inputs_frame, font=self.font_h2, width=65, height=38, justify="center")
        self.entry_col_name.insert(0, "B")
        self.entry_col_name.grid(row=1, column=1, pady=(0, 10))

        # 3. สแกนนิ้วดิบ
        lbl_c_scan = ctk.CTkLabel(col_inputs_frame, text="สแกนดิบ", font=self.font_sm, text_color=self.color_text_dark)
        lbl_c_scan.grid(row=0, column=2, pady=(10, 2))
        self.entry_col_scan = ctk.CTkEntry(col_inputs_frame, font=self.font_h2, width=65, height=38, justify="center")
        self.entry_col_scan.insert(0, "E")
        self.entry_col_scan.grid(row=1, column=2, pady=(0, 10))

        # 4. ประเภทลา / หมายเหตุ
        lbl_c_leave = ctk.CTkLabel(col_inputs_frame, text="ประเภทลา", font=self.font_sm, text_color=self.color_text_dark)
        lbl_c_leave.grid(row=0, column=3, pady=(10, 2))
        self.entry_col_leave = ctk.CTkEntry(col_inputs_frame, font=self.font_h2, width=65, height=38, justify="center")
        self.entry_col_leave.insert(0, "G")
        self.entry_col_leave.grid(row=1, column=3, pady=(0, 10))

        # 5. เวลาเข้า
        lbl_c_in = ctk.CTkLabel(col_inputs_frame, text="เวลาเข้า (In)", font=self.font_sm, text_color=self.color_text_dark)
        lbl_c_in.grid(row=0, column=4, pady=(10, 2))
        self.entry_col_in = ctk.CTkEntry(col_inputs_frame, font=self.font_h2, width=65, height=38, justify="center")
        self.entry_col_in.insert(0, "L")
        self.entry_col_in.grid(row=1, column=4, pady=(0, 10))

        # 6. เวลาออก
        lbl_c_out = ctk.CTkLabel(col_inputs_frame, text="เวลาออก (Out)", font=self.font_sm, text_color=self.color_text_dark)
        lbl_c_out.grid(row=0, column=5, pady=(10, 2))
        self.entry_col_out = ctk.CTkEntry(col_inputs_frame, font=self.font_h2, width=65, height=38, justify="center")
        self.entry_col_out.insert(0, "M")
        self.entry_col_out.grid(row=1, column=5, pady=(0, 10))

        # 7. รหัสย่อ
        lbl_c_code = ctk.CTkLabel(col_inputs_frame, text="รหัสย่อ (Code)", font=self.font_sm, text_color=self.color_text_dark)
        lbl_c_code.grid(row=0, column=6, pady=(10, 2))
        self.entry_col_code = ctk.CTkEntry(col_inputs_frame, font=self.font_h2, width=65, height=38, justify="center")
        self.entry_col_code.insert(0, "Q")
        self.entry_col_code.grid(row=1, column=6, pady=(0, 10))

        # 8. เริ่มแถวที่
        lbl_c_row = ctk.CTkLabel(col_inputs_frame, text="เริ่มแถวที่", font=self.font_sm, text_color=self.color_text_dark)
        lbl_c_row.grid(row=0, column=7, pady=(10, 2))
        self.entry_start_row = ctk.CTkEntry(col_inputs_frame, font=self.font_h2, width=65, height=38, justify="center")
        self.entry_start_row.insert(0, "2")
        self.entry_start_row.grid(row=1, column=7, pady=(0, 10))

        add_tooltip(self.btn_toggle_cols, "คลิกเพื่อเปิด/ปิด แผงกำหนดตำแหน่งคอลัมน์ในไฟล์ STAFF TIME")
        add_tooltip(btn_save_cols, "บันทึกตำแหน่งคอลัมน์เหล่านี้ไว้ถาวร (ระบบจะจำไว้ใช้ในครั้งต่อไป)")
        add_tooltip(btn_reset_cols, "คืนค่าตำแหน่งคอลัมน์กลับเป็นค่ามาตรฐาน (A, B, E, G, L, M, Q, แถว 2)")
        add_tooltip(self.entry_col_emp, "คอลัมน์รหัสพนักงาน (ค่ามาตรฐาน: A)")
        add_tooltip(self.entry_col_name, "คอลัมน์ชื่อ-นามสกุลพนักงาน (ค่ามาตรฐาน: B)")
        add_tooltip(self.entry_col_scan, "คอลัมน์บันทึกเวลาสแกนนิ้วดิบ (ค่ามาตรฐาน: E)")
        add_tooltip(self.entry_col_leave, "คอลัมน์ประเภทการลา / เหตุผล (ค่ามาตรฐาน: G)")
        add_tooltip(self.entry_col_in, "คอลัมน์เวลาเข้างานที่บันทึก (ค่ามาตรฐาน: L)")
        add_tooltip(self.entry_col_out, "คอลัมน์เวลาเลิกงานที่บันทึก (ค่ามาตรฐาน: M)")
        add_tooltip(self.entry_col_code, "คอลัมน์รหัสย่อการลา (ค่ามาตรฐาน: Q)")
        add_tooltip(self.entry_start_row, "แถวเริ่มต้นของข้อมูลพนักงานคนแรกในไฟล์ STAFF TIME (ค่ามาตรฐาน: 2)")

        # ผูก Event บันทึกอัตโนมัติทันทีที่พิมพ์หรือเปลี่ยนช่อง
        for entry_w in [self.entry_col_emp, self.entry_col_name, self.entry_col_scan,
                        self.entry_col_leave, self.entry_col_in, self.entry_col_out,
                        self.entry_col_code, self.entry_start_row]:
            entry_w.bind("<KeyRelease>", lambda e: self.save_current_settings())
            entry_w.bind("<FocusOut>", lambda e: self.save_current_settings())

        # แถบตั้งค่า GitHub Repository สำหรับระบบอัปเดตอัตโนมัติ
        repo_bar = ctk.CTkFrame(self.frame_cols_body, fg_color="#ffffff", corner_radius=10, border_width=1, border_color="#cbd5e1")
        repo_bar.pack(fill="x", padx=6, pady=(4, 10))

        lbl_repo = ctk.CTkLabel(
            repo_bar,
            text="ระบบอัปเดตอัตโนมัติ (GitHub Repository):",
            font=self.font_btn,
            text_color=self.color_text_dark
        )
        lbl_repo.pack(side="left", padx=(14, 8), pady=8)

        self.entry_github_repo = ctk.CTkEntry(
            repo_bar,
            font=self.font_body,
            width=280,
            height=36,
            fg_color="#f8fafc",
            border_color="#cbd5e1"
        )
        self.entry_github_repo.insert(0, self.github_repo)
        self.entry_github_repo.pack(side="left", padx=4, pady=8)

        btn_save_repo = ctk.CTkButton(
            repo_bar,
            text="บันทึก Repo",
            font=self.font_sm,
            fg_color=self.color_primary,
            hover_color=self.color_primary_hover,
            text_color="#ffffff",
            height=36,
            width=110,
            corner_radius=8,
            command=self.save_github_repo_from_entry
        )
        btn_save_repo.pack(side="left", padx=(4, 10), pady=8)
        add_tooltip(btn_save_repo, "บันทึก GitHub Repository สำหรับระบบอัปเดตอัตโนมัติ")

        # CARD 4: รูปแบบการบันทึกไฟล์ (Save Options - Single Line Compact)
        self.card_save = ctk.CTkFrame(
            main_scroll,
            corner_radius=12,
            fg_color=self.color_card_bg,
            border_width=1.5,
            border_color=self.color_border
        )
        self.card_save.grid(row=3, column=0, sticky="ew", pady=(0, 6), padx=4)

        save_row = ctk.CTkFrame(self.card_save, fg_color="transparent")
        save_row.pack(fill="x", padx=14, pady=6)

        badge_c4 = ctk.CTkLabel(
            save_row,
            text="การบันทึกไฟล์",
            font=self.font_sm,
            fg_color="#475569",
            text_color="#ffffff",
            corner_radius=6,
            height=26,
            width=85
        )
        badge_c4.pack(side="left", padx=(0, 16))

        self.save_mode_var = tk.StringVar(value="new_file")

        r1 = ctk.CTkRadioButton(
            save_row,
            text="สร้างไฟล์ใหม่ (_updated_วันที่ - แนะนำ)",
            variable=self.save_mode_var,
            value="new_file",
            font=self.font_btn,
            radiobutton_width=20,
            radiobutton_height=20,
            fg_color=self.color_primary,
            command=self.save_current_settings
        )
        r1.pack(side="left", padx=(0, 24))

        r2 = ctk.CTkRadioButton(
            save_row,
            text="บันทึกทับไฟล์เดิม (สร้างสำรอง _backup.xlsx อัตโนมัติ)",
            variable=self.save_mode_var,
            value="overwrite",
            font=self.font_btn,
            radiobutton_width=20,
            radiobutton_height=20,
            fg_color=self.color_primary,
            command=self.save_current_settings
        )
        r2.pack(side="left")

        add_tooltip(r1, "สร้างไฟล์ใหม่โดยต่อท้ายชื่อเดิมด้วย _updated_วันที่ ปลอดภัยต่อไฟล์ต้นฉบับ (แนะนำ)")
        add_tooltip(r2, "บันทึกทับไฟล์เดิมโดยตรง โดยระบบจะสร้างไฟล์สำรอง _backup.xlsx ป้องกันข้อมูลสูญหาย")

        # 5. ปุ่มแอคชันหลัก: ตรวจสอบรายการแก้ไข & ปุ่มเริ่มประมวลผล (Compact & Senior-Friendly)
        action_frame = ctk.CTkFrame(main_scroll, fg_color="transparent")
        action_frame.grid(row=4, column=0, sticky="ew", pady=(2, 6), padx=4)
        action_frame.grid_columnconfigure(0, weight=1)

        self.btn_preview = ctk.CTkButton(
            action_frame,
            text="ตรวจสอบรายการแก้ไขทั้งหมดก่อนบันทึกจริง (Pre-Check Preview)",
            font=self.font_h2,
            fg_color=self.color_accent,
            hover_color=self.color_accent_hover,
            height=42,
            corner_radius=8,
            command=self.open_preview
        )
        self.btn_preview.grid(row=0, column=0, sticky="ew", pady=(0, 5))

        self.btn_process = ctk.CTkButton(
            action_frame,
            text="เริ่มประมวลผลและบันทึกข้อมูลลงไฟล์ STAFF TIME ทันที",
            font=self.font_cta,
            fg_color=self.color_emerald,
            hover_color=self.color_emerald_hover,
            height=52,
            corner_radius=10,
            command=self.start_processing_thread
        )
        self.btn_process.grid(row=1, column=0, sticky="ew", pady=(0, 5))

        add_tooltip(self.btn_preview, "จำลองการประมวลผลเพื่อตรวจสอบรายการแก้ไขทุกจุดแบบละเอียดก่อนบันทึกลงไฟล์จริง")
        add_tooltip(self.btn_process, "เริ่มประมวลผล นำข้อมูลการลาและเวลาทำงานมาบันทึกลงไฟล์ STAFF TIME ทันที")

        self.progress_bar = ctk.CTkProgressBar(
            action_frame,
            height=10,
            corner_radius=5,
            progress_color=self.color_emerald,
            fg_color="#e2e8f0"
        )
        self.progress_bar.set(0.0)
        self.progress_bar.grid(row=2, column=0, sticky="ew")

        # 6. หน้าต่างแสดงผลการทำงาน (Live Log Box - Compact)
        self.card_log = ctk.CTkFrame(
            main_scroll,
            corner_radius=12,
            fg_color=self.color_card_bg,
            border_width=1.5,
            border_color=self.color_border
        )
        self.card_log.grid(row=5, column=0, sticky="ew", pady=(2, 4), padx=4)
        self.card_log.grid_columnconfigure(0, weight=1)

        log_top_frame = ctk.CTkFrame(self.card_log, fg_color="transparent")
        log_top_frame.grid(row=0, column=0, sticky="ew", padx=14, pady=(8, 4))
        log_top_frame.grid_columnconfigure(0, weight=1)

        log_h_box = ctk.CTkFrame(log_top_frame, fg_color="transparent")
        log_h_box.grid(row=0, column=0, sticky="w")

        badge_log = ctk.CTkLabel(
            log_h_box,
            text="ระบบตรวจสอบ",
            font=self.font_sm,
            fg_color="#475569",
            text_color="#ffffff",
            corner_radius=6,
            height=24,
            width=80
        )
        badge_log.pack(side="left", padx=(0, 8))

        lbl_log_title = ctk.CTkLabel(
            log_h_box,
            text="บันทึกสถานะการทำงาน (Activity Log)",
            font=self.font_h2,
            text_color=self.color_text_dark
        )
        lbl_log_title.pack(side="left")

        self.btn_open_report = ctk.CTkButton(
            log_top_frame,
            text="เปิดรายงานตรวจสอบ",
            font=self.font_sm,
            fg_color=self.color_emerald,
            hover_color=self.color_emerald_hover,
            height=30,
            corner_radius=6,
            state="disabled",
            command=self.open_audit_report
        )
        self.btn_open_report.grid(row=0, column=1, padx=(0, 6))

        self.btn_open_folder = ctk.CTkButton(
            log_top_frame,
            text="เปิดโฟลเดอร์ผลลัพธ์",
            font=self.font_sm,
            fg_color=self.color_primary,
            hover_color=self.color_primary_hover,
            height=30,
            corner_radius=6,
            state="disabled",
            command=self.open_result_folder
        )
        self.btn_open_folder.grid(row=0, column=2, padx=(0, 6))

        btn_clear_log = ctk.CTkButton(
            log_top_frame,
            text="ล้างข้อความ",
            font=self.font_sm,
            fg_color="#e2e8f0",
            hover_color="#cbd5e1",
            text_color="#334155",
            height=30,
            width=80,
            corner_radius=6,
            command=self.clear_log
        )
        btn_clear_log.grid(row=0, column=3)

        add_tooltip(self.btn_open_report, "เปิดไฟล์รายงานตรวจสอบ (Audit Report) สรุปรายการทั้งหมดที่อัปเดต")
        add_tooltip(self.btn_open_folder, "เปิดโฟลเดอร์ที่เก็บไฟล์ผลลัพธ์ใน File Explorer")
        add_tooltip(btn_clear_log, "ล้างข้อความประวัติในกล่องบันทึกสถานะ")

        self.txt_log = ctk.CTkTextbox(
            self.card_log,
            height=130,
            corner_radius=8,
            font=self.font_body,
            fg_color="#f8fafc",
            text_color="#0f172a",
            border_width=1.5,
            border_color="#cbd5e1"
        )
        self.txt_log.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 8))

        # กำหนดแท็กสีข้อความสำหรับบันทึกสถานะ (Semantic Status Badges - UI/UX Pro Max)
        try:
            tb = self.txt_log._textbox
            tb.tag_configure("tag_success", foreground="#16a34a", font=(self.font_family, 12, "bold"))
            tb.tag_configure("tag_skip", foreground="#d97706", font=(self.font_family, 12, "bold"))
            tb.tag_configure("tag_warn", foreground="#dc2626", font=(self.font_family, 12, "bold"))
            tb.tag_configure("tag_info", foreground="#2563eb", font=(self.font_family, 12, "bold"))
            tb.tag_configure("tag_normal", foreground="#0f172a", font=(self.font_family, 12))
        except Exception:
            pass

        # -------------------------------------------------------------
        # 3. แถบล่างสุดแสดงเครดิตผู้พัฒนา (Footer Status & Credits)
        # -------------------------------------------------------------
        self.root.grid_rowconfigure(2, weight=0)
        footer_bar = ctk.CTkFrame(self.root, fg_color=self.color_navy, corner_radius=0, height=32)
        footer_bar.grid(row=2, column=0, sticky="ew")
        lbl_footer = ctk.CTkLabel(
            footer_bar,
            text="CNT STAFF TIME AUTOMATION SUITE 2026  •  BY WERAWAT THAMRONGDULLAPARK  •  CHRISTIANI & NIELSEN (THAI) PCL.",
            font=self.font_sm,
            text_color="#94a3b8"
        )
        lbl_footer.pack(pady=5)

        self.log("ระบบพร้อมทำงาน... กรุณาเลือกไฟล์ STAFF TIME และข้อมูลที่ต้องการ")
        self.log("คำแนะนำ: สามารถกด 'ตรวจสอบรายการแก้ไขทั้งหมดก่อนบันทึกจริง' เพื่อดูพรีวิวทุกจุดก่อนบันทึกจริงได้เสมอ")

    # -------------------------------------------------------------
    # ระบบการบันทึก & โหลดการตั้งค่า (Settings Persistence)
    # -------------------------------------------------------------

    def load_saved_settings(self):
        settings = load_app_settings()
        if not settings:
            return

        main_f = settings.get("main_file", "")
        if main_f and os.path.exists(main_f):
            self.txt_main.insert(0, main_f)

        appr_f = settings.get("appr_file", "")
        if appr_f and os.path.exists(appr_f):
            self.txt_appr.insert(0, appr_f)

        att_f = settings.get("att_file", "")
        if att_f and os.path.exists(att_f):
            self.txt_att.insert(0, att_f)

        col_cfg = settings.get("col_config", {})
        if col_cfg:
            self.entry_col_emp.delete(0, 'end'); self.entry_col_emp.insert(0, col_cfg.get('emp_id_col', 'A'))
            self.entry_col_name.delete(0, 'end'); self.entry_col_name.insert(0, col_cfg.get('emp_name_col', 'B'))
            self.entry_col_scan.delete(0, 'end'); self.entry_col_scan.insert(0, col_cfg.get('scan_col', 'E'))
            self.entry_col_leave.delete(0, 'end'); self.entry_col_leave.insert(0, col_cfg.get('leave_type_col', 'G'))
            self.entry_col_in.delete(0, 'end'); self.entry_col_in.insert(0, col_cfg.get('checkin_col', 'L'))
            self.entry_col_out.delete(0, 'end'); self.entry_col_out.insert(0, col_cfg.get('checkout_col', 'M'))
            self.entry_col_code.delete(0, 'end'); self.entry_col_code.insert(0, col_cfg.get('leave_code_col', 'Q'))
            self.entry_start_row.delete(0, 'end'); self.entry_start_row.insert(0, str(col_cfg.get('start_row', '2')))

        sm = settings.get("save_mode", "new_file")
        if sm in ["new_file", "overwrite"]:
            self.save_mode_var.set(sm)

        # โหลดค่าสเกลตัวอักษรที่บันทึกไว้ (Senior-Friendly Scaler)
        scale_val = settings.get("font_scale", "100%")
        if scale_val in ["100%", "115%", "125%"]:
            self.font_scale_var.set(scale_val)
            mapping = {"100%": 1.0, "115%": 1.15, "125%": 1.25}
            self._apply_font_scale(mapping.get(scale_val, 1.0))

        self.github_repo = settings.get("github_repo", DEFAULT_GITHUB_REPO)
        if hasattr(self, "entry_github_repo"):
            self.entry_github_repo.delete(0, 'end')
            self.entry_github_repo.insert(0, self.github_repo)

    def save_current_settings(self):
        repo_val = self.github_repo
        if hasattr(self, "entry_github_repo"):
            val = self.entry_github_repo.get().strip()
            if val:
                repo_val = val
                self.github_repo = repo_val
        data = {
            "main_file": self.txt_main.get().strip(),
            "appr_file": self.txt_appr.get().strip(),
            "att_file": self.txt_att.get().strip(),
            "save_mode": self.save_mode_var.get(),
            "github_repo": repo_val,
            "col_config": self.get_column_config(),
            "font_scale": self.font_scale_var.get()
        }
        save_app_settings(data)

    def on_close_app(self):
        """บันทึกการตั้งค่าทั้งหมดก่อนปิดหน้าต่างโปรแกรม"""
        try:
            self.save_current_settings()
        except Exception:
            pass
        self.root.destroy()

    def save_column_settings_explicit(self):
        """บันทึกการตั้งค่าตำแหน่งคอลัมน์โดยผู้ใช้กดปุ่ม"""
        self.save_current_settings()
        self.log("บันทึกการตั้งค่าตำแหน่งคอลัมน์เป้าหมายลงระบบเรียบร้อย (ระบบจดจำถาวร)")
        self.show_toast("บันทึกตำแหน่งคอลัมน์เป้าหมายเรียบร้อย (จดจำถาวร)", "success")

    def on_files_dropped(self, files):
        """จัดการเมื่อผู้ใช้ลากไฟล์มาวางบนหน้าต่างโปรแกรม (Drag & Drop พร้อม Visual Feedback)"""
        if not files:
            return

        valid_paths = []
        for f in files:
            if isinstance(f, bytes):
                for enc in ['utf-8', 'mbcs', 'cp874']:
                    try:
                        f = f.decode(enc)
                        break
                    except UnicodeDecodeError:
                        continue
                if isinstance(f, bytes):
                    f = f.decode('utf-8', errors='ignore')
            path_str = os.path.normpath(str(f))
            if os.path.isfile(path_str) and path_str.lower().endswith(('.xlsx', '.xlsm', '.xls')):
                valid_paths.append(path_str)

        if not valid_paths:
            self.log("ไฟล์ที่ลากมาวางไม่ใช่ไฟล์ Excel (.xlsx, .xlsm, .xls)")
            self.show_toast("ไฟล์ที่ลากมาไม่ใช่ไฟล์ Excel", "warning")
            return

        if len(valid_paths) == 1:
            p = valid_paths[0]
            fn = os.path.basename(p).lower()
            if "อนุมัติ" in fn or "leave" in fn:
                self.txt_appr.delete(0, 'end')
                self.txt_appr.insert(0, p)
                self.save_current_settings()
                self._flash_card(self.card_sources)
                self.log(f"ลากและวางไฟล์การลาที่ได้รับอนุมัติ: {os.path.basename(p)}")
                self.show_toast(f"📥 ระบุไฟล์วันลา: {os.path.basename(p)}", "success")
            elif "ใบแจ้ง" in fn or "attendance" in fn or "ลืม" in fn:
                self.txt_att.delete(0, 'end')
                self.txt_att.insert(0, p)
                self.save_current_settings()
                self._flash_card(self.card_sources)
                self.log(f"ลากและวางไฟล์ใบแจ้งเวลาทำงาน: {os.path.basename(p)}")
                self.show_toast(f"📥 ระบุไฟล์ใบแจ้งเวลา: {os.path.basename(p)}", "success")
            else:
                self.txt_main.delete(0, 'end')
                self.txt_main.insert(0, p)
                self.save_current_settings()
                self._flash_card(self.card_main)
                self.log(f"ลากและวางไฟล์ STAFF TIME หลัก: {os.path.basename(p)}")
                self.show_toast(f"📥 ระบุไฟล์ STAFF TIME: {os.path.basename(p)}", "success")
        else:
            # ลากหลายไฟล์มาพร้อมกัน จัดเข้าแต่ละช่องให้อัตโนมัติ
            assigned_main = False
            for p in valid_paths:
                fn = os.path.basename(p).lower()
                if "อนุมัติ" in fn or "leave" in fn:
                    self.txt_appr.delete(0, 'end')
                    self.txt_appr.insert(0, p)
                    self.log(f"ระบุไฟล์การลาที่ได้รับอนุมัติ: {os.path.basename(p)}")
                elif "ใบแจ้ง" in fn or "attendance" in fn or "ลืม" in fn:
                    self.txt_att.delete(0, 'end')
                    self.txt_att.insert(0, p)
                    self.log(f"ระบุไฟล์ใบแจ้งเวลาทำงาน: {os.path.basename(p)}")
                elif not assigned_main:
                    self.txt_main.delete(0, 'end')
                    self.txt_main.insert(0, p)
                    assigned_main = True
                    self.log(f"ระบุไฟล์ STAFF TIME หลัก: {os.path.basename(p)}")
            self.save_current_settings()
            self._flash_card(self.card_main)
            self._flash_card(self.card_sources)
            self.log("จัดไฟล์นำเข้าลงช่องทั้งหมดเรียบร้อยแล้ว")
            self.show_toast(f"📥 จัดระเบียบ {len(valid_paths)} ไฟล์ลงช่องสำเร็จ", "success")

    # -------------------------------------------------------------
    # UI Actions & Modal Openers
    # -------------------------------------------------------------

    def open_updater(self):
        UpdateDialog(
            self.root,
            current_version=APP_VERSION,
            github_repo=self.github_repo,
            icon_path=self.icon_path,
            on_repo_save=self.on_updater_repo_saved
        )

    def on_updater_repo_saved(self, new_repo):
        self.github_repo = new_repo
        if hasattr(self, "entry_github_repo"):
            self.entry_github_repo.delete(0, 'end')
            self.entry_github_repo.insert(0, new_repo)
        self.save_current_settings()
        self.log(f"บันทึก GitHub Repository: {new_repo}")
        self.show_toast(f"บันทึก GitHub Repo: {new_repo}", "success")
        self.start_silent_update_check()

    def save_github_repo_from_entry(self):
        new_repo = self.entry_github_repo.get().strip()
        if not new_repo:
            new_repo = DEFAULT_GITHUB_REPO
            self.entry_github_repo.delete(0, 'end')
            self.entry_github_repo.insert(0, new_repo)
        self.github_repo = new_repo
        self.save_current_settings()
        self.log(f"บันทึก GitHub Repository เรียบร้อย: {new_repo}")
        self.show_toast(f"บันทึก GitHub Repo เรียบร้อย: {new_repo}", "success")
        self.start_silent_update_check()

    def start_silent_update_check(self):
        threading.Thread(target=self._worker_silent_check, daemon=True).start()

    def _worker_silent_check(self):
        try:
            res = check_github_update(self.github_repo, APP_VERSION)
            if res and res.get("success") and res.get("update_available"):
                tag = res.get("tag_name", "")
                def _update_ui():
                    if hasattr(self, "btn_update"):
                        self.btn_update.configure(
                            text=f"มีอัปเดตใหม่ ({tag})!",
                            fg_color="#f59e0b",
                            text_color="#ffffff",
                            hover_color="#d97706"
                        )
                self.root.after(0, _update_ui)
        except Exception:
            pass

    def open_manual(self):
        ManualDialog(self.root, self.icon_path)

    def open_inspector(self):
        main_file = self.txt_main.get().strip()
        appr_file = self.txt_appr.get().strip()
        att_file = self.txt_att.get().strip()
        col_cfg = self.get_column_config()

        if not main_file or not os.path.exists(main_file):
            messagebox.showerror("แจ้งเตือน", "กรุณาเลือกไฟล์ STAFF TIME (.xlsx) ให้ถูกต้องก่อนใช้งานระบบตรวจสอบรายคน")
            return

        InspectorDialog(self.root, main_file, appr_file, att_file, col_cfg, self.icon_path)

    def open_analytics(self):
        main_file = self.txt_main.get().strip()
        appr_file = self.txt_appr.get().strip()
        att_file = self.txt_att.get().strip()
        col_cfg = self.get_column_config()

        if not main_file or not os.path.exists(main_file):
            messagebox.showerror("แจ้งเตือน", "กรุณาเลือกไฟล์ STAFF TIME (.xlsx) ให้ถูกต้องก่อนดูสถิติภาพรวม")
            return

        AnalyticsDialog(self.root, main_file, appr_file, att_file, col_cfg, self.icon_path)

    def open_anomaly_dialog(self):
        main_file = self.txt_main.get().strip()
        if not main_file or not os.path.exists(main_file):
            messagebox.showerror("แจ้งเตือน", "กรุณาระบุไฟล์รายงาน STAFF TIME (.xlsx) ให้ถูกต้องก่อนเปิดระบบวิเคราะห์เวลาทำงานผิดปกติ")
            return
        AnomalyDialog(self.root, main_file=main_file, col_config=self.get_column_config(), icon_path=self.icon_path)

    def open_preview(self):
        main_file = self.txt_main.get().strip()
        appr_file = self.txt_appr.get().strip()
        att_file = self.txt_att.get().strip()
        col_cfg = self.get_column_config()

        if not main_file or not os.path.exists(main_file):
            messagebox.showerror("แจ้งเตือน", "กรุณาเลือกไฟล์ STAFF TIME (.xlsx) ให้ถูกต้องก่อนตรวจสอบรายการแก้ไข")
            return

        if (not appr_file or not os.path.exists(appr_file)) and (not att_file or not os.path.exists(att_file)):
            messagebox.showwarning("แจ้งเตือน", "กรุณาเลือกอย่างน้อย 1 ไฟล์ (การลาที่ได้รับการอนุมัติ หรือ ใบแจ้งการมาทำงาน)")
            return

        self.btn_preview.configure(text="⏳ กำลังจำลองและดึงรายการแก้ไขทั้งหมด...", state="disabled")
        self.progress_bar.set(0.05)

        def _worker():
            try:
                def _log(msg):
                    self.root.after(0, lambda m=msg: self.log(m))
                def _prog(val):
                    self.root.after(0, lambda v=val: self.progress_bar.set(v / 100.0))

                summary, items = simulate_all_changes(
                    main_file, appr_file, att_file, col_cfg,
                    log_callback=_log, progress_callback=_prog
                )

                def _show():
                    self.btn_preview.configure(
                        text="ตรวจสอบรายการแก้ไขทั้งหมดก่อนบันทึกจริง (Pre-Check Preview)",
                        state="normal"
                    )
                    self.progress_bar.set(1.0)
                    PreviewDialog(self.root, summary, items, on_proceed_callback=self.start_processing_thread, icon_path=self.icon_path)

                self.root.after(0, _show)
            except Exception as e:
                def _err(err_msg):
                    self.btn_preview.configure(
                        text="ตรวจสอบรายการแก้ไขทั้งหมดก่อนบันทึกจริง (Pre-Check Preview)",
                        state="normal"
                    )
                    self.progress_bar.set(0.0)
                    messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถจำลองการแก้ไขได้:\n{err_msg}")
                self.root.after(0, lambda err=str(e): _err(err))

        threading.Thread(target=_worker, daemon=True).start()

    def toggle_column_settings(self):
        if self.cols_expanded:
            self.frame_cols_body.pack_forget()
            self.btn_toggle_cols.configure(
                text="[+]  การกำหนดตำแหน่งคอลัมน์ในไฟล์ STAFF TIME (คลิกเพื่อเปิดแก้ไข  |  ค่ามาตรฐาน: A, B, E, G, L, M, Q, แถว 2)",
                fg_color="#f0f9ff"
            )
            self.cols_expanded = False
        else:
            self.frame_cols_body.pack(fill="x", padx=10, pady=(0, 10))
            self.btn_toggle_cols.configure(
                text="[-]  ซ่อนการกำหนดตำแหน่งคอลัมน์ในไฟล์ STAFF TIME (คลิกเพื่อพับเก็บ)",
                fg_color="#e0f2fe"
            )
            self.cols_expanded = True

    def reset_column_defaults(self):
        defaults = get_default_column_config()
        self.entry_col_emp.delete(0, 'end'); self.entry_col_emp.insert(0, defaults['emp_id_col'])
        self.entry_col_name.delete(0, 'end'); self.entry_col_name.insert(0, defaults['emp_name_col'])
        self.entry_col_scan.delete(0, 'end'); self.entry_col_scan.insert(0, defaults['scan_col'])
        self.entry_col_leave.delete(0, 'end'); self.entry_col_leave.insert(0, defaults['leave_type_col'])
        self.entry_col_in.delete(0, 'end'); self.entry_col_in.insert(0, defaults['checkin_col'])
        self.entry_col_out.delete(0, 'end'); self.entry_col_out.insert(0, defaults['checkout_col'])
        self.entry_col_code.delete(0, 'end'); self.entry_col_code.insert(0, defaults['leave_code_col'])
        self.entry_start_row.delete(0, 'end'); self.entry_start_row.insert(0, str(defaults['start_row']))
        self.save_current_settings()
        self.log("คืนค่าตำแหน่งคอลัมน์เป็นค่ามาตรฐานเรียบร้อย: A, B, E, G, L, M, Q, แถวที่ 2")
        self.show_toast("คืนค่าตำแหน่งคอลัมน์เป็นค่ามาตรฐานเรียบร้อย", "info")

    def get_column_config(self):
        return {
            'emp_id_col': self.entry_col_emp.get().strip().upper() or 'A',
            'emp_name_col': self.entry_col_name.get().strip().upper() or 'B',
            'scan_col': self.entry_col_scan.get().strip().upper() or 'E',
            'leave_type_col': self.entry_col_leave.get().strip().upper() or 'G',
            'checkin_col': self.entry_col_in.get().strip().upper() or 'L',
            'checkout_col': self.entry_col_out.get().strip().upper() or 'M',
            'leave_code_col': self.entry_col_code.get().strip().upper() or 'Q',
            'start_row': self.entry_start_row.get().strip() or '2'
        }

    def log(self, message):
        msg_str = str(message)
        tag = "tag_normal"
        if any(w in msg_str for w in ["สำเร็จ", "อัปเดต", "พร้อมทำงาน", "สมบูรณ์"]):
            tag = "tag_success"
        elif any(w in msg_str for w in ["ข้าม", "ละเว้น"]):
            tag = "tag_skip"
        elif any(w in msg_str for w in ["ผิดพลาด", "เตือน", "ไม่พบ", "❌"]):
            tag = "tag_warn"
        elif any(w in msg_str for w in ["โหลด", "เลือกไฟล์", "บันทึก", "คำแนะนำ", "ลากและวาง", "ระบุไฟล์"]):
            tag = "tag_info"

        try:
            tb = self.txt_log._textbox
            tb.insert("end", msg_str + "\n", tag)
            tb.see("end")
        except Exception:
            self.txt_log.insert("end", msg_str + "\n")
            self.txt_log.see("end")

    def clear_log(self):
        self.txt_log.delete("1.0", "end")

    def browse_main(self):
        filename = filedialog.askopenfilename(
            title="เลือกไฟล์ STAFF TIME (Staff Time Report)",
            filetypes=[("Excel Files", "*.xlsx *.xlsm"), ("All Files", "*.*")]
        )
        if filename:
            self.txt_main.delete(0, "end")
            self.txt_main.insert(0, filename)
            self.save_current_settings()
            self.log(f"เลือกไฟล์ STAFF TIME: {os.path.basename(filename)}")

    def browse_appr(self):
        filename = filedialog.askopenfilename(
            title="เลือกการลาที่ได้รับการอนุมัติ (CNTIWEB)",
            filetypes=[("Excel / CSV Files", "*.xlsx *.xls *.csv"), ("All Files", "*.*")]
        )
        if filename:
            self.txt_appr.delete(0, "end")
            self.txt_appr.insert(0, filename)
            self.save_current_settings()
            self.log(f"เลือกไฟล์การลาที่ได้รับอนุมัติ: {os.path.basename(filename)}")

    def browse_att(self):
        filename = filedialog.askopenfilename(
            title="เลือกใบแจ้งการมาทำงาน (CNTIWEB)",
            filetypes=[("Excel / CSV Files", "*.xlsx *.xls *.csv"), ("All Files", "*.*")]
        )
        if filename:
            self.txt_att.delete(0, "end")
            self.txt_att.insert(0, filename)
            self.save_current_settings()
            self.log(f"⏰ เลือกใบแจ้งการมาทำงาน: {os.path.basename(filename)}")

    def open_result_folder(self):
        if self.output_file and os.path.exists(self.output_file):
            folder = os.path.dirname(self.output_file)
            try:
                os.startfile(folder)
            except Exception:
                import subprocess
                subprocess.Popen(f'explorer "{folder}"')

    def open_audit_report(self):
        if self.report_file and os.path.exists(self.report_file):
            try:
                os.startfile(self.report_file)
            except Exception:
                import subprocess
                subprocess.Popen(f'explorer "{self.report_file}"')

    def start_processing_thread(self):
        if self.is_processing:
            return

        main_file = self.txt_main.get().strip()
        appr_file = self.txt_appr.get().strip()
        att_file = self.txt_att.get().strip()
        col_cfg = self.get_column_config()
        self.save_current_settings()

        if not main_file or not os.path.exists(main_file):
            messagebox.showerror("แจ้งเตือน", "กรุณาเลือกไฟล์ STAFF TIME (.xlsx) ให้ถูกต้องก่อนเริ่มประมวลผล")
            return

        if (not appr_file or not os.path.exists(appr_file)) and (not att_file or not os.path.exists(att_file)):
            messagebox.showwarning("แจ้งเตือน", "กรุณาเลือกอย่างน้อย 1 ไฟล์ (การลาที่ได้รับการอนุมัติ หรือ ใบแจ้งการมาทำงาน)")
            return

        self.is_processing = True
        self.btn_process.configure(
            text="⏳  กำลังประมวลผลข้อมูล กรุณารอสักครู่...",
            fg_color="#94a3b8",
            state="disabled"
        )
        self.btn_preview.configure(state="disabled")
        self.progress_bar.set(0.0)
        self.btn_open_folder.configure(state="disabled")
        self.btn_open_report.configure(state="disabled")

        t = threading.Thread(
            target=self._run_processing, 
            args=(main_file, appr_file or None, att_file or None, self.save_mode_var.get(), col_cfg),
            daemon=True
        )
        t.start()

    def _run_processing(self, main_file, appr_file, att_file, save_mode, col_cfg):
        try:
            out_file, rep_file, total_up, total_sk = process_all_files(
                main_filepath=main_file,
                approval_filepath=appr_file,
                attendance_filepath=att_file,
                save_mode=save_mode,
                col_config=col_cfg,
                log_callback=lambda msg: self.root.after(0, self.log, msg),
                progress_callback=lambda p: self.root.after(0, self.progress_bar.set, p / 100.0)
            )
            self.output_file = out_file
            self.report_file = rep_file
            
            def on_success():
                self.btn_process.configure(
                    text="เริ่มประมวลผลและบันทึกข้อมูลลงไฟล์ STAFF TIME ทันที",
                    fg_color=self.color_emerald,
                    state="normal"
                )
                self.btn_preview.configure(state="normal")
                self.is_processing = False
                if out_file:
                    self.btn_open_folder.configure(state="normal")
                if rep_file:
                    self.btn_open_report.configure(state="normal")
                    
                rep_msg = f"\n\nไฟล์รายงานตรวจสอบ (Audit Report):\n{rep_file}" if rep_file else ""
                messagebox.showinfo(
                    "ประมวลผลสำเร็จ",
                    f"อัปเดตข้อมูลเรียบร้อยแล้ว!\n\n"
                    f"• รายการที่อัปเดตสำเร็จ: {total_up} จุด\n"
                    f"• รายการที่ข้าม/ติดเงื่อนไข: {total_sk} จุด\n\n"
                    f"ไฟล์ผลลัพธ์:\n{out_file}"
                    f"{rep_msg}"
                )
            self.root.after(0, on_success)
            
        except Exception as e:
            def on_error(err_msg):
                self.log(f"\n❌ เกิดข้อผิดพลาด: {err_msg}")
                self.btn_process.configure(
                    text="เริ่มประมวลผลและบันทึกข้อมูลลงไฟล์ STAFF TIME ทันที",
                    fg_color=self.color_emerald,
                    state="normal"
                )
                self.btn_preview.configure(state="normal")
                self.is_processing = False
                messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถประมวลผลได้:\n{err_msg}")
            self.root.after(0, on_error, str(e))


# -------------------------------------------------------------
# Main Entry Point
# -------------------------------------------------------------

def main():
    root = ctk.CTk()
    app = ModernCustomTkinterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
