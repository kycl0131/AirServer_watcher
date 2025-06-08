import time
import ctypes
import winsound
import psutil
from datetime import datetime, timedelta
import logging

import pyautogui
import pygetwindow as gw
import win32gui
import win32con
import keyboard   # pip install keyboard

# ─────────── 로깅 설정 ───────────
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger(__name__)

# ─────────── Win32 디스플레이 절전 플래그 ───────────
ES_CONTINUOUS       = 0x80000000
ES_DISPLAY_REQUIRED = 0x00000002

def prevent_display_off():
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_DISPLAY_REQUIRED)
    logger.info("⏱️ Display sleep prevented")

def clear_execution_state():
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
    logger.info("↩️ Execution state cleared")

# ─────────── Auto-fullscreen 토글 ───────────
AUTO_FULLSCREEN = True
def toggle_fullscreen_flag():
    global AUTO_FULLSCREEN
    AUTO_FULLSCREEN = not AUTO_FULLSCREEN
    logger.info(f"🔀 Auto-fullscreen: {'ON' if AUTO_FULLSCREEN else 'OFF'}")
keyboard.add_hotkey('ctrl+alt+f', toggle_fullscreen_flag)

# ─────────── AirPlay 감시 설정 ───────────
PROCESS_NAME = "AirServer.exe"
WATCH_PORTS  = {5000, 7000}

def count_established_sessions():
    cnt = 0; video_cnt = 0
    for conn in psutil.net_connections(kind='tcp'):
        if conn.status == psutil.CONN_ESTABLISHED and conn.laddr.port in WATCH_PORTS:
            cnt += 1
            if conn.laddr.port == 7000:
                video_cnt += 1
    return cnt, video_cnt

def is_airplay_active():
    proc_found = any(p.info.get('name','').lower()==PROCESS_NAME.lower()
                     for p in psutil.process_iter(['name']))
    sessions, video_sessions = count_established_sessions()
    return proc_found and sessions>0, video_sessions>0

# ─────────── 스크린세이버 체크 ───────────
SPI_GETSCREENSAVERRUNNING = 114
def is_screensaver_running():
    val = ctypes.c_int()
    ctypes.windll.user32.SystemParametersInfoW(
        SPI_GETSCREENSAVERRUNNING, 0,
        ctypes.byref(val), 0
    )
    return bool(val.value)

def wait_for_screensaver_off(timeout=5, poll=0.2):
    start = time.time()
    while time.time()-start < timeout:
        if not is_screensaver_running():
            return True
        time.sleep(poll)
    return False

# ─────────── 창 핸들 찾기/강제 포그라운드 ───────────
def find_airserver_window():
    for w in gw.getAllWindows():
        if "AirServer Windows Desktop Edition" in w.title:
            return w._hWnd
    return None

def force_set_foreground(hwnd):
    user32   = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    fg       = user32.GetForegroundWindow()
    cur_tid  = kernel32.GetCurrentThreadId()
    fg_tid   = user32.GetWindowThreadProcessId(fg, None)
    user32.AttachThreadInput(cur_tid, fg_tid, True)
    user32.SetForegroundWindow(hwnd)
    user32.AttachThreadInput(cur_tid, fg_tid, False)

# ─────────── 글로벌 상태 ───────────
airserver_hwnd      = None
last_was_video      = False
_original_mouse_pos = None

# ─────────── 유틸리티 ───────────
def jiggle_mouse_and_save(delta=1):
    global _original_mouse_pos
    _original_mouse_pos = pyautogui.position()
    ctypes.windll.user32.mouse_event(0x0001, delta, 0, 0, 0)
    logger.info("🖱️ Low-level jiggle & saved original position")

def restore_mouse_position():
    global _original_mouse_pos
    if _original_mouse_pos:
        pyautogui.moveTo(_original_mouse_pos)
        logger.info(f"🖱️ Restored mouse to {_original_mouse_pos}")
        _original_mouse_pos = None

def move_mouse_away():
    if not airserver_hwnd or not win32gui.IsWindow(airserver_hwnd):
        return
    l,t,r,b = win32gui.GetWindowRect(airserver_hwnd)
    sw,sh   = pyautogui.size()
    x = max(10, min(sw-50, r+50))
    y = max(10, min(sh-50, b+50))
    pyautogui.moveTo(x, y)
    logger.info(f"🖱️ Mouse moved to ({x},{y})")

def raise_window():
    global airserver_hwnd
    if not airserver_hwnd or not win32gui.IsWindow(airserver_hwnd):
        airserver_hwnd = find_airserver_window()
    if not airserver_hwnd:
        logger.warning("AirServer window not found")
        return False

    if is_screensaver_running():
        wait_for_screensaver_off()

    # 복원 & 포그라운드
    if win32gui.IsIconic(airserver_hwnd):
        win32gui.ShowWindow(airserver_hwnd, win32con.SW_RESTORE)
        time.sleep(0.2)
    win32gui.ShowWindow(airserver_hwnd, win32con.SW_SHOWNORMAL)
    win32gui.BringWindowToTop(airserver_hwnd)
    try:
        win32gui.SetForegroundWindow(airserver_hwnd)
    except:
        force_set_foreground(airserver_hwnd)
    logger.info("🔝 AirServer window raised")
    return True

def double_fullscreen():
    l,t,r,b = win32gui.GetWindowRect(airserver_hwnd)
    pyautogui.moveTo((l+r)//2, (t+b)//2)
    pyautogui.doubleClick()
    logger.info("⛶ Fullscreen toggled")

def lower_window():
    global last_was_video
    if AUTO_FULLSCREEN and last_was_video:
        double_fullscreen()
        time.sleep(0.5)
    if airserver_hwnd and win32gui.IsWindow(airserver_hwnd):
        win32gui.SetWindowPos(
            airserver_hwnd, win32con.HWND_BOTTOM,
            0,0,0,0,
            win32con.SWP_NOMOVE|win32con.SWP_NOSIZE
        )
        logger.info("🔽 AirServer window lowered")

# ─────────── 메인 루프 ───────────
def main():
    global last_was_video
    was_active     = False
    last_keepalive = datetime.min

    logger.info("▶️ AirPlay Monitor started")
    while True:
        active, is_video = is_airplay_active()

        if active and not was_active:
            last_was_video = is_video
            logger.info(f"+ Session started (video={is_video})")

            jiggle_mouse_and_save()
            if raise_window():
                prevent_display_off()
                if AUTO_FULLSCREEN and last_was_video:
                    double_fullscreen()
                    time.sleep(2)    # 컨트롤 바 숨김 대기
                    move_mouse_away()

            was_active     = True
            last_keepalive = datetime.now()

        elif active and was_active:
            if datetime.now() - last_keepalive > timedelta(minutes=5):
                prevent_display_off()
                last_keepalive = datetime.now()

        elif not active and was_active:
            logger.info("- Session ended")
            # **변경**: 먼저 창 내리고
            lower_window()
            # 화면절전 해제 취소
            clear_execution_state()
            # 마우스 원위치 복원
            restore_mouse_position()
            was_active = False

        time.sleep(1)

if __name__ == "__main__":
    main()
