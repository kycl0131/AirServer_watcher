import time
import ctypes
import winsound
import psutil
from datetime import datetime, timedelta

import pyautogui
import pygetwindow as gw
import win32gui
import win32con
import keyboard   # pip install keyboard

# ─────────── Win32 디스플레이 절전 플래그 ───────────
ES_CONTINUOUS       = 0x80000000
ES_DISPLAY_REQUIRED = 0x00000002

def prevent_display_off():
    ctypes.windll.kernel32.SetThreadExecutionState(
        ES_CONTINUOUS | ES_DISPLAY_REQUIRED
    )
    print(f"[⏱️] prevent_display_off: {datetime.now()}")

def clear_execution_state():
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
    print(f"[↩️] clear_execution_state: {datetime.now()}")

# ─────────── Auto-fullscreen 토글 ───────────
AUTO_FULLSCREEN = True
def toggle_fullscreen_flag():
    global AUTO_FULLSCREEN
    AUTO_FULLSCREEN = not AUTO_FULLSCREEN
    print(f"[🔀] Auto-fullscreen: {'ON' if AUTO_FULLSCREEN else 'OFF'} ({datetime.now()})")

keyboard.add_hotkey('ctrl+alt+f', toggle_fullscreen_flag)

# ─────────── AirPlay 감시 설정 ───────────
PROCESS_NAME     = "AirServer.exe"
WATCH_PORTS      = {5000, 7000}

def count_established_sessions():
    """psutil로 ESTABLISHED된 TCP 연결 중 WATCH_PORTS에 속한 것을 센다."""
    cnt = 0
    for conn in psutil.net_connections(kind='tcp'):
        if conn.status == psutil.CONN_ESTABLISHED and conn.laddr.port in WATCH_PORTS:
            cnt += 1
    print(f"[PSUTIL] established sessions on {WATCH_PORTS}: {cnt}")
    return cnt

def is_airplay_active():
    """AirServer 프로세스 존재 + 최소 1개 이상의 연결이 있으면 활성."""
    # 프로세스 체크
    proc_found = any(p.name().lower() == PROCESS_NAME.lower() for p in psutil.process_iter())
    # 세션 체크
    sessions = count_established_sessions()
    active = proc_found and sessions >= 1
    print(f"[PSUTIL] process_found={proc_found}, sessions={sessions} -> active={active}")
    return active, sessions

# ─────────── 글로벌 상태 저장 ───────────
airserver_hwnd  = None
last_was_video  = False

# ─────────── 유틸리티 ───────────
def play_beep(freq, dur):
    winsound.Beep(freq, dur)

def jiggle_mouse():
    x, y = pyautogui.position()
    pyautogui.moveTo(x+1, y); pyautogui.moveTo(x, y)
    print(f"[🖱️] jiggle_mouse: {datetime.now()}")

def raise_airserver_window():
    global airserver_hwnd
    target = "AirServer Windows Desktop Edition"
    wins = [w for w in gw.getAllWindows() if target in w.title]
    if not wins:
        print("[!] AirServer 창을 못 찾음"); airserver_hwnd = None; return
    win = wins[0]; airserver_hwnd = win._hWnd
    if win.isMinimized: win.restore()
    win32gui.ShowWindow(airserver_hwnd, win32con.SW_SHOWNORMAL)
    win32gui.BringWindowToTop(airserver_hwnd)
    try: win32gui.SetForegroundWindow(airserver_hwnd)
    except: pass
    print(f"[🔝] raise_airserver_window: {datetime.now()}")

def double_click_fullscreen():
    if not airserver_hwnd or not win32gui.IsWindow(airserver_hwnd):
        print("[!] No valid hwnd"); return
    left, top, right, bottom = win32gui.GetWindowRect(airserver_hwnd)
    pyautogui.moveTo((left+right)//2, (top+bottom)//2)
    pyautogui.doubleClick()
    print(f"[⛶] double_click_fullscreen: {datetime.now()}")

def lower_airserver_window():
    global airserver_hwnd, last_was_video
    if AUTO_FULLSCREEN and last_was_video:
        double_click_fullscreen()
        time.sleep(0.5)
    if not airserver_hwnd or not win32gui.IsWindow(airserver_hwnd):
        return
    win32gui.SetWindowPos(
        airserver_hwnd, win32con.HWND_BOTTOM,
        0, 0, 0, 0,
        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
    )
    print(f"[🔽] lower_airserver_window: {datetime.now()}")

# ─────────── 메인 루프 ───────────
was_active     = False
last_keepalive = datetime.min

print(f"[DEBUG] start-up check")
while True:
    active, sessions = is_airplay_active()
    # 포트 7000 연결이 있으면 video 모드
    is_video = sessions and any(conn.laddr.port == 7000 for conn in psutil.net_connections(kind='tcp') if conn.status == psutil.CONN_ESTABLISHED)
    print(f"[DEBUG] loop: active={active}, was_active={was_active}, is_video={is_video}")

    if active and not was_active:
        last_was_video = is_video
        print(f"[+] AirPlay 시작됨 (video={is_video})")
        raise_airserver_window()
        play_beep(1200, 200)
        jiggle_mouse()
        prevent_display_off()
        if AUTO_FULLSCREEN and last_was_video:
            double_click_fullscreen()
        last_keepalive = datetime.now()
        was_active = True

    elif active and was_active:
        if datetime.now() - last_keepalive > timedelta(minutes=5):
            prevent_display_off()
            last_keepalive = datetime.now()

    elif not active and was_active:
        print("[-] AirPlay 해제됨")
        play_beep(700, 300)
        lower_airserver_window()
        clear_execution_state()
        was_active = False

    time.sleep(1)  # 1초 폴링으로 반응 속도 ↑
