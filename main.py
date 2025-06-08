import subprocess
import time
import locale
import winsound
import ctypes
from datetime import datetime, timedelta

import pyautogui
import pygetwindow as gw
import win32gui
import win32con
import keyboard   # pip install keyboard

# ─────────── Win32 디스플레이 절전 플래그 ───────────
ES_CONTINUOUS       = 0x80000000
ES_DISPLAY_REQUIRED = 0x00000002  # 디스플레이 유휴 타이머만 리셋

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
    state = "ON" if AUTO_FULLSCREEN else "OFF"
    print(f"[🔀] Auto-fullscreen toggled: {state} ({datetime.now()})")

keyboard.add_hotkey('ctrl+alt+f', toggle_fullscreen_flag)

# ─────────── AirPlay 감시 설정 ───────────
PORTS_TO_WATCH   = [5000, 7000]
PROCESS_HINT     = "AirServer.exe"
MIN_CONNECTIONS  = 1

def count_established_sessions():
    out = subprocess.check_output('netstat -an', shell=True)\
           .decode(locale.getpreferredencoding(), errors='replace')
    return sum(
        1
        for line in out.splitlines()
        if any(f":{p}" in line for p in PORTS_TO_WATCH)
        and "ESTABLISHED" in line
    )

def is_airplay_active():
    tl = subprocess.check_output('tasklist', shell=True)\
         .decode(locale.getpreferredencoding(), errors='replace')
    return (PROCESS_HINT.lower() in tl.lower()
            and count_established_sessions() >= MIN_CONNECTIONS)

# ─────────── 글로벌 저장 ───────────
airserver_hwnd = None

# ─────────── 유틸리티 ───────────
def play_beep(freq, dur):
    winsound.Beep(freq, dur)

def jiggle_mouse():
    x, y = pyautogui.position()
    pyautogui.moveTo(x + 1, y)
    pyautogui.moveTo(x, y)
    print(f"[🖱️] jiggle_mouse: {datetime.now()}")

def raise_airserver_window():
    global airserver_hwnd
    target = "AirServer Windows Desktop Edition"
    wins = [w for w in gw.getAllWindows() if target in w.title]
    if not wins:
        print("[!] AirServer 창을 찾을 수 없음")
        airserver_hwnd = None
        return

    win = wins[0]
    airserver_hwnd = win._hWnd

    if win.isMinimized:
        win.restore()

    win32gui.ShowWindow(airserver_hwnd, win32con.SW_SHOWNORMAL)
    win32gui.BringWindowToTop(airserver_hwnd)
    try:
        win32gui.SetForegroundWindow(airserver_hwnd)
    except Exception as e:
        print("SetForegroundWindow 실패:", e)
    print(f"[🔝] raise_airserver_window: {datetime.now()}")

def double_click_fullscreen():
    global airserver_hwnd
    if not airserver_hwnd or not win32gui.IsWindow(airserver_hwnd):
        print("[!] double_click_fullscreen: 유효한 AirServer 창 핸들 없음")
        return

    left, top, right, bottom = win32gui.GetWindowRect(airserver_hwnd)
    cx = (left + right) // 2
    cy = (top + bottom) // 2
    pyautogui.moveTo(cx, cy)
    pyautogui.doubleClick()
    print(f"[⛶] double_click_fullscreen: {datetime.now()}")

def lower_airserver_window():
    """
    전체화면 해제(필요 시) → Z-순서 맨 아래로 내립니다.
    """
    global airserver_hwnd
    if AUTO_FULLSCREEN:
        if airserver_hwnd and win32gui.IsWindow(airserver_hwnd):
            double_click_fullscreen()
            time.sleep(0.5)
        else:
            print("[!] lower: AirServer 핸들 유효하지 않아 전체화면 해제 생략")

    if not airserver_hwnd or not win32gui.IsWindow(airserver_hwnd):
        return

    win = gw.getWindowsWithTitle("AirServer Windows Desktop Edition")
    # 복구(최소화 해제) 없이 바로 Z-순서만 내립니다
    win32gui.SetWindowPos(
        airserver_hwnd,
        win32con.HWND_BOTTOM,
        0, 0, 0, 0,
        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
    )
    print(f"[🔽] lower_airserver_window: {datetime.now()}")

# ─────────── 메인 루프 ───────────
was_active     = False
last_keepalive = datetime.min

print(f"[디버그] 시작: active={is_airplay_active()}, was_active={was_active}")
while True:
    active = is_airplay_active()
    print(f"[디버그] 상태: active={active}, was_active={was_active}")

    if active and not was_active:
        print("[+] AirPlay 시작됨")
        raise_airserver_window()
        play_beep(1200, 200)
        jiggle_mouse()
        prevent_display_off()
        last_keepalive = datetime.now()

        if AUTO_FULLSCREEN:
            double_click_fullscreen()

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

    time.sleep(5)
