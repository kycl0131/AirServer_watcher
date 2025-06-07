import subprocess
import time
import locale
import winsound
from datetime import datetime, timedelta

import pyautogui
import pygetwindow as gw
import win32gui
import win32con

# 감시할 포트와 프로세스 정보
PORTS_TO_WATCH = [5000, 7000]
PROCESS_HINT = "AirServer.exe"
MIN_CONNECTIONS = 1

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
    if PROCESS_HINT.lower() not in tl.lower():
        return False
    return count_established_sessions() >= MIN_CONNECTIONS

def play_beep(freq, dur):
    winsound.Beep(freq, dur)

def jiggle_mouse():
    x, y = pyautogui.position()
    pyautogui.moveTo(x + 1, y)
    pyautogui.moveTo(x, y)
    print(f"마우스 움직임: {datetime.now()}")

def raise_airserver_window():
    target = "AirServer Windows Desktop Edition"
    wins = [w for w in gw.getAllWindows() if target in w.title]
    if not wins:
        print("[!] AirServer 창을 찾을 수 없음")
        return

    win = wins[0]
    hwnd = win._hWnd

    # 최소화 해제
    if win.isMinimized:
        win.restore()

    # 정상 보이기 + Z-순서 앞으로
    win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
    win32gui.BringWindowToTop(hwnd)

    # 포커스 주기
    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception as e:
        print("SetForegroundWindow 실패:", e)

    print("AirServer 창을 포그라운드로 올림")

def lower_airserver_window():
    target = "AirServer Windows Desktop Edition"
    wins = [w for w in gw.getAllWindows() if target in w.title]
    if not wins:
        return

    hwnd = wins[0]._hWnd
    # 창이 최소화 되어 있으면 복구
    if wins[0].isMinimized:
        wins[0].restore()
    # Z-순서 맨 아래로
    win32gui.SetWindowPos(
        hwnd,
        win32con.HWND_BOTTOM,
        0, 0, 0, 0,
        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
    )
    print("AirServer 창을 Z-순서 맨 아래로 내림")

# 메인 루프
was_active = False
last_mouse_jiggle = datetime.min

print(f"[디버그] active={is_airplay_active()}, was_active={was_active}")
while True:
    active = is_airplay_active()
    print(f"[디버그] active={active}, was_active={was_active}")

    if active and not was_active:
        print("[+] AirPlay 시작됨")
        raise_airserver_window()
        play_beep(1200, 200)
        jiggle_mouse()
        last_mouse_jiggle = datetime.now()
        was_active = True

    elif active and was_active:
        if datetime.now() - last_mouse_jiggle > timedelta(minutes=5):
            jiggle_mouse()
            last_mouse_jiggle = datetime.now()

    elif not active and was_active:
        print("[-] AirPlay 해제됨")
        play_beep(700, 300)
        lower_airserver_window()
        was_active = False

    time.sleep(5)
