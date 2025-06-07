import subprocess
import time
import os
import winsound
import locale
import pyautogui
from datetime import datetime, timedelta
import pygetwindow as gw
import win32gui
import win32con



PORTS_TO_WATCH = [5000, 7000]
PROCESS_HINT = "AirServer.exe"
MIN_CONNECTIONS = 1

def count_established_sessions():
    try:
        output = subprocess.check_output('netstat -an', shell=True).decode(locale.getpreferredencoding(), errors='replace')
        return sum(1 for line in output.splitlines()
                   if any(f":{port}" in line for port in PORTS_TO_WATCH) and "ESTABLISHED" in line)
    except Exception as e:
        print("에러 발생:", e)
        return 0

def is_airplay_active():
    try:
        tasklist = subprocess.check_output('tasklist', shell=True).decode(locale.getpreferredencoding(), errors='replace')
        if PROCESS_HINT.lower() not in tasklist.lower():
            return False
        return count_established_sessions() >= MIN_CONNECTIONS
    except Exception as e:
        print("에러 발생:", e)
        return False

def play_beep(freq=1000, duration=300):
    winsound.Beep(freq, duration)

def jiggle_mouse():
    try:
        x, y = pyautogui.position()
        pyautogui.moveTo(x + 1, y)
        pyautogui.moveTo(x, y)
        print(f"[🖱️] 마우스 움직임: {datetime.now()}")
    except Exception as e:
        print("마우스 이동 실패:", e)

def raise_airserver_window():
    try:
        for win in gw.getWindowsWithTitle('AirServer'):
            if win.isMinimized:
                win.restore()
            hwnd = win._hWnd
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            win32gui.SetForegroundWindow(hwnd)
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            print("[🔝] AirServer 창을 최상단으로 올림")
            return
        print("[!] AirServer 창을 찾지 못함")
    except Exception as e:
        print("AirServer 최상단 올리기 실패:", e)

def untop_airserver_window():
    try:
        for win in gw.getWindowsWithTitle('AirServer'):
            hwnd = win._hWnd
            win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            print("[⬇️] AirServer 창을 일반 창으로 내림")
            return
    except Exception as e:
        print("AirServer 창 내리기 실패:", e)



# 상태 변수
was_active = False
last_mouse_jiggle = datetime.min

while True:
    active = is_airplay_active()

    if active and not was_active:
        
        print("[+] AirPlay 시작됨")
        raise_airserver_window()

        play_beep(1200, 200)
        jiggle_mouse()  # 최초 한 번
        last_mouse_jiggle = datetime.now()
        was_active = True

    elif active and was_active:
        # 5분마다 마우스 살짝 움직임
        if datetime.now() - last_mouse_jiggle > timedelta(minutes=5):
            jiggle_mouse()
            last_mouse_jiggle = datetime.now()

    elif not active and was_active:
        
        print("[-] AirPlay 해제됨")
        untop_airserver_window()
        play_beep(700, 300)
        was_active = False

    time.sleep(5)

