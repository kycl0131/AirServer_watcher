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
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | ES_DISPLAY_REQUIRED
        )
        logger.info("⏱️ Display sleep prevented")
    except Exception as e:
        logger.error(f"Failed to prevent display off: {e}")

def clear_execution_state():
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
        logger.info("↩️ Execution state cleared")
    except Exception as e:
        logger.error(f"Failed to clear execution state: {e}")

# ─────────── Auto-fullscreen 토글 ───────────
AUTO_FULLSCREEN = True
def toggle_fullscreen_flag():
    global AUTO_FULLSCREEN
    AUTO_FULLSCREEN = not AUTO_FULLSCREEN
    logger.info(f"🔀 Auto-fullscreen: {'ON' if AUTO_FULLSCREEN else 'OFF'}")

try:
    keyboard.add_hotkey('ctrl+alt+f', toggle_fullscreen_flag)
except Exception as e:
    logger.error(f"Failed to register hotkey: {e}")

# ─────────── AirPlay 감시 설정 ───────────
PROCESS_NAME     = "AirServer.exe"
WATCH_PORTS      = {5000, 7000}

def count_established_sessions():
    """psutil로 ESTABLISHED된 TCP 연결 중 WATCH_PORTS에 속한 것을 센다."""
    try:
        cnt = 0
        video_sessions = 0
        
        for conn in psutil.net_connections(kind='tcp'):
            if conn.status == psutil.CONN_ESTABLISHED and conn.laddr.port in WATCH_PORTS:
                cnt += 1
                if conn.laddr.port == 7000:  # 비디오 포트
                    video_sessions += 1
                    
        logger.debug(f"Established sessions on {WATCH_PORTS}: {cnt} (video: {video_sessions})")
        return cnt, video_sessions
    except Exception as e:
        logger.error(f"Failed to count sessions: {e}")
        return 0, 0

def is_airplay_active():
    """AirServer 프로세스 존재 + 최소 1개 이상의 연결이 있으면 활성."""
    proc_found = False
    
    try:
        for p in psutil.process_iter(['name']):
            if p.info.get('name', '').lower() == PROCESS_NAME.lower():
                proc_found = True
                break
    except Exception as e:
        logger.error(f"Failed to check processes: {e}")

    sessions, video_sessions = count_established_sessions()
    active = proc_found and sessions >= 1
    is_video = video_sessions > 0
    
    logger.debug(f"Process found: {proc_found}, Sessions: {sessions}, Video: {is_video} -> Active: {active}")
    return active, is_video

# ─────────── 스크린세이버 체크 ───────────
SPI_GETSCREENSAVERRUNNING = 114
def is_screensaver_running():
    running = ctypes.c_int()
    ctypes.windll.user32.SystemParametersInfoW(
        SPI_GETSCREENSAVERRUNNING, 0,
        ctypes.byref(running), 0
    )
    return bool(running.value)

def wait_for_screensaver_off(timeout=5, poll=0.2):
    start = time.time()
    while time.time() - start < timeout:
        if not is_screensaver_running():
            return True
        time.sleep(poll)
    return False

# ─────────── 창 핸들 찾기/강제 포그라운드 ───────────
def find_airserver_window():
    target = "AirServer Windows Desktop Edition"
    try:
        wins = [w for w in gw.getAllWindows() if target in w.title]
        if wins:
            return wins[0]._hWnd
    except Exception as e:
        logger.error(f"Failed to find AirServer window: {e}")
    return None

def force_set_foreground(hwnd):
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32

    fg_hwnd = user32.GetForegroundWindow()
    cur_tid = kernel32.GetCurrentThreadId()                         # ← kernel32에서 가져오기
    fg_tid = user32.GetWindowThreadProcessId(fg_hwnd, None)

    # 입력 스레드를 붙였다 떼기
    user32.AttachThreadInput(cur_tid, fg_tid, True)
    user32.SetForegroundWindow(hwnd)
    user32.AttachThreadInput(cur_tid, fg_tid, False)


# ─────────── 글로벌 상태 저장 ───────────
airserver_hwnd  = None
last_was_video  = False

# ─────────── 유틸리티 ───────────
def play_beep(freq, dur):
    try:
        winsound.Beep(freq, dur)
    except Exception as e:
        logger.error(f"Failed to play beep: {e}")

def jiggle_mouse():
    try:
        x, y = pyautogui.position()
        screen_w, screen_h = pyautogui.size()
        if 0 < x < screen_w-1 and 0 < y < screen_h-1:
            pyautogui.moveTo(x+1, y); time.sleep(0.1); pyautogui.moveTo(x, y)
            logger.info("🖱️ Mouse jiggled")
        else:
            logger.warning("Mouse at edge, skipping jiggle")
    except Exception as e:
        logger.error(f"Failed to jiggle mouse: {e}")

def move_mouse_away_from_airserver():
    global airserver_hwnd
    if not airserver_hwnd or not win32gui.IsWindow(airserver_hwnd):
        return
    try:
        left, top, right, bottom = win32gui.GetWindowRect(airserver_hwnd)
        cx, cy = pyautogui.position()
        if left <= cx <= right and top <= cy <= bottom:
            sw, sh = pyautogui.size()
            safe_x = max(10, min(sw-50, right+50))
            safe_y = max(10, min(sh-50, bottom+50))
            pyautogui.moveTo(safe_x, safe_y)
            logger.info(f"🖱️ Mouse moved to ({safe_x},{safe_y})")
    except Exception as e:
        logger.error(f"Failed to move mouse away: {e}")

def raise_airserver_window():
    global airserver_hwnd

    # 유효 핸들 아니면 다시 찾기
    if not airserver_hwnd or not win32gui.IsWindow(airserver_hwnd):
        airserver_hwnd = find_airserver_window()
    if not airserver_hwnd:
        logger.warning("AirServer window not found")
        return False

    # 화면보호기 동작 중일 때만 대기
    if is_screensaver_running():
        logger.info("Screensaver running, waiting to turn off…")
        if wait_for_screensaver_off():
            logger.info("Screensaver off")
        else:
            logger.warning("Timeout waiting for screensaver off")

    try:
        if win32gui.IsIconic(airserver_hwnd):
            win32gui.ShowWindow(airserver_hwnd, win32con.SW_RESTORE)
            time.sleep(0.2)

        win32gui.ShowWindow(airserver_hwnd, win32con.SW_SHOWNORMAL)
        win32gui.BringWindowToTop(airserver_hwnd)

        try:
            win32gui.SetForegroundWindow(airserver_hwnd)
        except Exception:
            force_set_foreground(airserver_hwnd)

        logger.info("🔝 AirServer window raised")
        return True
    except Exception as e:
        logger.error(f"Failed to raise AirServer window: {e}")
        airserver_hwnd = None
        return False

def double_click_fullscreen():
    if not airserver_hwnd or not win32gui.IsWindow(airserver_hwnd):
        logger.warning("No valid hwnd for fullscreen")
        return
    try:
        l, t, r, b = win32gui.GetWindowRect(airserver_hwnd)
        pyautogui.moveTo((l+r)//2, (t+b)//2)
        pyautogui.doubleClick()
        logger.info("⛶ double-click fullscreen")
    except Exception as e:
        logger.error(f"Failed to double-click fullscreen: {e}")

def lower_airserver_window():
    global last_was_video
    if AUTO_FULLSCREEN and last_was_video:
        double_click_fullscreen()
        time.sleep(0.5)
    if airserver_hwnd and win32gui.IsWindow(airserver_hwnd):
        try:
            win32gui.SetWindowPos(
                airserver_hwnd, win32con.HWND_BOTTOM,
                0,0,0,0,
                win32con.SWP_NOMOVE|win32con.SWP_NOSIZE
            )
            logger.info("🔽 AirServer window lowered")
        except Exception as e:
            logger.error(f"Failed to lower window: {e}")

# ─────────── 메인 루프 ───────────
def main():
    global last_was_video
    was_active = False
    last_keepalive = datetime.min
    consecutive_errors = 0
    MAX_ERRORS = 5

    logger.info("🚀 AirPlay Monitor started")

    while True:
        try:
            active, is_video = is_airplay_active()

            if active and not was_active:
                last_was_video = is_video
                logger.info(f"+ AirPlay session started (video={is_video})")

                if raise_airserver_window():
                    time.sleep(0.5)
                    jiggle_mouse()
                    prevent_display_off()
                    if AUTO_FULLSCREEN and last_was_video:
                        time.sleep(1.0)
                        double_click_fullscreen()
                        time.sleep(0.5)
                        move_mouse_away_from_airserver()

                last_keepalive = datetime.now()
                was_active = True
                consecutive_errors = 0

            elif active and was_active:
                if is_video:
                    move_mouse_away_from_airserver()
                if datetime.now() - last_keepalive > timedelta(minutes=5):
                    prevent_display_off()
                    last_keepalive = datetime.now()

            elif not active and was_active:
                logger.info("- AirPlay session ended")
                lower_airserver_window()
                clear_execution_state()
                was_active = False
                consecutive_errors = 0

            time.sleep(1)

        except KeyboardInterrupt:
            logger.info("🛑 Shutting down by user request")
            break
        except Exception as e:
            consecutive_errors += 1
            logger.error(f"Main loop error ({consecutive_errors}/{MAX_ERRORS}): {e}")
            if consecutive_errors >= MAX_ERRORS:
                logger.critical("Too many errors, exiting")
                break
            time.sleep(5)

    # 종료 시 정리
    try:
        clear_execution_state()
        logger.info("👋 AirPlay Monitor stopped")
    except Exception as e:
        logger.error(f"Cleanup error: {e}")

if __name__ == "__main__":
    main()
