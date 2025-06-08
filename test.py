# ─────────── Audio-mode 볼륨 변경 로그 ───────────
# pip install pycaw comtypes

from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import threading, time, datetime

def get_airserver_volume() -> float:
    """현재 AirServer.exe 의 볼륨(0.0~1.0)을 반환."""
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        proc = session.Process
        if proc and proc.name().lower() == "airserver.exe":
            vol = session._ctl.QueryInterface(ISimpleAudioVolume)
            level = vol.GetMasterVolume()
            return level
    return None

def watch_volume_changes(poll_interval=0.5):
    last = get_airserver_volume()
    print(f"[{datetime.datetime.now():%H:%M:%S}] 초기 볼륨: {last}")
    while True:
        time.sleep(poll_interval)
        cur = get_airserver_volume()
        if cur is None:
            print(f"[{datetime.datetime.now():%H:%M:%S}] AirServer 프로세스 미발견")
            continue
        if abs(cur - last) > 0.001:
            print(f"[{datetime.datetime.now():%H:%M:%S}] 볼륨 변경 감지: {last:.2f} → {cur:.2f}")
            last = cur

# 백그라운드로 폴링 스레드 시작
t = threading.Thread(target=watch_volume_changes, daemon=True)
t.start()

# ─────────── 이하 기존 메인 루프 ───────────
# … your AirPlay 감시 & 전체화면 제어 루프 …
