# AirPlay 감지 스크립트 (for AirServer)

AirServer (https://www.airserver.com/WindowsDesktop)를 사용할 때, 화면 보호기 상태에서 AirPlay를 시작해도 화면이 자동으로 깨어나지 않거나, AirServer 창이 다른 창 아래에 가려져 있어 직접 클릭해야 하는 번거로움이 있습니다.

이 스크립트는 AirPlay 연결 상태를 감지해 AirServer 창을 자동으로 맨 앞으로 올리고, 연결이 해제되면 다시 원래대로 돌려주는 기능을 수행합니다.
iPhone이나 iPad에서 미러링을 시작하면 자동으로 반응하며, 일정 주기로 마우스를 미세하게 움직여 화면 보호기 작동도 방지합니다.



## 기능
- AirPlay 연결 상태 감지
- AirPlay연결 감지시 -> AirServer 창을 항상 위로 설정 / 해제(아직 미작동)
- 연결 시 소리로 알림(수정 후 삭제 예정)
- 5분마다 자동으로 마우스를 미세하게 움직여 화면 꺼짐 방지

## 사용 환경

- Windows 10 이상
- Python 3.x
- Microsoft Store에서 설치한 AirServer 사용 중

## 설치 방법

1. 필요한 패키지 설치:
pip install pygetwindow pywin32 