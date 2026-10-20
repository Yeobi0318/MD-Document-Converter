@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

rem 사용법: build.bat          (라이브러리 설치 + 테스트 + exe 빌드)
rem         build.bat notest   (테스트 건너뛰기)

if not exist ".venv\Scripts\python.exe" (
    echo [1/4] 가상환경 만들기
    py -3.11 -m venv .venv || python -m venv .venv || goto :fail
)

echo [2/4] 라이브러리 설치
".venv\Scripts\python.exe" -m pip install -q --disable-pip-version-check -r requirements.txt || goto :fail

if /i "%~1"=="notest" goto :build
echo [3/4] 테스트
".venv\Scripts\python.exe" -m pytest -q -p no:cacheprovider tests || goto :fail

:build
echo [4/4] exe 빌드
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --log-level WARN md2doc.spec || goto :fail
echo.
echo 완료: dist 폴더에 exe가 만들어졌습니다.
exit /b 0

:fail
echo.
echo 실패했습니다. 위 메시지를 확인하세요.
pause
exit /b 1
