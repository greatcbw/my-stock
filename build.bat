@echo off
chcp 65001 >nul
echo ============================================
echo  트레이딩 PRO+ 빌드 스크립트
echo ============================================
echo.

:: Python 버전 확인
python --version
echo.

:: PyInstaller 설치/업데이트
echo [1/4] PyInstaller 최신버전 설치 중...
pip install --upgrade pyinstaller
echo.

:: 필수 패키지 확인
echo [2/4] 필수 패키지 확인 중...
pip install flask flask-cors yfinance FinanceDataReader requests anthropic python-dotenv
echo.

:: 이전 빌드 정리
echo [3/4] 이전 빌드 정리 중...
if exist "dist" rmdir /s /q dist
if exist "build" rmdir /s /q build
echo.

:: 빌드 실행
echo [4/4] 빌드 시작...
pyinstaller app.spec

echo.
if exist "dist\트레이딩PRO\트레이딩PRO.exe" (
    echo ============================================
    echo  빌드 성공!
    echo  실행파일: dist\트레이딩PRO\트레이딩PRO.exe
    echo ============================================
    echo.
    echo  배포 시 dist\트레이딩PRO\ 폴더 전체를 복사하세요.
) else (
    echo ============================================
    echo  빌드 실패 - 위 오류 메시지를 확인하세요
    echo ============================================
)
pause
