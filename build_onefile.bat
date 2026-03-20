@echo off
chcp 65001 >nul
echo ============================================
echo  트레이딩 PRO+ 단일파일 빌드 스크립트
echo ============================================
echo.

python --version
echo.

echo [1/4] PyInstaller 최신버전 설치 중...
py -m  install --upgrade pyinstaller
echo.

echo [2/4] 필수 패키지 확인 중...
py -m  install flask flask-cors yfinance FinanceDataReader requests anthropic python-dotenv websocket-client
echo.

echo [3/4] 이전 빌드 정리 중...
if exist "dist" rmdir /s /q dist
if exist "build" rmdir /s /q build
echo.

echo [4/4] 단일 파일 빌드 시작...
python -m PyInstaller --onefile --console ^
  --collect-all FinanceDataReader ^
  --collect-all flask ^
  --collect-all yfinance ^
  --collect-all anthropic ^
  --hidden-import=flask_cors ^
  --hidden-import=dotenv ^
  --hidden-import=lxml ^
  --hidden-import=lxml.etree ^
  --hidden-import=lxml._elementpath ^
  --hidden-import=bs4 ^
  --hidden-import=websocket ^
  --hidden-import=concurrent.futures ^
  --hidden-import=concurrent.futures.thread ^
  --add-data "stock_chart_app.html;." ^
  --add-data "*.csv;." ^
  --add-data ".env;." ^
  --name 트레이딩PRO ^
  app.py

echo.
if exist "dist\트레이딩PRO.exe" (
    echo ============================================
    echo  빌드 성공!
    echo  실행파일: dist\트레이딩PRO.exe
    echo ============================================
    echo.
    echo  아래 파일들을 exe 와 같은 폴더에 두세요:
    echo    - .env
    echo    - stock_chart_app.html
    echo    - 한국상장법인.csv
    echo    - 미국snp.csv
    echo    - 미국Nyse.csv
    echo    - ETF.csv
    echo    - ETN.csv
) else (
    echo ============================================
    echo  빌드 실패 - 위 오류 메시지를 확인하세요
    echo ============================================
)
pause