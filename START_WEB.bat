@echo off
setlocal EnableExtensions EnableDelayedExpansion

REM ================================
REM  Tez Kontrol - Tek Tik Baslat
REM ================================

title Tez Kontrol - Baslatici

REM 1) Script’in oldugu klasore gec
cd /d "%~dp0"

echo ==========================================
echo   TEZ KONTROL SISTEMI BASLATICI
echo   Klasor: %cd%
echo ==========================================

REM 2) Gerekli dosyalar var mi kontrol et
if not exist "app.py" (
  echo [HATA] app.py bulunamadi. Bu .bat dosyasi web_version klasorunde olmali.
  pause
  exit /b 1
)

if not exist "static\index.html" (
  echo [HATA] static\index.html bulunamadi. static klasoru ve index.html kontrol edin.
  pause
  exit /b 1
)

REM 3) Python var mi?
python --version >nul 2>&1
if errorlevel 1 (
  echo [HATA] Python bulunamadi. Python kurulu degil veya PATH'e ekli degil.
  echo Cozum: python.org’dan kurun ve "Add Python to PATH" secin.
  pause
  exit /b 1
)

REM 4) Sanal ortam varsa aktif et (opsiyonel ama iyi pratik)
if exist "venv\Scripts\activate.bat" (
  echo [OK] venv bulundu, aktif ediliyor...
  call "venv\Scripts\activate.bat"
) else (
  echo [BILGI] venv bulunamadi. (Istersen venv kurabiliriz.)
)

REM 5) uvicorn var mi? Yoksa kur
python -c "import uvicorn" >nul 2>&1
if errorlevel 1 (
  echo [BILGI] uvicorn yok. Yukleniyor...
  pip install uvicorn fastapi python-multipart >nul 2>&1
  if errorlevel 1 (
    echo [HATA] Paket kurulumu basarisiz. Internet/PIP ayarlarinizi kontrol edin.
    pause
    exit /b 1
  )
  echo [OK] Paketler yuklendi.
) else (
  echo [OK] uvicorn hazir.
)

REM 6) Port 8000 dolu mu? (Doluysa uyar)
for /f "tokens=5" %%a in ('netstat -ano ^| findstr /r /c:":8000 .*LISTENING"') do (
  set PID8000=%%a
)
if defined PID8000 (
  echo [UYARI] Port 8000 zaten kullaniliyor. PID: %PID8000%
  echo Bu durumda yeni sunucu baslamayabilir.
  echo.
  choice /m "Bu PID'i sonlandirmak ister misin"
  if errorlevel 2 (
    echo [BILGI] PID sonlandirilmadi. Mevcut sunucu calisiyorsa tarayiciyi aciyorum.
    goto OPEN_BROWSER
  ) else (
    echo [BILGI] PID %PID8000% sonlandiriliyor...
    taskkill /PID %PID8000% /F >nul 2>&1
    timeout /t 1 >nul
  )
)

REM 7) Sunucuyu baslat (ayri pencerede)
echo [BILGI] Sunucu baslatiliyor (reload acik)...
start "Tez Kontrol Server" cmd /k "cd /d %cd% && uvicorn app:app --reload --host 127.0.0.1 --port 8000"

REM 8) Biraz bekle
timeout /t 2 >nul

:OPEN_BROWSER
REM 9) Tarayiciyi ac
echo [BILGI] Tarayici aciliyor...
start "" "http://127.0.0.1:8000/?v=%random%"

echo [OK] Hazir. Sunucu penceresini kapatmak icin o pencerede CTRL+C kullan.
echo.
pause
endlocal
