@echo off
title Discord Order Bot - Installer
chcp 65001 >nul
echo ===================================
echo   Discord Order Bot - ติดตั้ง
echo ===================================
echo.

:: เช็ค Node.js
node --version >nul 2>&1
if %errorlevel% == 0 (
    echo [OK] Node.js พบแล้ว:
    node --version
    goto :install_deps
)

echo [!] ไม่พบ Node.js กำลังดาวน์โหลด...
echo.

:: ดาวน์โหลด Node.js LTS installer
set NODE_URL=https://nodejs.org/dist/v20.19.1/node-v20.19.1-x64.msi
set NODE_MSI=%TEMP%\node-installer.msi

echo กำลังดาวน์โหลด Node.js LTS...
powershell -Command "Invoke-WebRequest -Uri '%NODE_URL%' -OutFile '%NODE_MSI%' -UseBasicParsing"
if %errorlevel% neq 0 (
    echo [ERROR] ดาวน์โหลดล้มเหลว กรุณาติดตั้ง Node.js เอง:
    echo https://nodejs.org
    pause
    exit /b 1
)

echo กำลังติดตั้ง Node.js...
msiexec /i "%NODE_MSI%" /quiet /norestart
if %errorlevel% neq 0 (
    echo [ERROR] ติดตั้งล้มเหลว กรุณาติดตั้ง Node.js เอง:
    echo https://nodejs.org
    pause
    exit /b 1
)

:: Refresh PATH
call RefreshEnv.cmd >nul 2>&1
set "PATH=%PATH%;%ProgramFiles%\nodejs"

echo [OK] ติดตั้ง Node.js เสร็จแล้ว

:install_deps
echo.
echo กำลังติดตั้ง packages...
npm install
if %errorlevel% neq 0 (
    echo [ERROR] npm install ล้มเหลว
    pause
    exit /b 1
)

echo.
echo ===================================
echo   ติดตั้งเสร็จแล้ว!
echo ===================================
echo.
echo กดดับเบิ้ลคลิก "เปิดบอท.bat" เพื่อเปิดบอท
echo.
pause
