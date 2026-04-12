@echo off
title Discord Order Bot
chcp 65001 >nul

:: เช็ค Node.js
node --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] ไม่พบ Node.js กรุณารัน install.bat ก่อน
    pause
    exit /b 1
)

:: เช็ค node_modules
if not exist "node_modules" (
    echo [!] ยังไม่ได้ติดตั้ง packages กำลังติดตั้ง...
    npm install
)

echo ===================================
echo   Discord Order Bot
echo ===================================
echo.
node index.js

echo.
echo === บอทหยุดทำงาน ===
pause
