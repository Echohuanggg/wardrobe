@echo off
chcp 936 >nul
title 我的衣橱 - 本地服务
cd /d "%~dp0"

echo ============================================================
echo   我的衣橱 本地服务
echo   服务地址: http://localhost:8765
echo ============================================================
echo.

rem 第一步：检查服务是否已经在运行，是的话直接打开浏览器
netstat -ano | findstr ":8765" | findstr "LISTENING" >nul 2>&1
if %errorlevel%==0 (
    echo [提示] 服务已在运行，直接打开浏览器...
    start "" http://localhost:8765
    timeout /t 3 /nobreak >nul
    exit /b 0
)

rem 第二步：用绝对路径的 Python 启动服务（避免多Python环境解析错误）
set "PYTHON=C:\Users\echoh\AppData\Local\Programs\Python\Python312\python.exe"
if not exist "%PYTHON%" (
    where py >nul 2>&1
    if %errorlevel%==0 (
        set "PYTHON=py -3"
    ) else (
        set "PYTHON=python"
    )
)

echo 正在启动服务，请稍候...
start /min "" cmd /c ""%PYTHON%" serve.py"
timeout /t 3 /nobreak >nul

rem 第三步：验证服务是否真的启动成功
netstat -ano | findstr ":8765" | findstr "LISTENING" >nul 2>&1
if %errorlevel%==0 (
    echo [成功] 服务启动成功，正在打开浏览器...
    start "" http://localhost:8765
) else (
    echo.
    echo [错误] 服务启动失败，请手动运行测试：python serve.py 看报错信息
    pause
)
timeout /t 3 /nobreak >nul
