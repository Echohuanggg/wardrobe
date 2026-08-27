@echo off
chcp 936 >nul
title 我的衣橱 - 本地服务
cd /d "%~dp0"

echo ============================================================
echo   我的衣橱 本地服务
echo   服务地址: http://localhost:8765
echo ============================================================
echo.

rem 第零步：确定 Python 解释器（避免多Python环境解析错误）
set "PYTHON=C:\Users\echoh\AppData\Local\Programs\Python\Python312\python.exe"
if not exist "%PYTHON%" (
    where py >nul 2>&1
    if %errorlevel%==0 (
        set "PYTHON=py -3"
    ) else (
        set "PYTHON=python"
    )
)

rem 第一步：以 Excel 为权威源，自动同步 wardrobe.json 与缩略图
echo [同步] 正在从 Excel 同步衣橱数据...
"%PYTHON%" sync_excel_to_json.py
if %errorlevel% neq 0 (
    echo [警告] 数据同步未完成，继续使用旧数据启动
)

rem 第二步：检查服务是否已经在运行，是的话直接打开浏览器
netstat -ano | findstr ":8765" | findstr "LISTENING" >nul 2>&1
if %errorlevel%==0 (
    echo [提示] 服务已在运行，直接打开浏览器...
    start "" http://localhost:8765
    ping -n 3 127.0.0.1 >nul
    exit /b 0
)

echo 正在启动服务，请稍候...
start /min "" cmd /c ""%PYTHON%" serve.py"

rem 第三步：轮询等待服务就绪（最多约 15 秒，避免冷启动慢误报）
set /a attempts=0
:wait_loop
ping -n 2 127.0.0.1 >nul
netstat -ano | findstr ":8765" | findstr "LISTENING" >nul 2>&1
if %errorlevel%==0 goto service_ok
set /a attempts+=1
if %attempts% lss 15 goto wait_loop
goto service_fail

:service_ok
echo [成功] 服务启动成功，正在打开浏览器...
start "" http://localhost:8765
ping -n 3 127.0.0.1 >nul
exit /b 0

:service_fail
echo.
echo [错误] 服务启动失败，请手动运行测试：python serve.py 看报错信息
pause
exit /b 1