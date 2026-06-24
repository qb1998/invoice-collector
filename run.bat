@echo off
REM 发票收集器 - Windows 启动脚本
REM 使用方法: 双击 run.bat 或在命令行运行 run.bat

cd /d "%~dp0"

REM 设置 no_proxy 让 localhost 请求不走代理
set no_proxy=localhost,127.0.0.1
set NO_PROXY=localhost,127.0.0.1

REM 检查虚拟环境
if not exist ".venv" (
    echo 首次运行，正在创建虚拟环境...
    python -m venv .venv
    .venv\Scripts\pip install --upgrade pip
    .venv\Scripts\pip install fastapi uvicorn pdfplumber openpyxl python-multipart
)

REM 检查依赖
.venv\Scripts\python -c "import fastapi, uvicorn, pdfplumber, openpyxl" 2>nul
if errorlevel 1 (
    echo 正在安装依赖...
    .venv\Scripts\pip install fastapi uvicorn pdfplumber openpyxl python-multipart
)

echo ========================================
echo   发票收集器已启动
echo   浏览器打开: http://localhost:8000
echo   按 Ctrl+C 停止服务
echo ========================================

.venv\Scripts\python app.py
