#!/bin/bash
# 发票收集器 - macOS/Linux 启动脚本
# 使用方法: ./run.sh

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

# 设置 no_proxy 让 localhost 请求不走代理
export no_proxy="localhost,127.0.0.1"
export NO_PROXY="localhost,127.0.0.1"

# 检查虚拟环境
if [ ! -d ".venv" ]; then
    echo "首次运行，正在创建虚拟环境..."
    python3 -m venv .venv
    .venv/bin/pip install --upgrade pip
    .venv/bin/pip install fastapi uvicorn pdfplumber openpyxl python-multipart
fi

# 检查依赖
.venv/bin/python -c "import fastapi, uvicorn, pdfplumber, openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "正在安装依赖..."
    .venv/bin/pip install fastapi uvicorn pdfplumber openpyxl python-multipart
fi

echo "========================================"
echo "  发票收集器 v2.0 已启动"
echo "  浏览器打开: http://localhost:8000"
echo "  按 Ctrl+C 停止服务"
echo "========================================"

.venv/bin/python app.py
