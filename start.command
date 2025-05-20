#!/bin/bash
cd "$(dirname "$0")"

# 检查是否安装了必要的依赖
if ! command -v pip3 &> /dev/null; then
    echo "请先安装 Python3 和 pip3"
    exit 1
fi

# 安装依赖
pip3 install -r requirements.txt

# 启动服务器
python3 app.py