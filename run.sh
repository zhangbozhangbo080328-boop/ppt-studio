#!/bin/sh
set -e
cd "$(dirname "$0")"

PY=""
if [ -x ".venv/bin/python" ]; then
  PY=".venv/bin/python"
elif [ -x "../github-projects/ppt-master/.venv/bin/python" ]; then
  PY="../github-projects/ppt-master/.venv/bin/python"
elif command -v python3 >/dev/null 2>&1; then
  echo "正在创建本地虚拟环境并安装依赖…"
  python3 -m venv .venv
  .venv/bin/pip install -r requirements.txt
  PY=".venv/bin/python"
else
  echo "未找到 Python，请先安装 Python 3.9+。" >&2
  exit 1
fi

echo "启动 PPT Studio: http://127.0.0.1:${PORT:-5000}"
exec "$PY" app.py
