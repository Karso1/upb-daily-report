#!/bin/zsh
set -u
cd -- "${0:A:h}" || exit 1

report="UPB每日数据汇总.xlsx"
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"
python="${UPB_PYTHON:-}"
if [[ -z "$python" ]]; then
  if [[ -x .venv/bin/python ]]; then
    python="./.venv/bin/python"
  else
    python="python3"
  fi
fi

if ! "$python" -c 'import openpyxl, xlsxwriter' >/dev/null 2>&1; then
  print -u2 '找不到 Python 或日报依赖。请按 README 创建 .venv 并安装 requirements.txt。'
  [[ ! -t 0 ]] || read -rs '?按回车关闭窗口…'
  exit 1
fi

"$python" refresh.py "$@"
result=$?
if [[ $result -eq 0 && -f "$report" ]]; then
  open "$report"
elif [[ -t 0 ]]; then
  read -rs '?按回车关闭窗口…'
fi
exit "$result"
