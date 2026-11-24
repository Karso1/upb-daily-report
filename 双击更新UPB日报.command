#!/bin/zsh
set -u
cd -- "${0:A:h}" || exit 1

report="UPB每日数据汇总.xlsx"
python="${UPB_PYTHON:-python3}"

"$python" refresh.py "$@"
status=$?
if [[ $status -eq 0 && -f "$report" ]]; then
  open "$report"
elif [[ -t 0 ]]; then
  read -rs '?按回车关闭窗口…'
fi
exit "$status"
