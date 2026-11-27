@echo off
chcp 65001 >nul
echo 正在安装依赖（只有读 xlsx 才需要）...
python -m pip install --quiet openpyxl
echo 完成。试试：
echo   python tool\sheetkit.py recon --left A.xlsx --right B.xlsx --left-key 订单号 --right-key 订单号 --left-amount 金额 --right-amount 金额 --out-dir out --xlsx
pause