# sheetkit · 表格合并 / 双表对账 / 分组汇总

> 一个零依赖（纯 Python 标准库）的小工具：把「十几张表合成一张」「订单对收款」「按门店汇总」这类活，
> 从一下午压缩到几十秒。读 csv 和 xlsx，输出带样式的多 sheet Excel。

## 为什么开源

因为你会自己配 Python 环境，那就不该花钱。免费版给你源码和命令行；
不想折腾环境的人，可以买打包好的 Windows 双击版（含图形界面和模板）——**收费的是「省事」，不是「功能」**。

## 安装

```bash
git clone git@github.com:3shisan3/sheetkit.git
cd sheetkit
pip install openpyxl   # 只有读 xlsx 才需要
```

Windows 用户也可以直接双击 `install.bat`。

## 用法

```bash
# 多表合并（列名不一致就加映射）
python tool/sheetkit.py merge --inputs A.csv B.csv C.csv \
    --map 门店=店名 --map 营业额=销售额 --dedupe-by 单号 --out-dir out

# 双表对账（订单 vs 收款）
python tool/sheetkit.py recon --left 订单.xlsx --right 流水.xlsx \
    --left-key 订单号 --right-key 订单编号 \
    --left-amount 应收金额 --right-amount 到账金额 --out-dir out --xlsx

# 分组汇总（按地区、按月）
python tool/sheetkit.py report --input data.csv --group 地区 --value 金额 --out-dir out
```

## 输出

- `summary.md`：人话结论（匹配多少、差多少、问题在哪）
- `matched.csv` / `amount_diff.csv` / `only_left.csv` / `only_right.csv` / `dup_keys.csv`：明细
- `--xlsx`：把上面全部塞进一个带表头样式的 Excel，直接发给客户

## 已知的坑（也是它的价值所在）

- 重复单号会让金额翻倍 → 工具会单独出 `dup_keys.csv`，不静默合并
- 「¥1,234.00元」「1,234」「 1234 」都是同一个数 → 内置金额清洗
- 列名不一致（门店/店名、营业额/销售额）→ `--map` 一次对齐
- 日期格式混着来 → 自动识别 6 种格式

## 授权

MIT。随便用，包括商用。
