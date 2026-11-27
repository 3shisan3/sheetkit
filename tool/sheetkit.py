#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""sheetkit - 表格清洗 / 多表合并 / 双表对账 / 汇总报表

只依赖标准库；读 .xlsx 需要 openpyxl（可选）。
输出文件均为 UTF-8-SIG，Excel 双击直接不乱码。

用法示例:
  python sheetkit.py merge --inputs a.csv b.csv --out-dir out --map 门店=店名
  python sheetkit.py recon --left 订单.csv --right 流水.csv --left-key 订单号 --right-key 订单编号 --left-amount 应收金额 --right-amount 到账金额 --out-dir out
  python sheetkit.py report --input data.csv --group 地区 --value 金额 --out-dir out
"""
import argparse
import csv
import os
import re
import sys
from datetime import datetime
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

FW = str.maketrans("０１２３４５６７８９（）．，－／：％￥", "0123456789().,-/:%Y")
DATE_FMTS = ["%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d",
             "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"]


def clean_text(v):
    if v is None:
        return ""
    s = str(v).strip().translate(FW)
    return re.sub(r"\s+", " ", s)


def norm_key(v):
    return clean_text(v).upper()


def to_amount(v):
    s = clean_text(v)
    if s == "":
        return None
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg, s = True, s[1:-1]
    for ch in ("¥", "Y", "元", ",", " ", "_"):
        s = s.replace(ch, "")
    if s.startswith("-"):
        neg, s = True, s[1:]
    if s.endswith("-"):
        neg, s = True, s[:-1]
    if not re.fullmatch(r"\d+(\.\d+)?", s):
        return None
    val = float(s)
    return -val if neg else val


def to_date(v):
    s = clean_text(v)
    if not s:
        return None
    s = s.replace("年", "-").replace("月", "-").replace("日", "").strip()
    for cand in (s, s.split(" ")[0], s.split("T")[0]):
        for f in DATE_FMTS:
            try:
                return datetime.strptime(cand, f).date()
            except ValueError:
                continue
    return None


def read_rows(path):
    p = Path(path)
    if not p.exists():
        raise SystemExit(f"文件不存在: {p}")
    suf = p.suffix.lower()
    if suf in (".xlsx", ".xlsm"):
        try:
            import openpyxl
        except ImportError:
            raise SystemExit("读 xlsx 需要 openpyxl：pip install openpyxl（或先另存为 csv）")
        wb = openpyxl.load_workbook(filename=str(p), data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        raw = list(ws.iter_rows(values_only=True))
        wb.close()
        if not raw:
            return [], []
        header = [clean_text(h) for h in raw[0]]
        data = [[clean_text(c) for c in r] for r in raw[1:]]
        return header, [r for r in data if any(x for x in r)]
    if suf == ".xls":
        raise SystemExit(f"旧版 .xls 请先另存为 .xlsx 或 .csv: {p}")
    for enc in ("utf-8-sig", "gb18030", "utf-8"):
        try:
            with open(p, "r", encoding=enc, newline="") as f:
                rows = [r for r in csv.reader(f)]
        except UnicodeDecodeError:
            continue
        if not rows:
            return [], []
        header = [clean_text(h) for h in rows[0]]
        data = [r for r in rows[1:] if any(clean_text(c) for c in r)]
        return header, data
    raise SystemExit(f"无法识别文件编码: {p}")


def need_col(header, name, where, optional=False):
    if name is None and optional:
        return None
    for i, h in enumerate(header):
        if h == name:
            return i
    for i, h in enumerate(header):
        if clean_text(h).lower() == clean_text(name).lower():
            return i
    if optional:
        return None
    raise SystemExit(f"{where}找不到列「{name}」。实际列: {', '.join(header)}")


def write_csv(path, header, rows):
    path = Path(path)
    tmp = path.with_suffix(path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        for r in rows:
            w.writerow(r)
    os.replace(tmp, path)


def write_text(path, text):
    path = Path(path)
    tmp = path.with_suffix(path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(text)
    os.replace(tmp, path)


def fmt_money(x):
    if x is None:
        return ""
    return "{:,.2f}".format(x)


def write_workbook(path, sheets):
    """把若干 (sheet名, header, rows) 写成一个带样式的 xlsx；没有 openpyxl 就跳过。"""
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("!! 未安装 openpyxl，跳过 xlsx 输出")
        return False
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, header, rows in sheets:
        ws = wb.create_sheet(title=str(name)[:31])
        ws.append(list(header))
        for r in rows:
            ws.append(list(r))
        for c in range(1, len(header) + 1):
            cell = ws.cell(row=1, column=c)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="4F81BD")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[get_column_letter(c)].width = 20
        ws.freeze_panes = "A2"
        if rows:
            ws.auto_filter.ref = ws.dimensions
    path = Path(path)
    tmp = path.with_suffix(path.suffix + ".tmp")
    wb.save(str(tmp))
    os.replace(tmp, path)
    return True


def run_merge(a):
    mapping = {}
    for m in a.map or []:
        if "=" not in m:
            raise SystemExit(f"--map 需要 旧列名=新列名 的形式: {m}")
        k, v = m.split("=", 1)
        mapping[clean_text(k)] = clean_text(v)
    header_out, records = [], []
    per_file = []
    for path in a.inputs:
        h, rows = read_rows(path)
        if not h:
            continue
        mapped = [mapping.get(clean_text(x), clean_text(x)) for x in h]
        for name in mapped:
            if name and name not in header_out:
                header_out.append(name)
        kept = 0
        for r in rows:
            rec = {"_source": Path(path).name}
            for i, name in enumerate(mapped):
                if not name:
                    continue
                rec[name] = clean_text(r[i]) if i < len(r) else ""
            if any(v for k, v in rec.items() if k != "_source"):
                records.append(rec)
                kept += 1
        per_file.append((Path(path).name, len(rows), kept))
    if not records:
        raise SystemExit("没有读到任何数据行")

    issues, seen, deduped = [], {}, []
    key = a.dedupe_by
    norm_cols = [clean_text(c) for c in (a.amount_cols or [])]
    for rec in records:
        if key:
            k = norm_key(rec.get(clean_text(key), ""))
            if not k:
                issues.append(("缺少主键", rec))
                continue
            if k in seen:
                issues.append(("重复主键 " + k, rec))
                continue
            seen[k] = rec
        if norm_cols:
            for c in norm_cols:
                if c in rec:
                    amt = to_amount(rec[c])
                    if amt is None:
                        issues.append(("金额无法解析 " + c + "=" + str(rec[c]), rec))
                    else:
                        rec[c] = fmt_money(amt)
                    break
        deduped.append(rec)

    out = Path(a.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    cols = ["_source"] + header_out
    write_csv(out / "merged.csv", cols, [[rec.get(c, "") for c in cols] for rec in deduped])
    write_csv(out / "issues.csv", ["问题", "来源"] + header_out,
              [[why, rec.get("_source", "")] + [rec.get(c, "") for c in header_out] for why, rec in issues])

    lines = ["# 多表合并结果", ""]
    lines.append("- 输入文件: " + str(len(a.inputs)) + " 个")
    for name, total, kept in per_file:
        lines.append("  - " + name + ": 读入 " + str(total) + " 行，保留 " + str(kept) + " 行")
    lines.append("- 合并后列: " + str(len(header_out)) + " 列 -> " + ", ".join(header_out))
    lines.append("- 输出记录: " + str(len(deduped)) + " 行；问题记录: " + str(len(issues)) + " 条")
    if key:
        lines.append("- 去重主键: " + key + "（保留首次出现）")
    write_text(out / "summary.md", "\n".join(lines) + "\n")
    if getattr(a, "xlsx", False):
        write_workbook(out / "合并结果.xlsx", [
            ("合并总表", cols, [[rec.get(c, "") for c in cols] for rec in deduped]),
            ("问题清单", ["问题", "来源"] + header_out,
             [[why, rec.get("_source", "")] + [rec.get(c, "") for c in header_out] for why, rec in issues]),
            ("说明", ["项目", "内容"], [[l.lstrip("# ").strip(), ""] for l in lines if l.strip()]),
        ])
    print("OK merge:", len(deduped), "行 ->", out / "merged.csv", "（问题", len(issues), "条 -> issues.csv）")


def aggregate_side(header, rows, ki, ai):
    bucket, dup = {}, {}
    for r in rows:
        k = norm_key(r[ki]) if ki < len(r) else ""
        amt = to_amount(r[ai]) if (ai is not None and ai < len(r)) else None
        if not k:
            continue
        if k in bucket:
            dup[k] = dup.get(k, 1) + 1
        b = bucket.setdefault(k, {"count": 0, "amount": 0.0, "bad_amount": 0, "row": r})
        b["count"] += 1
        if amt is None:
            b["bad_amount"] += 1
        else:
            b["amount"] += amt
    return bucket, dup


def run_recon(a):
    lh, lr = read_rows(a.left)
    rh, rr = read_rows(a.right)
    lkey = a.left_key or a.key
    rkey = a.right_key or a.key
    if not lkey or not rkey:
        raise SystemExit("对账需要 --left-key 和 --right-key（或统一的 --key）")
    li = need_col(lh, lkey, "左表")
    ri = need_col(rh, rkey, "右表")
    la = need_col(lh, a.left_amount, "左表", optional=True) if a.left_amount else None
    ra = need_col(rh, a.right_amount or a.left_amount, "右表", optional=True) if (a.right_amount or a.left_amount) else None

    left, ldup = aggregate_side(lh, lr, li, la)
    right, rdup = aggregate_side(rh, rr, ri, ra)
    tol = a.tolerance

    matched, only_left, only_right, diffs = [], [], [], []
    for k, b in left.items():
        if k in right:
            c = right[k]
            diff = b["amount"] - c["amount"]
            matched.append((k, b["count"], c["count"], b["amount"], c["amount"], diff))
            if abs(diff) > tol:
                diffs.append((k, b["amount"], c["amount"], diff))
        else:
            only_left.append(k)
    for k in right:
        if k not in left:
            only_right.append(k)

    out = Path(a.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    write_csv(out / "matched.csv", ["主键", "左表条数", "右表条数", "左表金额", "右表金额", "差额"],
              [[k, lc, rc, fmt_money(la_), fmt_money(ra_), fmt_money(d)] for k, lc, rc, la_, ra_, d in matched])
    write_csv(out / "amount_diff.csv", ["主键", "左表金额", "右表金额", "差额"],
              [[k, fmt_money(x), fmt_money(y), fmt_money(d)] for k, x, y, d in diffs])
    write_csv(out / "only_left.csv", ["主键"] + lh, [[k] + left[k]["row"] for k in only_left])
    write_csv(out / "only_right.csv", ["主键"] + rh, [[k] + right[k]["row"] for k in only_right])
    write_csv(out / "dup_keys.csv", ["表", "主键", "出现次数"],
              [["左表", k, n] for k, n in sorted(ldup.items())] + [["右表", k, n] for k, n in sorted(rdup.items())])

    lsum = sum(b["amount"] for b in left.values())
    rsum = sum(b["amount"] for b in right.values())
    lines = ["# 双表对账结果", "",
             "- 左表: " + Path(a.left).name + " 共 " + str(len(lr)) + " 行，" + str(len(left)) + " 个唯一主键，金额合计 " + fmt_money(lsum),
             "- 右表: " + Path(a.right).name + " 共 " + str(len(rr)) + " 行，" + str(len(right)) + " 个唯一主键，金额合计 " + fmt_money(rsum),
             "- 匹配: " + str(len(matched)) + " 个；仅左表有: " + str(len(only_left)) + "；仅右表有: " + str(len(only_right)),
             "- 金额不一致（容差 " + str(tol) + "）: " + str(len(diffs)) + "；主键重复: " + str(len(ldup) + len(rdup)) + " 个",
             "",
             "## 交付文件", "",
             "- matched.csv 匹配上的明细与差额",
             "- amount_diff.csv 金额对不上的（重点看这个）",
             "- only_left.csv / only_right.csv 单边有、对方没有的",
             "- dup_keys.csv 主键重复的（合并后金额可能翻倍，需与客户确认口径）"]
    if diffs[:5]:
        lines += ["", "## 差额前 5 条", ""]
        lines += ["- " + k + ": 左 " + fmt_money(x) + " / 右 " + fmt_money(y) + " / 差 " + fmt_money(d) for k, x, y, d in diffs[:5]]
    write_text(out / "summary.md", "\n".join(lines) + "\n")
    if getattr(a, "xlsx", False):
        write_workbook(out / "对账结果.xlsx", [
            ("匹配明细", ["主键", "左表条数", "右表条数", "左表金额", "右表金额", "差额"],
             [[k, lc, rc, la_, ra_, round(d, 2)] for k, lc, rc, la_, ra_, d in matched]),
            ("金额差异", ["主键", "左表金额", "右表金额", "差额"],
             [[k, la_, ra_, round(d, 2)] for k, la_, ra_, d in diffs]),
            ("仅左表有", ["主键"] + lh, [[k] + left[k]["row"] for k in only_left]),
            ("仅右表有", ["主键"] + rh, [[k] + right[k]["row"] for k in only_right]),
            ("重复主键", ["表", "主键", "出现次数"],
             [["左表", k, n] for k, n in sorted(ldup.items())] + [["右表", k, n] for k, n in sorted(rdup.items())]),
            ("说明", ["项目", "内容"], [[l.lstrip("# ").strip(), ""] for l in lines if l.strip()]),
        ])
    print("OK recon: 匹配", len(matched), "仅左", len(only_left), "仅右", len(only_right), "差异", len(diffs), "->", out)


def run_report(a):
    header, rows = read_rows(a.input)
    gi = need_col(header, a.group, "表")
    vi = need_col(header, a.value, "表", optional=True) if a.value else None
    di = need_col(header, a.date, "表", optional=True) if a.date else None
    buckets = {}
    for r in rows:
        g = clean_text(r[gi]) if gi < len(r) else ""
        if a.month and di is not None:
            d = to_date(r[di]) if di < len(r) else None
            g = d.strftime("%Y-%m") if d else g
        b = buckets.setdefault(g, {"n": 0, "sum": 0.0})
        b["n"] += 1
        if vi is not None:
            amt = to_amount(r[vi]) if vi < len(r) else None
            if amt is not None:
                b["sum"] += amt
    ordered = sorted(buckets.items(), key=lambda kv: (-kv[1]["sum"], kv[0]))
    out = Path(a.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    hdr = ["分组", "条数"] + (["金额合计"] if vi is not None else [])
    body = [[k, v["n"]] + ([fmt_money(v["sum"])] if vi is not None else []) for k, v in ordered]
    if vi is not None:
        body.append(["合计", sum(v["n"] for v in buckets.values()), fmt_money(sum(v["sum"] for v in buckets.values()))])
    write_csv(out / "report.csv", hdr, body)
    lines = ["# 汇总报表", "", "- 来源: " + Path(a.input).name + "，" + str(len(rows)) + " 行",
             "- 分组列: " + a.group + (("，金额列: " + a.value) if a.value else "") + (("，按月(" + str(a.date) + ")") if a.month else ""), ""]
    write_text(out / "summary.md", "\n".join(lines) + "\n")
    print("OK report:", len(ordered), "个分组 ->", out / "report.csv")


def main():
    ap = argparse.ArgumentParser(description="sheetkit 表格清洗/合并/对账/报表")
    sub = ap.add_subparsers(dest="cmd", required=True)

    m = sub.add_parser("merge", help="多表合并 + 去重 + 列名对齐")
    m.add_argument("--inputs", nargs="+", required=True)
    m.add_argument("--out-dir", default="out")
    m.add_argument("--map", action="append", default=[], help="旧列名=新列名，可重复")
    m.add_argument("--dedupe-by", default=None, help="去重主键列名")
    m.add_argument("--amount-cols", nargs="*", default=[], help="需要规范为 2 位小数的金额列")
    m.add_argument("--xlsx", action="store_true", help="额外输出带样式的 xlsx")
    m.set_defaults(func=run_merge)

    r = sub.add_parser("recon", help="双表对账")
    r.add_argument("--left", required=True)
    r.add_argument("--right", required=True)
    r.add_argument("--out-dir", default="out")
    r.add_argument("--key", default=None)
    r.add_argument("--left-key", default=None)
    r.add_argument("--right-key", default=None)
    r.add_argument("--left-amount", default=None)
    r.add_argument("--right-amount", default=None)
    r.add_argument("--tolerance", type=float, default=0.01)
    r.add_argument("--xlsx", action="store_true", help="额外输出带样式的 xlsx")
    r.set_defaults(func=run_recon)

    p = sub.add_parser("report", help="分组汇总")
    p.add_argument("--input", required=True)
    p.add_argument("--group", required=True)
    p.add_argument("--value", default=None)
    p.add_argument("--date", default=None, help="配合 --month 使用日期列")
    p.add_argument("--month", action="store_true", help="按年月分组")
    p.add_argument("--out-dir", default="out")
    p.add_argument("--xlsx", action="store_true", help="额外输出带样式的 xlsx")
    p.set_defaults(func=run_report)

    a = ap.parse_args()
    a.func(a)


if __name__ == "__main__":
    main()
