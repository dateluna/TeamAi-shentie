#!/usr/bin/env python3
"""
量化T0营销 · 持仓交集与收益对比分析脚本

解析《量化标的》Excel 与客户持仓文件（Excel / CSV），
找出客户持仓中与《量化标的》一致的股票，并输出结构化的对比数据。

用法:
    analyze_t0.py --targets <量化标的.xlsx> --holding <客户持仓.xlsx|csv> [--market-data <行情.csv>] [--output <result.json>]

参数:
    --targets       《量化标的》文件路径（必填）。列：开始时间、结束时间、股票代码、股票名称、
                    股票评级、日均授权市值(元）、实现费后盈亏额（元）、区间盈利率
    --holding       客户持仓文件路径（必填）。支持 xlsx / xls(OOXML) / csv，
                    自动识别列：证券代码/股票代码/代码、证券名称/股票名称/名称、市值等
    --market-data   (可选) 行情数据 CSV，表头: 日期,股票代码,收盘价（用于计算区间实际涨跌幅）。
                    若不提供，market_change_pct 置 null，由调用方后续通过网络查询补充
    --output        输出 JSON 路径（可选，默认打印到 stdout）

输出:
    JSON: { targets_period, holdings_matched, matched: [...], unmatched_holdings: [...] }
    每个 matched 项含:
        code / name / rating / authorized_value / t0_profit / t0_interval_profit_pct(区间盈利率)
        market_change_pct(区间市场实际涨跌幅, 可能为 null)
"""

import argparse
import csv
import json
import re
import sys
from pathlib import Path


def sniff_format(path: Path):
    """按文件头嗅探真实格式（不信任扩展名，兼容 .xls 实为 OOXML 的情况）。"""
    with open(path, "rb") as f:
        magic = f.read(8)
    if magic.startswith(b"PK\x03\x04"):
        return "xlsx"
    if magic.startswith(b"\xD0\xCF\x11\xE0"):
        return "xls_ole"
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return "csv"
    raise ValueError(f"无法识别文件格式: {path}（支持 .xlsx / .xls / .csv）")


def read_table(path: Path):
    """读取 xlsx / xls / csv，返回 (行列表[dict], 表头[list])。"""
    import tempfile
    fmt = sniff_format(path)
    if fmt in ("xlsx", "xls_ole"):
        import openpyxl
        if fmt == "xls_ole":
            # 老式二进制 xls 先经 xlrd 转 OOXML 再读
            import io
            import xlrd
            xb = xlrd.open_workbook(str(path))
            ws0 = xb.sheet_by_index(0)
            xlsx_io = io.BytesIO()
            wbx = openpyxl.Workbook()
            wxs = wbx.active
            for r in range(ws0.nrows):
                wxs.append([ws0.cell_value(r, c) for c in range(ws0.ncols)])
            wbx.save(xlsx_io)
            xlsx_io.seek(0)
            wb = openpyxl.load_workbook(xlsx_io, data_only=True, read_only=True)
        else:
            # openpyxl 按扩展名校验：OOXML 但扩展名非 .xlsx 时复制为临时 .xlsx
            if path.suffix.lower() != ".xlsx":
                tmp = Path(tempfile.mkstemp(suffix=".xlsx")[1])
                tmp.write_bytes(path.read_bytes())
                wb = openpyxl.load_workbook(tmp, data_only=True, read_only=True)
                tmp.unlink(missing_ok=True)
            else:
                wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb.worksheets[0]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
        rows = [r for r in rows if any(c is not None for c in r)]
        if not rows:
            raise ValueError(f"文件为空: {path}")
        header = [str(c).strip() if c is not None else "" for c in rows[0]]
        return [dict(zip(header, r)) for r in rows[1:]], header
    elif fmt == "csv":
        with open(path, encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            rows = [r for r in reader if any(c.strip() for c in r)]
        if not rows:
            raise ValueError(f"文件为空: {path}")
        header = [c.strip() for c in rows[0]]
        return [dict(zip(header, r)) for r in rows[1:]], header
    raise ValueError(f"不支持的文件格式: {path}")


def norm_code(raw):
    """股票代码规范化：去掉 .SH/.SZ 后缀，数字补零到 6 位。"""
    if raw is None:
        return None
    s = str(raw).strip().upper()
    s = re.sub(r"\.(SH|SZ|BJ)$", "", s)
    if s.endswith((".SH", ".SZ", ".BJ")):
        s = s[:-3]
    s = s.replace(" ", "")
    if re.fullmatch(r"\d+", s):
        return s.zfill(6)
    if re.fullmatch(r"\d{6}", s):
        return s
    return s or None


def find_col(header, candidates):
    """在表头中按候选名模糊匹配列（支持包含匹配，如 '证券代码' 匹配 '证券代码'）。"""
    for cand in candidates:
        for h in header:
            if cand.lower() == h.lower() or cand.lower() in h.lower() or h.lower() in cand.lower():
                return h
    return None


def parse_targets(path: Path):
    rows, header = read_table(path)
    code_col = find_col(header, ["股票代码", "证券代码", "代码"])
    name_col = find_col(header, ["股票名称", "证券名称", "名称"])
    rating_col = find_col(header, ["股票评级"])
    auth_col = find_col(header, ["日均授权市值"])
    profit_col = find_col(header, ["实现费后盈亏额"])
    pct_col = find_col(header, ["区间盈利率"])
    start_col = find_col(header, ["开始时间"])
    end_col = find_col(header, ["结束时间"])

    if not (code_col and pct_col):
        raise ValueError(
            f"《量化标的》缺少必要列（需要 股票代码 + 区间盈利率）。实际表头: {header}"
        )

    targets = []
    for r in rows:
        code = norm_code(r.get(code_col))
        if not code:
            continue
        def num(key):
            v = r.get(key)
            if v is None or v == "":
                return None
            try:
                return float(v)
            except (TypeError, ValueError):
                return None
        targets.append({
            "code": code,
            "name": str(r.get(name_col)).strip() if name_col and r.get(name_col) is not None else None,
            "rating": r.get(rating_col) if rating_col else None,
            "authorized_value": num(auth_col) if auth_col else None,
            "t0_profit": num(profit_col) if profit_col else None,
            "t0_interval_profit_pct": num(pct_col),
            "period_start": str(r.get(start_col)).strip() if start_col else None,
            "period_end": str(r.get(end_col)).strip() if end_col else None,
        })
    return targets, (str(r.get(start_col)).strip() if start_col else None,
                     str(r.get(end_col)).strip() if end_col else None)


def parse_holdings(path: Path):
    rows, header = read_table(path)
    code_col = find_col(header, ["证券代码", "股票代码", "代码", "基金代码"])
    name_col = find_col(header, ["证券名称", "股票名称", "名称", "基金名称"])
    value_col = find_col(header, ["持仓市值", "市值", "最新市值", "市价", "持仓金额"])
    qty_col = find_col(header, ["持仓数量", "数量", "股份余额", "持有份额"])

    if not (code_col or name_col):
        raise ValueError(f"客户持仓无法识别代码/名称列。实际表头: {header}")

    holdings = []
    for r in rows:
        code = norm_code(r.get(code_col)) if code_col else None
        name = None
        if name_col and r.get(name_col) is not None:
            name = str(r.get(name_col)).strip()
        # 无代码列时，尝试从名称列兜底（不匹配则跳过）
        if not code:
            continue
        def num(key):
            v = r.get(key)
            if v is None or v == "":
                return None
            try:
                return float(v)
            except (TypeError, ValueError):
                return None
        holdings.append({
            "code": code,
            "name": name,
            "value": num(value_col) if value_col else None,
            "qty": num(qty_col) if qty_col else None,
            "raw": {k: v for k, v in r.items() if v is not None},
        })
    return holdings, header


def parse_market_data(path: Path):
    """行情 CSV：日期,股票代码,收盘价。返回 {(code): [(date, close)]}，日期升序。"""
    result = {}
    with open(path, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            date = (row.get("日期") or row.get("date") or "").strip()
            code = norm_code(row.get("股票代码") or row.get("代码") or row.get("symbol"))
            close_raw = row.get("收盘价") or row.get("close")
            if not date or not code or not close_raw:
                continue
            try:
                close = float(close_raw)
            except (TypeError, ValueError):
                continue
            result.setdefault(code, []).append((date, close))
    for code in result:
        result[code].sort(key=lambda x: x[0])
    return result


def compute_market_change(code, period_start, period_end, market):
    """计算区间 [period_start, period_end] 的市场实际涨跌幅（%）。"""
    if not market or code not in market:
        return None, None
    series = market[code]
    # 找区间内第一个(开始日或之后最近)与最后一个(结束日或之前最近)
    if period_start:
        ps = period_start[:8]
        start = next((s for s in series if s[0] >= ps), None)
        if not start:
            return None, None
    else:
        start = series[0]
    if period_end:
        pe = period_end[:8]
        end = next((s for s in reversed(series) if s[0] <= pe), None)
        if not end:
            return None, None
    else:
        end = series[-1]
    if start[1] == 0:
        return None, None
    return (end[1] - start[1]) / start[1] * 100.0, {"start": start, "end": end}


def fmt_pct(v):
    return f"{v:.2f}%" if v is not None else "N/A"


def main():
    ap = argparse.ArgumentParser(description="量化T0营销 · 持仓交集与收益对比分析")
    ap.add_argument("--targets", required=True, help="《量化标的》xlsx 路径")
    ap.add_argument("--holding", required=True, help="客户持仓 xlsx/xls/csv 路径")
    ap.add_argument("--market-data", help="可选行情 CSV（日期,股票代码,收盘价）")
    ap.add_argument("--output", help="输出 JSON 路径（默认打印 stdout）")
    args = ap.parse_args()

    t_path = Path(args.targets)
    h_path = Path(args.holding)
    if not t_path.exists():
        sys.exit(f"❌ 找不到《量化标的》文件: {t_path}")
    if not h_path.exists():
        sys.exit(f"❌ 找不到客户持仓文件: {h_path}")

    targets, period = parse_targets(t_path)
    holdings, h_header = parse_holdings(h_path)
    market = parse_market_data(Path(args.market_data)) if args.market_data else {}

    t_by_code = {t["code"]: t for t in targets}
    h_by_code = {}
    for h in holdings:
        h_by_code.setdefault(h["code"], []).append(h)

    matched = []
    for code, t in sorted(t_by_code.items(), key=lambda x: x[1]["t0_interval_profit_pct"] or 0, reverse=True):
        if code in h_by_code:
            mkt_chg, mkt_info = compute_market_change(code, t["period_start"], t["period_end"], market)
            matched.append({
                "code": code,
                "name": t["name"],
                "rating": t["rating"],
                "authorized_value": t["authorized_value"],
                "t0_profit": t["t0_profit"],
                "t0_interval_profit_pct": t["t0_interval_profit_pct"],
                "market_change_pct": round(mkt_chg, 4) if mkt_chg is not None else None,
                "market_data_points": mkt_info,
                "holding": h_by_code[code],
                "period_start": t["period_start"],
                "period_end": t["period_end"],
            })

    matched_codes = {m["code"] for m in matched}
    unmatched_holdings = [h for h in holdings if h["code"] not in matched_codes]

    result = {
        "targets_count": len(targets),
        "holdings_count": len(holdings),
        "matched_count": len(matched),
        "period_start": period[0],
        "period_end": period[1],
        "matched": matched,
        "unmatched_holdings": unmatched_holdings,
        "note": (
            "market_change_pct 为《量化标的》开始时间~结束时间区间内该标的的市场实际涨跌幅；"
            "t0_interval_profit_pct 为《量化标的》给出的区间盈利率（量化T0策略收益）。"
            "若 market_change_pct 为 null，需用网络行情补充后重新计算或标注'数据待补'。"
        ),
    }

    # 控制台摘要
    print("=" * 60)
    print(f"《量化标的》标的数: {len(targets)} | 客户持仓数: {len(holdings)} | 交集标的数: {len(matched)}")
    if period[0]:
        print(f"回测区间: {period[0]} ~ {period[1]}")
    print("=" * 60)
    for m in matched:
        flag = "✅" if (m["market_change_pct"] is not None) else "⚠️ 市场涨跌幅待补"
        print(f"{m['code']} {m['name'] or ''} | T0区间盈利率: {fmt_pct(m['t0_interval_profit_pct'])} | "
              f"市场实际涨跌幅: {fmt_pct(m['market_change_pct'])} {flag}")
    if unmatched_holdings:
        print("-" * 60)
        print(f"持仓中未匹配《量化标的》的标的 ({len(unmatched_holdings)} 只):")
        for h in unmatched_holdings:
            print(f"  {h['code']} {h['name'] or ''}")

    out = args.output
    if out:
        Path(out).parent.mkdir(parents=True, exist_ok=True)
        Path(out).write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        print(f"\n📄 结果已写入: {out}")
    else:
        print(json.dumps(result, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    main()
