#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""快乐8开奖数据自动更新：从中国福利彩票官网抓取最新开奖，合并到 JSON/Excel。

用法：
  python3 fetch_kl8.py            # 增量更新（只抓缺失期号后合并）
  python3 fetch_kl8.py --check    # 只检查有没有新数据，不写文件（有新数据退出码 0，否则 1）

约定：所有路径相对当前工作目录（默认工作区根目录）。
"""
import json
import os
import re
import subprocess
import sys
import time
from collections import Counter
from datetime import datetime

API = (
    "https://www.cwl.gov.cn/cwl_admin/front/cwlkj/search/kjxx/findDrawNotice"
    "?name=kl8&issueCount={n}&issueStart=&issueEnd=&dayStart=&dayEnd=&pageNo=1&pageSize={n}"
)
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
JSON_FILE = "快乐8开奖数据.json"
XLSX_FILE = "快乐8开奖数据.xlsx"

WD_MAP = {"一": "周一", "二": "周二", "三": "周三", "四": "周四", "五": "周五", "六": "周六", "日": "周日"}


def fetch(count=30):
    """调用福彩官网接口，返回原始记录列表。"""
    url = API.format(n=count)
    out = subprocess.run(
        ["curl", "-sL", "-m", "30", url,
         "-H", f"User-Agent: {UA}",
         "-H", "Referer: https://www.cwl.gov.cn/ygkj/wjkj/kjxx/kl8/",
         "-H", "Accept: application/json"],
        capture_output=True, text=True,
    ).stdout
    data = json.loads(out)
    if data.get("state") != 0:
        raise RuntimeError(f"接口返回异常: {data.get('message')}")
    return data["result"]


def parse(raw_list):
    """原始记录 -> 标准记录。"""
    records = []
    for r in raw_list:
        nums = sorted(int(x) for x in r["red"].split(",") if x)
        if len(nums) != 20:
            raise ValueError(f"期号 {r['code']} 号码数量异常: {nums}")
        date_s = r["date"].split("(")[0]
        wd = WD_MAP.get(r.get("week", ""), "")
        records.append({
            "issue": int(r["code"]),
            "date": date_s,
            "weekday": wd,
            "numbers": nums,
        })
    return records


def load_json():
    if not os.path.exists(JSON_FILE):
        return None
    with open(JSON_FILE, encoding="utf-8") as f:
        return json.load(f)


def save_json(data):
    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def rebuild_xlsx(records):
    """由全量记录重建 Excel（结构与首次生成一致）。"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    rows = sorted(records, key=lambda x: x["issue"], reverse=True)
    thin = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    center = Alignment(horizontal="center", vertical="center")

    wb = Workbook()
    ws = wb.active
    ws.title = "开奖明细"
    headers = ["期号", "日期", "星期"] + [f"号码{i}" for i in range(1, 21)]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font, cell.alignment, cell.border = header_fill, header_font, center, border

    for rec in rows:
        d = datetime.strptime(rec["date"], "%Y-%m-%d").date()
        ws.append([rec["issue"], d, rec["weekday"]] + rec["numbers"])

    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=1).number_format = "0"
        ws.cell(row=r, column=2).number_format = "yyyy-mm-dd"
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            cell.alignment, cell.border = center, border

    ws.freeze_panes = "D2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"
    for i, w in enumerate([10, 12, 8] + [7] * 20, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws2 = wb.create_sheet("号码频次")
    cnt = Counter()
    for rec in rows:
        cnt.update(rec["numbers"])
    ws2.append(["号码", "出现次数", "出现频率"])
    for c in range(1, 4):
        cell = ws2.cell(row=1, column=c)
        cell.fill, cell.font, cell.alignment, cell.border = header_fill, header_font, center, border
    total = len(rows)
    maxc = max(cnt.values())
    heat = PatternFill("solid", fgColor="FFC7CE")
    for n in range(1, 81):
        c = cnt.get(n, 0)
        ws2.append([n, c, c / total])
        ws2.cell(row=n + 1, column=3).number_format = "0.0%"
        for col in range(1, 4):
            cell = ws2.cell(row=n + 1, column=col)
            cell.alignment, cell.border = center, border
        if c == maxc:
            ws2.cell(row=n + 1, column=2).fill = heat
    ws2.freeze_panes = "A2"
    for i, w in enumerate([8, 10, 10], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    wb.save(XLSX_FILE)


def main():
    check_only = "--check" in sys.argv

    data = load_json()
    known = {rec["issue"] for rec in data["issues"]} if data else set()

    # 抓最近 30 期，足够覆盖任何中断
    raw = fetch(30)
    fresh = [rec for rec in parse(raw) if rec["issue"] not in known]

    if not fresh:
        print("NO_NEW")
        return 1 if check_only else 0

    if check_only:
        for rec in sorted(fresh, key=lambda x: x["issue"]):
            print(f"NEW {rec['issue']} {rec['date']} {rec['numbers']}")
        return 0

    if data is None:
        data = {
            "game": "快乐8",
            "total_issues": 0,
            "date_range": {},
            "issues": [],
            "number_frequency": {},
        }

    data["issues"] = sorted(data["issues"] + fresh, key=lambda x: x["issue"], reverse=True)
    data["total_issues"] = len(data["issues"])
    dates = sorted(rec["date"] for rec in data["issues"])
    data["date_range"] = {"start": dates[0], "end": dates[-1]}
    cnt = Counter()
    for rec in data["issues"]:
        cnt.update(rec["numbers"])
    data["number_frequency"] = {str(n): cnt.get(n, 0) for n in range(1, 81)}

    save_json(data)
    rebuild_xlsx(data["issues"])

    for rec in sorted(fresh, key=lambda x: x["issue"]):
        print(f"ADDED {rec['issue']} {rec['date']} {','.join(map(str, rec['numbers']))}")
    print(f"TOTAL {data['total_issues']} issues, range {data['date_range']['start']} ~ {data['date_range']['end']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
