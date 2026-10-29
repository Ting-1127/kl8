#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""快乐8开奖数据自动更新：多数据源抓取 + 增量合并到 JSON/Excel。

数据源（按顺序自动降级）：
  1. 中国福利彩票官网接口  cwl.gov.cn  （官方 JSON，可批量拉取）
  2. 500彩票网开奖页      kaijiang.500.com （HTML 解析，仅最新一期）

用法：
  python3 fetch_kl8.py                  # 自动选源，增量更新
  python3 fetch_kl8.py --check          # 只检查不写文件（有新数据退出码 0，否则 1）
  python3 fetch_kl8.py --source 500     # 强制指定数据源（cwl | 500 | auto）

约定：所有路径相对当前工作目录。
"""
import json
import os
import re
import sys
import time
import urllib.request
from collections import Counter
from datetime import datetime

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
JSON_FILE = "快乐8开奖数据.json"
XLSX_FILE = "快乐8开奖数据.xlsx"

CWL_API = ("https://www.cwl.gov.cn/cwl_admin/front/cwlkj/search/kjxx/findDrawNotice"
           "?name=kl8&issueCount={n}&issueStart=&issueEnd=&dayStart=&dayEnd=&pageNo=1&pageSize={n}")
CWL_REFERER = "https://www.cwl.gov.cn/ygkj/wjkj/kjxx/kl8/"
PAGE_500 = "https://kaijiang.500.com/kl8.shtml"

WD = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]


class FetchError(Exception):
    pass


def _decode(raw):
    """响应体解码：兼容 UTF-8 / GBK 页面。"""
    for enc in ("utf-8", "gb18030", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def http_get(url, referer=None, retries=3, timeout=30):
    """纯 Python 标准库实现的 GET（无需 curl / subprocess，iOS 沙箱可跑）。
    返回 (body, note)，网络问题会在 note 里留痕并重试。"""
    headers = {
        "User-Agent": UA,
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        "Accept": "text/html,application/json;q=0.9,*/*;q=0.8",
        "Connection": "close",
    }
    if referer:
        headers["Referer"] = referer
    last = ""
    for i in range(retries):
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return _decode(resp.read()), ""
        except Exception as e:  # URLError / timeout / HTTPError 等
            last = f"{type(e).__name__}: {e}"
        time.sleep(2 * (i + 1))
    raise FetchError(f"GET 失败 {url[:80]}... | {last}")


# ---------------------------------------------------------------- 数据源 1：福彩官网
def fetch_cwl(count=100):
    """官方 JSON 接口，可一次拉多期。返回标准记录列表。"""
    body, _ = http_get(CWL_API.format(n=count), referer=CWL_REFERER)
    try:
        data = json.loads(body)
    except json.JSONDecodeError:
        raise FetchError(f"官网返回非 JSON（前 120 字符）: {body[:120]!r}")
    if data.get("state") != 0:
        raise FetchError(f"官网接口报错: {data.get('message')}")
    out = []
    for r in data.get("result", []):
        nums = sorted(int(x) for x in r.get("red", "").split(",") if x)
        if len(nums) != 20:
            raise FetchError(f"期号 {r.get('code')} 号码数异常: {nums}")
        out.append({
            "issue": int(r["code"]),
            "date": r["date"].split("(")[0],
            "weekday": WD[datetime.strptime(r["date"].split("(")[0], "%Y-%m-%d").weekday()],
            "numbers": nums,
        })
    return out


# ---------------------------------------------------------------- 数据源 2：500彩票网
def fetch_500():
    """解析 500 彩票网快乐8主页，取最新一期。返回标准记录列表（1 条）。"""
    body, _ = http_get(PAGE_500)
    m_issue = re.search(r'快乐8 第\s*<font[^>]*><strong>(\d{7})</strong>', body)
    m_date = re.search(r'开奖日期：(\d{4})年(\d{1,2})月(\d{1,2})日', body)
    m_balls = re.search(r'ball_box01.*?</ul>', body, re.S)
    if not (m_issue and m_date and m_balls):
        raise FetchError("500 彩票网页面结构解析失败（页面可能已改版）")
    nums = sorted(int(x) for x in re.findall(r'ball_red">(\d{2})<', m_balls.group(0)))
    if len(nums) != 20:
        raise FetchError(f"500 页面号码数异常: {nums}")
    d = datetime(int(m_date.group(1)), int(m_date.group(2)), int(m_date.group(3)))
    return [{
        "issue": int(m_issue.group(1)),
        "date": d.strftime("%Y-%m-%d"),
        "weekday": WD[d.weekday()],
        "numbers": nums,
    }]


SOURCES = {"cwl": fetch_cwl, "500": fetch_500}


# ---------------------------------------------------------------- 本地数据读写
def load_json():
    if not os.path.exists(JSON_FILE):
        return None
    with open(JSON_FILE, encoding="utf-8") as f:
        return json.load(f)


def save_json(data):
    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def rebuild_xlsx(records):
    """由全量记录重建 Excel。"""
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
    maxc = max(cnt.values()) if total else 0
    heat = PatternFill("solid", fgColor="FFC7CE")
    for n in range(1, 81):
        c = cnt.get(n, 0)
        ws2.append([n, c, (c / total) if total else 0])
        ws2.cell(row=n + 1, column=3).number_format = "0.0%"
        for col in range(1, 4):
            cell = ws2.cell(row=n + 1, column=col)
            cell.alignment, cell.border = center, border
        if total and c == maxc:
            ws2.cell(row=n + 1, column=2).fill = heat
    ws2.freeze_panes = "A2"
    for i, w in enumerate([8, 10, 10], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    wb.save(XLSX_FILE)


# ---------------------------------------------------------------- 主流程
def main():
    args = sys.argv[1:]
    check_only = "--check" in args
    source = "auto"
    if "--source" in args:
        source = args[args.index("--source") + 1]
    if source not in ("auto", "cwl", "500"):
        print(f"未知数据源: {source}", file=sys.stderr)
        return 2

    data = load_json()
    known = {rec["issue"] for rec in data["issues"]} if data else set()

    order = ["cwl", "500"] if source == "auto" else [source]
    errors, fresh, used = [], [], None
    for name in order:
        try:
            records = SOURCES[name]()
            if not records:
                raise FetchError("返回 0 条记录")
            fresh = [r for r in records if r["issue"] not in known]
            used = name
            break
        except FetchError as e:
            errors.append(f"[{name}] {e}")

    if used is None:
        print("FAIL 所有数据源均不可用:", file=sys.stderr)
        for e in errors:
            print("  " + e, file=sys.stderr)
        return 2

    for e in errors:
        print("WARN 降级: " + e)

    if not fresh:
        print("NO_NEW")
        return 1 if check_only else 0

    if check_only:
        for rec in sorted(fresh, key=lambda x: x["issue"]):
            print(f"NEW {rec['issue']} {rec['date']} {rec['numbers']}")
        return 0

    if data is None:
        data = {"game": "快乐8", "total_issues": 0, "date_range": {},
                "issues": [], "number_frequency": {}}

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
    print(f"SOURCE {used}")
    print(f"TOTAL {data['total_issues']} issues, range {data['date_range']['start']} ~ {data['date_range']['end']}")

    # 缺口提示：本地最大期号 与 新抓最大期号 之间若断档，提醒人工补
    if used == "500":
        print("NOTE 备用源仅含最新一期；如本地存在缺期，建议网络恢复后用官方源补齐")
    return 0


if __name__ == "__main__":
    sys.exit(main())
