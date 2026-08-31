# -*- coding: utf-8 -*-
"""当前应核季度自动推算 + 日历公司名↔真库行名匹配（2026-08-31 建）。
build_site.py 与 _caibao_watch.py 共用，消除「财报季推进时手工调 CUR_Q」。

规则（披露季后各公司陆续出数，取主波开启日）：
  Q1 季 04-20 起核、Q2 季 08-10 起核、Q3 季 10-25 起核、Q4 季 次年 02-20 起核。
CUR_Q 形如 "26Q2"；PREV_Q = 去年同季（守望者对照列）。
"""
from datetime import date

# (开始核对的 (月,日), 季度序号)
_OPEN = [((2, 20), 4), ((4, 20), 1), ((8, 10), 2), ((10, 25), 3)]


def cur_check_quarter(today: date | None = None) -> tuple[str, str]:
    t = today or date.today()
    yy = t.year % 100
    cur_year, cur_q = yy - 1, 4  # 年初默认核上年 Q4
    for (m, d), q in _OPEN:
        if (t.month, t.day) >= (m, d):
            if q == 4:
                cur_year, cur_q = yy - 1, 4
            else:
                cur_year, cur_q = yy, q
    cur = f"{cur_year:02d}Q{cur_q}"
    prev = f"{cur_year - 1:02d}Q{cur_q}"
    return cur, prev


def cur_check_quarter_iso(today: date | None = None) -> tuple[str, str]:
    """网站 data 键格式版：("2026Q2", "2025Q2")。"""
    cur, prev = cur_check_quarter(today)
    return "20" + cur, "20" + prev


# 日历公司名 → 真库行名（别名表；两边写法不一致时在此登记）
ALIASES = {
    "Teck": "TECK", "Boliden": "BOLIDEN", "Nexa": "NEXA",
    "明苏尔": "明苏尔", "Minsur": "明苏尔", "Vedanta": "Vedanta",
    "大全能源": "新疆大全新能源",
}


def norm(name: str) -> str:
    return "".join(ch for ch in str(name).lower() if ch.isalnum() or ch > "ÿ")


def company_match(cal_name: str, row_names: list[str]) -> str | None:
    """日历公司名匹配真库行名：别名 → 精确 → 互相包含（取最短包含串防误配）。"""
    target = ALIASES.get(cal_name, cal_name)
    n0 = norm(target)
    for rn in row_names:
        if norm(rn) == n0:
            return rn
    cands = [rn for rn in row_names if n0 and (n0 in norm(rn) or norm(rn) in n0)]
    if len(cands) == 1:
        return cands[0]
    if cands:
        cands.sort(key=len)
        return cands[0] if norm(cands[0]) == norm(cands[-1]) else None
    return None
