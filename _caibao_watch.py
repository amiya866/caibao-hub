# -*- coding: utf-8 -*-
"""财报平台欠账守望者（2026-08-23 建，替代已消亡的 Kimi 会话 cron）
每周一 09:11（Windows 任务 CaibaoWatch）：
  ① 扫 8 品种真库：有 25Q2 但 26Q2 空的公司行 = 疑似「已披露未入库」；
  ② 写待核清单 D:\Kimi\caibao-hub\_待核清单.md；
  ③ 有欠账就弹窗提醒（叫 Kimi 来补）。
另：提醒信息速递每周扫描（印尼镍/铝讯/阿拉丁/爱择/抖音传言）。
手动跑: python _caibao_watch.py
"""
import sys
from datetime import date
from pathlib import Path

import openpyxl

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

FILES = {
    "铜": r"D:\拷贝文件\E\永安\铜\全球铜企季度产量梳理.xlsx",
    "铝": r"D:\拷贝文件\E\永安\铝\全球铝企季度产量梳理.xlsx",
    "铅": r"D:\拷贝文件\E\永安\铅\全球铅企季度产量梳理.xlsx",
    "锌": r"D:\拷贝文件\E\永安\锌\全球锌企季度产量梳理.xlsx",
    "镍": r"D:\拷贝文件\E\永安\镍\全球镍企季度产量梳理.xlsx",
    "锂": r"D:\拷贝文件\E\永安\锂\全球锂企季度产量梳理.xlsx",
    "硅": r"D:\拷贝文件\E\永安\硅产业\全球硅企季度产量梳理.xlsx",
    "锡": r"D:\拷贝文件\E\永安\锡\海外主要公司产量.xlsx",
}
OUT = Path(r"D:\Kimi\caibao-hub\_待核清单.md")

# 当前最新应有季度（财报季推进时手工调：Q2 季=26Q2，Q3 季=26Q3）
CUR_Q, PREV_Q = "26Q2", "25Q2"


def scan(path):
    """返回 [(sheet, 公司, 国家/项目)]：上一季有数、当季空。"""
    owes = []
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        return [("(文件读取失败)", str(e)[:60], "", "")]
    for ws in wb:
        if "日志" in ws.title:
            continue
        hdr = None
        for r in ws.iter_rows(min_row=1, max_row=5, values_only=True):
            if r and any(str(v).strip() == CUR_Q for v in r if v):
                hdr = r
                break
        if not hdr:
            continue
        cols = {str(v).strip(): j for j, v in enumerate(hdr) if v}
        if CUR_Q not in cols or PREV_Q not in cols:
            continue
        for r in ws.iter_rows(min_row=6, values_only=True):
            name = r[0]
            if not name or "总计" in str(name) or "合计" in str(name):
                continue
            cur, prev = r[cols[CUR_Q]], r[cols[PREV_Q]]
            if prev not in (None, "", "-") and cur in (None, "", "-"):
                owes.append((ws.title, str(name), str(r[1] or ""), str(r[2] or "") if len(r) > 2 else ""))
    wb.close()
    return owes


def popup(msg):
    import subprocess
    try:
        esc = msg.replace("'", "''").replace("\n", "`n")
        subprocess.Popen(["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command",
                          "Add-Type -AssemblyName System.Windows.Forms;"
                          f"[System.Windows.Forms.MessageBox]::Show('{esc}','财报平台欠账提醒','OK','Warning')"])
    except Exception:
        pass


def main():
    lines = [f"# 财报平台待核清单（{date.today()} 自动生成，守望者 _caibao_watch.py）",
             f"\n规则：{PREV_Q} 有数但 {CUR_Q} 空 = 疑似已披露未入库（财报季推进时改脚本 CUR_Q）。\n"]
    total = 0
    for comm, path in FILES.items():
        owes = scan(path)
        total += len(owes)
        lines.append(f"\n## {comm}（{len(owes)} 行）")
        for sheet, name, country, proj in owes:
            lines.append(f"- [{sheet}] {name} {country} {proj}".rstrip())
    OUT.write_text("\n".join(lines), encoding="utf-8")
    print(f"欠账 {total} 行 → {OUT}")
    if total:
        popup(f"财报平台疑似欠账 {total} 行（{CUR_Q} 未入库）。\n清单：{OUT}\n叫 Kimi：「按待核清单更新财报平台」。\n\n另：信息速递每周扫描（印尼镍/铝讯/阿拉丁/爱择/抖音传言）也别忘。")


if __name__ == "__main__":
    main()
