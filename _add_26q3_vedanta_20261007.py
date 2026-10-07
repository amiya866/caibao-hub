# -*- coding: utf-8 -*-
"""一次性：开 26Q3 季度列（锌/铝/铅三真库）+ Vedanta Q2FY27(2026年7-9月) 产量入库
- 每sheet 在「26Q2同比」后插入两列：26Q3 / 26Q3同比（样式拷贝自 26Q2 两列）
- 锌走 _zinc_write 网关；铝/铅 先备份再写；写后同步 caibao-hub\excel\ 镜像
- Vedanta 数值来源：Vedanta Q2FY27 Production Release（2026-10-03，多源交叉核实）
"""
import shutil, sys
from copy import copy
from datetime import datetime
from pathlib import Path

import openpyxl

sys.stdout.reconfigure(encoding='utf-8', errors="replace")

BASE = Path(r"D:\Kimi\金属总网\网站构建\财报汇总")
MIRROR = Path(r"D:\Kimi\caibao-hub\excel")

FILLS = {  # file: {sheet: {公司行名: 26Q3值}}, log: (内容, 来源)
    "锌": {
        "sheets": {
            "锌矿企业·季度产量": {"Vedanta": 27.1},          # HZL 采矿金属 271kt(+5%)
            "锌锭冶炼企业·季度产量": {"Zawar & Rampura": 21.2},  # HZL 精炼锌 212kt(+5%)
        },
        "log": ("开 26Q3 列。Vedanta Q2FY27(7-9月)入库：Zinc India 采矿金属 27.1万吨(+5% YoY)→锌矿 Vedanta 行；"
                "HZL 精炼锌 21.2万吨(+5%)→冶炼 Zawar & Rampura 行。Zinc International 采矿 5.3万吨(-12%，"
                "Gamsberg 二期已投产、10 月商业放量)——锌矿表 Zinc India/Zinc International 两行历史值口径存疑，本季暂不写入，待核口径。",
                "Vedanta Q2FY27 Production Release 2026-10-03"),
    },
    "铝": {
        "sheets": {
            "氧化铝·季度产量": {"Vedanta（Lanjigarh）": 89.5},   # 895kt(+37% 纪录)
            "电解铝·季度产量": {"Vedanta": 64.9},                # 649kt(+5% 纪录; Jharsuguda 470+BALCO 178)
        },
        "log": ("开 26Q3 列。Vedanta Q2FY27(7-9月)入库：氧化铝(Lanjigarh) 89.5万吨(+37% YoY，季度纪录，扩产爬坡)；"
                "电解铝 64.9万吨(+5%，季度纪录；Jharsuguda 47.0万吨+BALCO 17.8万吨+19%)。",
                "Vedanta Q2FY27 Production Release 2026-10-03"),
    },
    "铅": {
        "sheets": {
            "精炼铅·季度产量": {"Vedanta/HZL": 5.1},   # 51kt(+14%)
        },
        "log": ("开 26Q3 列。Vedanta/HZL Q2FY27(7-9月)精炼铅 5.1万吨(+14% YoY)入库。"
                "（铅精矿 Vedanta 行维持 '-'：HZL 采矿金属为锌铅合并口径，不单列铅精矿。）",
                "Vedanta Q2FY27 Production Release 2026-10-03"),
    },
}


def open_quarter(ws):
    """在 26Q2同比 后插入 26Q3/26Q3同比 两列，返回新列号。"""
    hdr_row = None
    for i, r in enumerate(ws.iter_rows(min_row=1, max_row=8, values_only=True), 1):
        if r and any(str(v).strip() == "26Q2" for v in r if v):
            hdr_row = i
            hdr = list(r)
            break
    assert hdr_row, f"{ws.title} 找不到表头"
    c_q = [j + 1 for j, v in enumerate(hdr) if str(v).strip() == "26Q2"][0]
    c_y = [j + 1 for j, v in enumerate(hdr) if str(v).strip() == "26Q2同比"][0]
    assert c_y == c_q + 1, f"{ws.title} 26Q2/同比 列不相邻"
    if any(str(v).strip() == "26Q3" for v in hdr if v):
        return c_y + 1  # 幂等：已开过
    ws.insert_cols(c_y + 1, 2)
    for r in range(1, ws.max_row + 1):
        for src_c, dst_c in ((c_q, c_y + 1), (c_y, c_y + 2)):
            s, d = ws.cell(row=r, column=src_c), ws.cell(row=r, column=dst_c)
            d.font, d.fill, d.border = copy(s.font), copy(s.fill), copy(s.border)
            d.alignment, d.number_format = copy(s.alignment), s.number_format
    ws.cell(row=hdr_row, column=c_y + 1, value="26Q3")
    ws.cell(row=hdr_row, column=c_y + 2, value="26Q3同比")
    for letter_src, dst_c in ((openpyxl.utils.get_column_letter(c_q), c_y + 1),
                              (openpyxl.utils.get_column_letter(c_y), c_y + 2)):
        w = ws.column_dimensions[letter_src].width
        if w:
            ws.column_dimensions[openpyxl.utils.get_column_letter(dst_c)].width = w
    return c_y + 1


def fill_and_log(wb, sheets_fill, log_text, log_src):
    for sheet, fills in sheets_fill.items():
        ws = wb[sheet]
        c3 = open_quarter(ws)
        for name, val in fills.items():
            hit = False
            for r in range(2, ws.max_row + 1):
                if str(ws.cell(row=r, column=1).value or "").strip() == name:
                    ws.cell(row=r, column=c3, value=val)
                    hit = True
                    break
            assert hit, f"{ws.title} 找不到行「{name}」"
            print(f"  {sheet} | {name} 26Q3 = {val}")
    ws = wb["更新日志"]
    ws.append(["2026-10-07", log_text, log_src])


def backup(path):
    bak = path.parent / "backup"
    bak.mkdir(exist_ok=True)
    dst = bak / f"{path.stem}_写入前_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    shutil.copy2(path, dst)
    return dst


# ---- 锌：走网关 ----
sys.path.insert(0, str(BASE / "锌"))
from _zinc_write import guarded_update  # noqa: E402

def zinc_edit(ws_dict):
    class FakeWB:  # fill_and_log 只需要 wb[sheet] 索引
        def __getitem__(self, k):
            return ws_dict[k]
    fill_and_log(FakeWB(), FILLS["锌"]["sheets"], FILLS["锌"]["log"][0], FILLS["锌"]["log"][1])

guarded_update(zinc_edit, note="开26Q3列+Vedanta Q2FY27入库")
print("[锌] 网关写入完成")

# ---- 铝/铅：备份后直写 + 同步镜像 ----
for comm in ("铝", "铅"):
    path = BASE / comm / f"全球{comm}企季度产量梳理.xlsx"
    bak = backup(path)
    wb = openpyxl.load_workbook(path)
    fill_and_log(wb, FILLS[comm]["sheets"], FILLS[comm]["log"][0], FILLS[comm]["log"][1])
    wb.save(path)
    print(f"[{comm}] 写入完成（备份 {bak.name}）")

for comm in ("锌", "铝", "铅"):
    src = BASE / comm / f"全球{comm}企季度产量梳理.xlsx"
    for cand in (MIRROR / f"{comm}.xlsx", MIRROR / f"全球{comm}企季度产量梳理.xlsx"):
        if cand.exists():
            shutil.copy2(src, cand)
            print(f"镜像已同步: {cand.name}")
            break
print("全部完成")
