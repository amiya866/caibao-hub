# -*- coding: utf-8 -*-
"""caibao-hub 工作流 db 生成器（2026-09-02）
把财报入库平台的真源/入库链/脚本/定时任务/军规结构化存进 caibao_hub_工作流.db，供 agent 直接查询。
重跑即刷新：python _build_workflow_db.py
"""
import io, os, sqlite3, datetime, sys
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
HERE = Path(__file__).resolve().parent
OUT = HERE / "caibao_hub_工作流.db"

META = {
    "site": "caibao-hub 财报入库平台（全球有色产业财报梳理跟踪）",
    "url": "https://amiya866.github.io/caibao-hub/",
    "repo": "amiya866/caibao-hub（GitHub Pages，无密码门）",
    "source_root": r"D:\Kimi\caibao-hub",
    "builder": "_build_workflow_db.py v1",
}

DATABASES = [
    ("站点数据契约", r"D:\Kimi\caibao-hub\data\data.js", "js", "window.SITE_DATA（契约不动）", "build_site.py 生成", "UI 改动只碰 index.html/assets，改完必须跑 build_site.py 更新哈希引用"),
    ("品种Excel真库", r"D:\Kimi\金属总网\网站构建\财报汇总\{铜铝锌镍锡锂硅铅}", "xlsx", "8 品种唯一真源（2025-08-25 起）", "财报入库流程", "永安目录降级为存档、禁止编辑；真库优先于 excel\\ 本地副本"),
    ("财报数据集db", r"D:\Kimi\金属总网\网站构建\财报汇总\财报数据集.db", "sqlite", "8 品种结构化数据集", "财报汇总工作流", ""),
    ("新闻与扰动", r"D:\Kimi\caibao-hub\data\news.json / disruptions.json", "json", "信息速递+供应扰动表", "追加条目后 build_site.py 重建", "news 字段 date/commodity/category/title/summary/source/url/impact/affects；扰动减=红增=绿"),
    ("UI 设计真源", r"D:\大胖鱼\财报汇总\独立面板\index.html", "html", "Caibao 财报面板 v5 侧栏形态（2026-08-21 用户指定）", "—", "黑顶栏+左侧栏 ticker chip+品种色；设计文档在 caibao-collect_品种SOP_v5.md 第七节"),
    ("永安存档", r"D:\拷贝文件\E\永安\{品种}", "xlsx", "历史存档（只读）", "不同步写入", "_sync_caibao.py 反向分发用"),
]

STEPS = [
    (1, "核财报", "联网/公告", "财报节点联网核官方财报来源", "公司官网/交易所/SEC", "核实数据+口径备注", "禁止编造，找不到标「未找到」，推算值注明"),
    (2, "写Excel", r"财报汇总\{品种}\ 真库", "写入品种 Excel 真库", "核实数据", "真库更新", "锌表必须走 财报汇总\\锌\\_zinc_write.py 网关（guarded_update）；写前先备份；编辑前关 Excel 以磁盘为准"),
    (3, "build", r"build_site.py", "读真库重建 data.js", "品种真库 Excel", "data/data.js + 披露日历三档状态", "公司名匹配走 _curq.py ALIASES；真库优先于 excel\\ 副本"),
    (4, "push", r"_gh_push_caibao_hub.py", "Trees API 全量推 amiya866/caibao-hub", "站点文件", "线上 Pages", "device token；fine-grained PAT 不覆盖新仓"),
    (5, "目检", "—", "打开线上站点目检", "—", "—", "交付时主动打开给用户检查"),
]

SCRIPTS = [
    (r"build_site.py", "站点构建", "COMMODITIES 注册表扩展新品种；披露日历三档（待披露灰/已入库绿/已披露待核·未入库橙）"),
    (r"_caibao_watch.py", "守望者 v2", "扫 8 品种真库「上季有数当季空」+ 披露日历×真库交叉核对 → _待核清单.md + 弹窗；CUR_Q 由 _curq.py 自动推算勿手调"),
    (r"_curq.py", "当季推算", "Q1 季 4/20、Q2 季 8/10、Q3 季 10/25、Q4 季次年 2/20 起核；含公司名 ALIASES"),
    (r"_gh_push_caibao_hub.py", "部署", "Trees API 全量推"),
    (r"..\金属总网\网站构建\财报汇总\工作流\run_all_check.py", "体检 13 项", "含真库防覆盖行数基线（只增不减，缩水=FAIL）"),
    (r"..\金属总网\网站构建\财报汇总\锌\_zinc_write.py", "锌表写入网关", "guarded_update"),
    (r"..\金属总网\网站构建\财报汇总\锌\_zinc_guard.py", "锌表哨兵", "每日 09:05 丢行自动恢复"),
    (r"..\金属总网\网站构建\财报汇总\_sync_caibao.py", "存档反向分发", "真库→永安存档"),
]

SCHEDULES = [
    ("CaibaoWatch", "每周一 09:11", "python D:\\Kimi\\caibao-hub\\_caibao_watch.py", "Windows 任务；只提醒不执行扫描"),
    ("Kimi cron 信息速递周扫描", "每周一 09:41（cron 41 9 * * 1）", "Kimi 会话级 cron", "扫描实际执行者：news.json/disruptions.json 补录→build→push；⚠️会话消亡即失效，新会话按册07 F5 存档 prompt 重建（2026-09-05 断更事件后立）"),
    ("ZincTableGuard", "每日 09:05", "python 财报汇总\\锌\\_zinc_guard.py", "Windows 任务（UTF-8-BOM ps1 重注册修复过乱码）"),
]

RULES = [
    ("事故教训", "锌表四次覆盖事故：旧内存副本覆盖保存静默丢行——编辑锌表前关 Excel、以磁盘为准、写入走 _zinc_write.py 网关、每日哨兵兜底"),
    ("军规", "缺季拟合：有年报总量缺季度值时用平均/季节性拟合填入，前端斜体异色+备注注明拟合方法"),
    ("军规", "产量指引全部入卡片并算年化进度 vs 指引；新品种从 2023Q1 起建序列"),
    ("军规", "突发供应事件收录 news.json 时同步更新 disruptions.json（恢复状态变化改 recovery 列），重跑 build_site.py 上线"),
    ("军规", "事故减产年公司 2027 展望人工覆盖 FY2027_OUTLOOK 为恢复性增长（标「事件」），不靠线性外推"),
    ("分工", "披露季欠账=守望者提醒、Kimi 会话内执行入库（联网核财报→写 Excel→build→push）"),
    ("口径", "数据账号已有：Mysteel 钢联终端/百川盈孚/zhiji API（含 SMM 全系）/Wind/iFinD 公司侧；文档标「需要购买」前先核对"),
    ("结构", "工作流八册+部署包+品种检索索引在 财报汇总\\工作流\\；品种 SOP=caibao-collect skill（~/.agents/skills/caibao-collect）"),
]

LINKS = [
    ("线上站点", "https://amiya866.github.io/caibao-hub/"),
    ("金属总网（总台）", "https://amiya866.github.io/metals-framework/"),
    ("品种 SOP", r"D:\Kimi\金属总网\网站构建\财报汇总\caibao-collect_品种SOP_v5.md"),
    ("工作流八册", r"D:\Kimi\金属总网\网站构建\财报汇总\工作流"),
]


def main():
    now = datetime.datetime.now().isoformat(timespec="seconds")
    if OUT.exists():
        OUT.unlink()
    con = sqlite3.connect(OUT)
    cur = con.cursor()
    cur.execute("CREATE TABLE meta(k TEXT PRIMARY KEY, v TEXT)")
    cur.execute("CREATE TABLE databases(name TEXT, path TEXT, type TEXT, caliber TEXT, update_method TEXT, notes TEXT)")
    cur.execute("CREATE TABLE steps(ord INT, key TEXT, script TEXT, purpose TEXT, inputs TEXT, outputs TEXT, pitfalls TEXT)")
    cur.execute("CREATE TABLE scripts(path TEXT, role TEXT, notes TEXT)")
    cur.execute("CREATE TABLE schedules(task TEXT, trigger TEXT, command TEXT, notes TEXT)")
    cur.execute("CREATE TABLE rules(category TEXT, rule TEXT)")
    cur.execute("CREATE TABLE links(name TEXT, url TEXT)")
    for k, v in META.items():
        cur.execute("INSERT INTO meta VALUES(?,?)", (k, v))
    cur.execute("INSERT INTO meta VALUES('built_at', ?)", (now,))
    cur.executemany("INSERT INTO databases VALUES(?,?,?,?,?,?)", DATABASES)
    cur.executemany("INSERT INTO steps VALUES(?,?,?,?,?,?,?)", STEPS)
    cur.executemany("INSERT INTO scripts VALUES(?,?,?)", SCRIPTS)
    cur.executemany("INSERT INTO schedules VALUES(?,?,?,?)", SCHEDULES)
    cur.executemany("INSERT INTO rules VALUES(?,?)", RULES)
    cur.executemany("INSERT INTO links VALUES(?,?)", LINKS)
    con.commit()
    con.close()
    print(f"DB -> {OUT}（{OUT.stat().st_size//1024} KB；databases {len(DATABASES)}/steps {len(STEPS)}/scripts {len(SCRIPTS)}/rules {len(RULES)}）")


if __name__ == "__main__":
    main()
