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
    "agent_guide": "本库是 caibao-hub 财报入库平台的工作流地图（入库链路/脚本/定时任务/军规）。用法：①找数据资产先查 D:\\Kimi\\金属总网\\网站构建\\db\\数据资产目录_v1.db 的 assets 表；②季度产量数据本体在 8 品种 Excel 真库（网站构建\\财报汇总\\{品种}\\，写锌表必须走 财报汇总\\锌\\_zinc_write.py 网关），结构化快照在 网站构建\\db\\财报数据集.db（SELECT * FROM dataset WHERE comm='锌'）；③入库链路 6 步看 steps 表，调度看 schedules 表（守望者只提醒、入库由 Kimi 会话执行）；④zhiji 断供/断网时指标数据直接 SQL 查 网站构建\\db\\数据缓存_v1.db 的 series_points（取法见 databases 表「数据缓存db」行）；⑤改本库内容只改本脚本，跑 D:\\Kimi\\金属总网\\_refresh_workflow_dbs.py 或等每日 09:21 KimiWorkflowDbRefresh 自动重建。",
}

DATABASES = [
    ("站点数据契约", r"D:\Kimi\caibao-hub\data\data.js", "js", "window.SITE_DATA（契约不动）", "build_site.py 生成", "UI 改动只碰 index.html/assets，改完必须跑 build_site.py 更新哈希引用"),
    ("品种Excel真库", r"D:\Kimi\金属总网\网站构建\财报汇总\{铜铝锌镍锡锂硅铅}", "xlsx", "8 品种唯一真源（2025-08-25 起）", "财报入库流程", "永安目录降级为存档、禁止编辑；真库优先于 excel\\ 本地副本"),
    ("财报数据集db", r"D:\Kimi\金属总网\网站构建\db\财报数据集.db", "sqlite", "8 品种结构化数据集", r"财报汇总\_caibao_export.py 全量重建", ""),
    ("新闻与扰动", r"D:\Kimi\caibao-hub\data\news.json / disruptions.json", "json", "信息速递+供应扰动表", "追加条目后 build_site.py 重建", "news 字段 date/commodity/category/title/summary/source/url/impact/affects；扰动减=红增=绿"),
    ("UI 设计真源", r"D:\大胖鱼\财报汇总\独立面板\index.html", "html", "Caibao 财报面板 v5 侧栏形态（2026-08-21 用户指定）", "—", "黑顶栏+左侧栏 ticker chip+品种色；设计文档在 caibao-collect_品种SOP_v5.md 第七节"),
    ("永安存档", r"D:\拷贝文件\E\永安\{品种}", "xlsx", "历史存档（只读）", "不同步写入", "_sync_caibao.py 反向分发用"),
    ("caibao_hub_工作流db", r"D:\Kimi\caibao-hub\caibao_hub_工作流.db", "sqlite", "meta/databases/steps/scripts/schedules/rules/links（本库）", r"caibao-hub\_build_workflow_db.py 全量重建", "自描述库；agent 了解本平台先看这里"),
    ("数据缓存db", r"D:\Kimi\金属总网\网站构建\db\数据缓存_v1.db", "sqlite", "series_meta(240 指标)/series_points(ind_id,date,value 约12万点)", "_build_data_cache.py 增量（每日 09:21）", "断网/异地取数：SELECT date,value FROM series_points WHERE ind_id='指标ID' ORDER BY date；指标清单查 series_meta；zhiji 全镜像 2020 起"),
]

STEPS = [
    (1, "核财报", "联网/公告", "财报节点联网核官方财报来源", "公司官网/交易所/SEC", "核实数据+口径备注", "禁止编造，找不到标「未找到」，推算值注明"),
    (2, "写Excel", r"财报汇总\{品种}\ 真库", "写入品种 Excel 真库", "核实数据", "真库更新", "锌表必须走 财报汇总\\锌\\_zinc_write.py 网关（guarded_update）；写前先备份；编辑前关 Excel 以磁盘为准"),
    (3, "build", r"build_site.py", "读真库重建 data.js", "品种真库 Excel", "data/data.js + 披露日历三档状态", "公司名匹配走 _curq.py ALIASES；真库优先于 excel\\ 副本"),
    (4, "push", r"_gh_push_caibao_hub.py", "Trees API 全量推 amiya866/caibao-hub", "站点文件", "线上 Pages", "device token；fine-grained PAT 不覆盖新仓"),
    (5, "目检", "—", "打开线上站点目检", "—", "—", "交付时主动打开给用户检查"),
    (6, "信息速递周扫描", "Kimi cron 周一 09:41", "联网扫描 8 品种近一周可核实事件 → 追加 news.json + 同步 disruptions.json → build → push", "印尼镍/天下铝讯/阿拉丁(yaqh)/爱择/传言/公司公告", "news/disruptions 增量上线", "CaibaoWatch(周一09:11)只提醒不执行；cron 会话级消亡即失效，新会话按册07 F5 重建；schema/传言规范见册03 第五节"),
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
    ("CaibaoWatch", "每周一 09:11", r'"C:\Users\Yitian Shen\AppData\Local\Programs\Python\Python314\python.exe" "D:\Kimi\caibao-hub\_caibao_watch.py"', "Windows 任务；只提醒不执行扫描"),
    ("Kimi cron 信息速递周扫描", "每周一 09:41（cron 41 9 * * 1）", "Kimi 会话级 cron", "扫描实际执行者：news.json/disruptions.json 补录→build→push；⚠️会话消亡即失效，新会话按册07 F5 存档 prompt 重建（2026-09-05 断更事件后立）"),
    ("ZincTableGuard", "每日 09:05", r'"C:\Users\Yitian Shen\AppData\Local\Programs\Python\Python314\python.exe" "D:\Kimi\金属总网\网站构建\财报汇总\锌\_zinc_guard.py"', "Windows 任务（UTF-8-BOM ps1 重注册修复过乱码）；锌表丢行自动恢复"),
    ("MetalsFrameworkDailyUpdate", "每天 08:47（StartWhenAvailable 错过补跑）", r'"C:\Users\Yitian Shen\AppData\Local\Programs\Python\Python314\python.exe" "D:\Kimi\金属总网\_daily_update.py"', "Windows 任务（总台侧）；含 zsxq 知识星球步；日志 金属总网\\_daily_update.log"),
    ("KimiWorkflowDbRefresh", "每天 09:21", r'"C:\Users\Yitian Shen\AppData\Local\Programs\Python\Python314\python.exe" "D:\Kimi\金属总网\_refresh_workflow_dbs.py"', "Windows 任务；重建本库+金属总网工作流db+数据资产目录db+数据缓存db（缓存步失败容错跳过）"),
    ("ZsxqFeishuPush", "每 20 分钟（Repetition PT20M）", r'"C:\Users\Yitian Shen\AppData\Local\Programs\Python\Python314\python.exe" "D:\Kimi\金属总网\网站构建\小作文平台\知识星球\_zsxq_dynamics.py"', r"Windows 任务（总台侧）；知识星球增量→飞书"),
    ("ZsxqSiteRealtime", "每 20 分钟（偏移 10 分）", r'"C:\Users\Yitian Shen\AppData\Local\Programs\Python\Python314\python.exe" "D:\Kimi\金属总网\网站构建\小作文平台\知识星球\_zsxq_realtime.py"', r"Windows 任务（总台侧）；近 3h 主题入库 zsxq_metals.json+注入 HTML+部署"),
]

RULES = [
    ("事故教训", "锌表四次覆盖事故：旧内存副本覆盖保存静默丢行——编辑锌表前关 Excel、以磁盘为准、写入走 _zinc_write.py 网关、每日哨兵兜底"),
    ("军规", "缺季拟合：有年报总量缺季度值时用平均/季节性拟合填入，前端斜体异色+备注注明拟合方法"),
    ("军规", "产量指引全部入卡片并算年化进度 vs 指引；新品种从 2023Q1 起建序列"),
    ("军规", "突发供应事件收录 news.json 时同步更新 disruptions.json（恢复状态变化改 recovery 列），重跑 build_site.py 上线"),
    ("军规", "事故减产年公司 2027 展望人工覆盖 FY2027_OUTLOOK 为恢复性增长（标「事件」），不靠线性外推"),
    ("分工", "披露季欠账=守望者提醒、Kimi 会话内执行入库（联网核财报→写 Excel→build→push）"),
    ("事故教训", "2026-09-05 信息速递断更（停在 08-29）：扫描执行者是 Kimi 会话级 cron（旧 id 51265eed），随旧会话消亡；CaibaoWatch 只弹窗提醒不执行——新会话开工先 CronList 检查「信息速递周扫描」cron（周一 09:41），缺失按册07 F5 存档 prompt 重建"),
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
