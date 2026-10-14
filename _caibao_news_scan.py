# -*- coding: utf-8 -*-
"""caibao-hub 信息速递自动扫描（2026-10-14 建，Windows 任务 CaibaoNewsScan 每日 4 档）
背景：周扫描一直靠 Kimi 会话 cron，会话消亡即断更（9/4 断到 10/7、10/12 又漏）——改 Windows 任务常驻。
数据源：整合研究站快讯快照 D:\\Kimi\\整合研究站_中哲复制\\site\\data\\news_latest.json
        （RwBenchQuick 每 2h 刷新，含金十/CS财经/东财7x24/新浪7x24/westmetall）
规则：商品关键词必须命中【】标题（命中正文会误收汇编，10-14 已在 geo_events 验证过）；
      去重=已有标题前12字+URL+同快照跨源重复；⚙️标注自动收录，impact 启发式（人工可改）；
      每次最多 8 条、单品种最多 3 条；只收近 3 天。有新增才 build_site + 推送。
用法: python _caibao_news_scan.py [--dry]
"""
import io, os, re, sys, json, subprocess, datetime

BASE = os.path.dirname(os.path.abspath(__file__))
NEWS_JSON = os.path.join(BASE, 'data', 'news.json')
SNAPSHOT = r'D:\Kimi\整合研究站_中哲复制\site\data\news_latest.json'
LOG = os.path.join(BASE, '_news_scan.log')
MAX_PER_RUN, MAX_PER_COMM, DAYS = 8, 3, 3

COMM_KW = {
    '铜': re.compile(r'铜'),
    '铝': re.compile(r'铝|氧化铝|铝土矿'),
    '锌': re.compile(r'锌'),
    '铅': re.compile(r'铅'),
    '镍': re.compile(r'镍|NPI|印尼.{0,6}(镍|RKAB|配额)'),
    '锡': re.compile(r'锡'),
    '锂': re.compile(r'锂'),
    '硅': re.compile(r'多晶硅|工业硅|硅料|硅片|有机硅'),
    '宏观': re.compile(r'美联储|FOMC|加息|降息|非农|CPI|PCE|美债|黄金|金价|霍尔木兹|红海|原油|油价|关税|美元指数'),
}
# 宏观之外的品种标题若只是价格快讯且无关键词，noise 偏多——要求标题带基本面/政策词才收
FUND_KW = re.compile(r'停产|复产|减产|增产|事故|罢工|矿|冶炼|产量|库存|仓单|开工|产能|出口|进口|关税|政策|配额|'
                     r'加工费|TC|升水|贴水|涨跌|上涨|下跌|创新高|创新低|财报|业绩|盈利|亏损|投产|检修|需求|供应')
HIGH_KW = re.compile(r'停产|事故|火灾|爆炸|罢工|封锁|加征|加息|降息|超预期|暴涨|暴跌|创新高|创新低|断供|危机')
MID_KW = re.compile(r'复产|减产|增产|政策|关税|配额|出口|进口|库存|产量|开工|投产|检修|财报|业绩|指引')


def title_of(content):
    m = re.match(r'\s*【(.{4,80}?)】', content or '')
    return m.group(1) if m else ''  # 无【】标题的快讯不收（多为碎片）


def classify(comm, text):
    if comm == '宏观':
        return '宏观'
    if re.search(r'停产|复产|减产|增产|事故|罢工|矿|冶炼|开工|产能|投产|检修|供应', text):
        return '供应'
    if re.search(r'需求|消费|订单|排产|下游', text):
        return '需求'
    if re.search(r'政策|关税|配额|出口管制|补贴', text):
        return '政策'
    return '公司'


def impact_of(text):
    if HIGH_KW.search(text):
        return 'high'
    if MID_KW.search(text):
        return 'mid'
    return 'low'


def norm(s):
    return re.sub(r'\s+', '', s or '')


def main():
    dry = '--dry' in sys.argv
    news = json.load(io.open(NEWS_JSON, encoding='utf-8'))
    snap = json.load(io.open(SNAPSHOT, encoding='utf-8'))
    cutoff = (datetime.date.today() - datetime.timedelta(days=DAYS)).isoformat()
    exist_titles = {norm(it.get('title'))[:12] for it in news}
    exist_urls = {it.get('url', '') for it in news if it.get('url')}

    items = []
    for k in ('cs', 'eastmoney', 'sina', 'jin10', 'lme'):
        for it in (snap.get(k) or []):
            t, c = it.get('time', ''), it.get('content', '')
            if t[:10] >= cutoff and c:
                items.append((t, c))
    items.sort(key=lambda x: x[0], reverse=True)

    added, per_comm, seen = [], {}, set()
    for t, c in items:
        if len(added) >= MAX_PER_RUN:
            break
        title = title_of(c)
        if not title:
            continue
        comm = None
        for cm, kw in COMM_KW.items():
            if kw.search(title):
                comm = cm
                break
        if not comm:
            continue
        if comm != '宏观' and not FUND_KW.search(title):
            continue
        if per_comm.get(comm, 0) >= MAX_PER_COMM:
            continue
        key = norm(title)[:12]
        if key in exist_titles or key in seen:
            continue
        body = norm(re.sub(r'^\s*【.{4,80}?】', '', c))[:120]
        src = {'cs': 'CS财经', 'eastmoney': '东方财富7x24', 'sina': '新浪7x24',
               'jin10': '金十数据', 'lme': 'westmetall'}.get('cs')
        added.append({
            'date': t[:10],
            'commodity': comm,
            'category': classify(comm, title + body),
            'impact': impact_of(title + body),
            'title': title[:60],
            'summary': (body + '…' if body else title) + '（⚙️快讯自动收录·%s）' % src,
            'source': src + '快讯',
            'url': '',
        })
        seen.add(key)
        per_comm[comm] = per_comm.get(comm, 0) + 1

    line = '%s added=%d %s' % (datetime.datetime.now().strftime('%F %T'), len(added),
                               [(a['commodity'], a['title'][:18]) for a in added])
    with io.open(LOG, 'a', encoding='utf-8') as f:
        f.write(line + '\n')
    print(line)
    if not added or dry:
        return
    news.extend(added)
    json.dump(news, io.open(NEWS_JSON, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    # 重建 + 推送（失败留到下一档重试，news.json 已落盘不丢）
    r1 = subprocess.run([sys.executable, os.path.join(BASE, 'build_site.py')],
                        cwd=BASE, capture_output=True, text=True, encoding='utf-8',
                        errors='replace', timeout=600)
    print('build rc=%d' % r1.returncode, (r1.stdout or '')[-200:], (r1.stderr or '')[-200:])
    if r1.returncode == 0:
        r2 = subprocess.run([sys.executable, os.path.join(BASE, '_gh_push_caibao_hub.py')],
                            cwd=BASE, capture_output=True, text=True, encoding='utf-8',
                            errors='replace', timeout=900)
        print('push rc=%d' % r2.returncode, (r2.stdout or '')[-200:], (r2.stderr or '')[-200:])


if __name__ == '__main__':
    main()
