# -*- coding: utf-8 -*-
"""2026-08-31 信息速递双周扫描入库：追加 news.json 8 条 + 更新 disruptions.json 银漫 recovery。"""
import json, io, sys

sys.stdout.reconfigure(encoding='utf-8')
NEWS = r'D:\Kimi\caibao-hub\data\news.json'
DISR = r'D:\Kimi\caibao-hub\data\disruptions.json'

new_items = [
 {
  "date": "2026-08-29",
  "commodity": "宏观",
  "category": "宏观",
  "impact": "high",
  "title": "美联储主席沃什 Jackson Hole 首秀偏鹰：9月大概率按兵不动，12月前加息概率约70%",
  "summary": "沃什8/28在Jackson Hole年会发表就任后首次主旨演讲，聚焦强化抗通胀信誉、未给出明确反应函数。会前市场对9月加息的定价已从7月中约82%降至约30%；会后主流预期9月不动，但中东局势推高油价、通胀黏性下，交易员定价12月前至少加息25bp概率约70%（路透）。金属宏观面从「降息预期」切换为「加息尾部风险」，美元获支撑，与上半年宽松基调反转。",
  "source": "Morningstar/Reuters",
  "url": "https://www.morningstar.com/economy/warsh-sounds-hawkish-will-there-be-september-rate-hike"
 },
 {
  "date": "2026-08-28",
  "commodity": "锡",
  "category": "公司",
  "impact": "high",
  "title": "兴业银锡半年报确认银漫仍全面停产（采矿+选矿尾矿），H1归母净利预增169-198%靠宇邦+价格对冲",
  "summary": "兴业银锡2026半年报（8/28披露）：7/26事故（1人死亡）后银漫矿业收西乌旗应急局257/260号决定书，截至目前采矿系统、选矿尾矿系统均停产，事故原因仍在调查，复产须经事故调查与安监验收、时间未定。H1业绩预告归母净利21.4-23.7亿元（+169%~198%），由宇邦矿业产能释放、银锡价格上涨及转让收益贡献。银漫2025年营收占合并55%、净利占近八成；按2025年月均约557吨（26Q1月均259吨）计，全面停产每月矿产锡约-260~560吨。",
  "source": "兴业银锡2026年半年度报告（巨潮/新浪公告）",
  "url": "https://money.finance.sina.com.cn/corp/view/vCB_AllBulletinDetail.php?stockid=000426&id=12570967",
  "affects": [
   {"company": "兴业银锡", "note": "半年报确认银漫采矿+选矿尾矿仍全面停产、复产时间未定；H1归母预增169-198%（宇邦放量+银锡涨价对冲）"}
  ]
 },
 {
  "date": "2026-08-27",
  "commodity": "锌",
  "category": "供应",
  "impact": "high",
  "title": "伦锌盘中逼近4000美元：LME库存9.51万吨（较6月底-20.7%），1.7万吨交仓约2/3入香港、出口窗口关闭",
  "summary": "8/27伦锌盘中逼近4000美元/吨，近端供应结构仍偏紧：LME锌库存约9.51万吨、较6月底下降20.7%，仓单持有集中度仍高。近期约1.7万吨锌锭交入LME仓库，其中约三分之二流入香港仓库；国内锌锭流出+海外交仓后出口窗口关闭，挤仓情绪阶段缓和但伦锌back结构继续走强（创元期货8/23）。挤仓从「库存枯竭」转入「低库存+高集中度」僵持阶段。",
  "source": "新浪财经/创元期货",
  "url": "https://finance.sina.com.cn/money/future/fmnews/2026-08-27/doc-inipsmim5893512.shtml"
 },
 {
  "date": "2026-08-20",
  "commodity": "镍",
  "category": "政策",
  "impact": "high",
  "title": "ESDM确认2026镍矿RKAB配额不再增加（APNI 8/19宣布）；印尼今年暂不对镍征出口税/暴利税",
  "summary": "据钢联/光大期货有色日报：8/19印尼镍矿商协会APNI宣布ESDM已正式确认2026年镍矿RKAB配额不再增加；ESDM局长Tri Winarno同日称已开始逐企批准2026年煤炭和镍RKAB修订（延续「仅限供应短缺冶炼厂」口径）。经济统筹部长艾尔朗加表示今年暂不对煤炭和镍征收出口税和暴利税，明年政策尚未讨论。另据ESDM预估2026年镍矿产量约2.09亿吨。年中配额博弈落定偏紧，沪镍主力12.8万附近震荡（8/19收127,850）。",
  "source": "光大期货有色日报（据钢联）/APNI",
  "url": "https://finance.sina.com.cn/money/future/fmnews/2026-08-26/doc-iniprcpz6239464.shtml"
 },
 {
  "date": "2026-08-20",
  "commodity": "锡",
  "category": "公司",
  "impact": "mid",
  "title": "锡业股份半年报：H1有色金属总产量19.34万吨（锡5.18万吨），营收同比+49.68%",
  "summary": "云南锡业2026半年报（8/20披露）：H1有色金属总产量19.34万吨，其中锡5.18万吨、铜6.92万吨、锌7.23万吨；营收同比+49.68%。分季度看锡Q1为2.59万吨、Q2约2.59万吨环比持平——高锡价下以量稳兑现业绩弹性。",
  "source": "锡业股份2026年半年度报告/搜狐财经",
  "url": "https://www.sohu.com/a/1068502079_122014422",
  "affects": [
   {"company": "云南锡业", "note": "H1锡产量5.18万吨（Q2约2.59万吨环比持平）；营收同比+49.68%"}
  ]
 },
 {
  "date": "2026-08-18",
  "commodity": "锌",
  "category": "供应",
  "impact": "mid",
  "title": "ILZSG将2026年锌平衡大幅调转30万吨至短缺，LME库存降至9.5万吨；中国需求疲弱压制上方",
  "summary": "Crux Investor（8/18）：ILZSG把2026年锌供需平衡下调30万吨、由过剩转为短缺，LME库存降至约9.5万吨；但中国需求疲弱限制锌价上行空间——短缺叙事与弱需求并存。",
  "source": "Crux Investor",
  "url": "https://www.cruxinvestor.com/posts/zinc-flips-into-deficit-as-lme-stocks-hit-95-000-tonnes-despite-weak-demand"
 },
 {
  "date": "2026-08-13",
  "commodity": "镍",
  "category": "供应",
  "impact": "mid",
  "title": "印尼镍铁让电电解铝进入执行：IWIP 3×380MW机组H2分批交付，印尼年内铝供应增量约73万吨",
  "summary": "上海金属网（8/13，上轮扫描遗漏补录）：电解铝利润远高于镍铁（SMM测算园区单位电力毛利贡献电解铝为NPI的数十倍），园区镍铁减产让电、电力优先供给电解铝，印尼电解铝项目得以快于预期落地；IWIP园区3×380MW电力机组2026H2起分批交付，电力瓶颈将妥善解决，印尼年内预计贡献铝供应增量73万吨。8月调研遗留的「镍铁让电」变量已从协商进入执行：对镍铁供应为边际减量、对铝为加速增量。",
  "source": "上海金属网/SMM",
  "url": "https://www.shmet.com/news/newsDetail-2-921552.html"
 },
 {
  "date": "2026-08-07",
  "commodity": "锡",
  "category": "公司",
  "impact": "mid",
  "title": "Minsur 26Q2精炼锡7,034吨（同比-1%、环比-15.4%检修），H1累计15,349吨（-2%）",
  "summary": "SMM锡快讯（8/7，上轮扫描遗漏补录）：Minsur秘鲁Pisco冶炼厂26Q2精炼锡7,034吨，同比-1%、环比-15.4%（检修），入选品位下降被处理量提升部分对冲；H1累计15,349吨（同比-2%）。San Rafael地下矿Q2锡金属产量7,359吨。",
  "source": "SMM锡快讯/ITA",
  "url": "https://news.metal.com/en/newscontent/104036828-smm-tin-express-minsur-q2-refined-tin-production-7034-mt-yoy-down-1",
  "affects": [
   {"company": "明苏尔", "note": "26Q2精炼锡7,034吨（同比-1%）；H1累计15,349吨（-2%），检修+品位下降"}
  ]
 }
]

# URL 去重
news = json.load(open(NEWS, encoding='utf-8'))
existing_urls = {x.get('url', '') for x in news}
added = [x for x in new_items if x['url'] not in existing_urls]
skipped = [x['title'][:30] for x in new_items if x['url'] in existing_urls]
if skipped:
    print('SKIP(url dup):', skipped)
news = added + news
json.dump(news, open(NEWS, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print(f'news.json: {len(news)-len(added)} + {len(added)} = {len(news)}')

# disruptions.json：更新银漫 recovery
disr = json.load(open(DISR, encoding='utf-8'))
for it in disr['items']:
    if '银漫' in it['company']:
        it['recovery'] = '截至8/28半年报：采矿+选矿尾矿系统仍全面停产，事故原因调查中，复产须经安监验收、时间未定；2027恢复性增长至事故前水平'
        print('disruptions 银漫 recovery 已更新')
disr['updated'] = '2026-08-31'
json.dump(disr, open(DISR, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('disruptions.json updated =', disr['updated'], 'items =', len(disr['items']))
