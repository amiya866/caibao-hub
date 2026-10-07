import json, sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

BASE = r'D:\Kimi\caibao-hub\data'

# ---- news.json ----
news = json.load(open(f'{BASE}\\news.json', encoding='utf-8'))
patch = json.load(open(f'{BASE}\\_news_patch_20261007.json', encoding='utf-8'))
exist = {(n['date'], n['title']) for n in news}
added = [n for n in patch if (n['date'], n['title']) not in exist]
news.extend(added)
json.dump(news, open(f'{BASE}\\news.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print(f'news: {len(news)-len(added)} + {len(added)} = {len(news)}')

# ---- disruptions.json ----
d = json.load(open(f'{BASE}\\disruptions.json', encoding='utf-8'))
items = d['items']
# Alba 恢复进展更新
for it in items:
    if 'Alba' in it.get('company', ''):
        it['recovery'] = ('2026-09-30：1-3号受损产线（约31万吨）设施修复完毕，电解槽逐渐启动；'
                          '4-6号线正常运行，8月年化产量130万吨；欧美现货升水反弹、海外偏紧')
        it['type'] = '地缘受损→修复重启'
        print('disruptions: Alba 行已更新')
# WBN 复产行（无则加）
if not any('Weda Bay' in it.get('company', '') or 'WBN' in it.get('company', '') for it in items):
    items.append({
        "commodity": "镍", "date": "2026-09-10", "ongoing": False,
        "company": "Eramet / PT Weda Bay Nickel", "country": "印尼",
        "type": "配额耗尽停产→获批复产", "dir": "up",
        "capacity": "2026 初始 RKAB 12Mwmt（较 2025 年 42Mwmt -70%）",
        "impact": "5 月起矿山转保养停产约 4 个月；9/10 获批逐步重启采矿（2026 补充 RKAB，市场报道约 2,500 万湿吨）",
        "recovery": "渐进复产中，全年产销量指引待 Eramet 更新；2027 RKAB 申请窗口 10/1-11/15",
        "source": "Eramet 官网 2026-09-10"})
    print('disruptions: WBN 复产行已加')
d['updated'] = '2026-10-07'
json.dump(d, open(f'{BASE}\\disruptions.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('disruptions 条数:', len(items))
