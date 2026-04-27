import json, sys
sys.stdout.reconfigure(encoding='utf-8')
cache = json.load(open('ocr_cache.json'))
for name in list(cache.keys())[:5]:
    print(f'=== {name} ===')
    for s in cache[name].get('subject_wise', []):
        tc = s.get('total_classes', 0)
        ac = s.get('attended_classes', 0)
        pct = round(ac/tc*100, 1) if tc and tc > 0 else 0
        subj = s.get('subject', '')[:55]
        flag = ' <<< OVER 100%!' if ac > tc else ''
        print(f'  {subj} | T:{tc} A:{ac} %:{pct}{flag}')
    print()
