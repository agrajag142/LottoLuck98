"""Builds docs/index.html from tools/template.html, the draw history and the wheel tables.
Run from the repo root:  python3 tools/build.py
(Run tools/test_wheels.py first if wheels.json has changed.)"""
import json, re

wheels = json.load(open('tools/wheels.json'))
tests = json.load(open('tools/wheel_tests.json'))

PRIZE = {6: 'Division 1', 5: 'Division 3', 4: 'Division 4', 3: 'Division 6 (3 numbers)', 2: 'no prize'}
CLAIM = {'1st': 'Division 1', '2nd': 'Division 2', '3rd': 'Division 3', '4th': 'Division 4'}

out = []
for k, v in wheels.items():
    t = tests.get(k)
    if not t or not t['ok']:
        continue  # TenOfFive: incomplete table in the original program
    worst = {int(m): x for m, x in t['worst'].items()}
    original = ' '.join(v['text']).replace("'", '')
    m = re.search(r'(1st|2nd|3rd|4th)', original, re.I)
    claimed = CLAIM[m.group(1).lower()] if m else 'Division 1'
    prize = 'Division 2' if (t['div2'] and worst[6] == 5) else PRIZE[worst[6]]
    combos = [v['combos'][i] for i in sorted(v['combos'], key=int)]
    s = 's' if v['games'] > 1 else ''
    if prize == 'Division 2':
        text = (f"These {v['games']} game{s} from a field of {v['nums']} numbers guarantee at least Division 2 if all six "
                f"winning numbers and a supplementary fall within your field, and at least Division 3 if only the six do.")
    else:
        text = (f"These {v['games']} game{s} from a field of {v['nums']} numbers guarantee at least "
                f"{'a ' if worst[6] != 2 else ''}{prize} if all six winning numbers fall within your field.")
    level = lambda p: int(re.search(r'\d', p).group()) if 'Division' in p else 9
    out.append(dict(id=k, text=text, field=v['nums'], games=v['games'], combos=combos, prize=prize,
                    rank=worst[6], claimed=claimed, overclaimed=level(claimed) < level(prize),
                    worst=worst, dist6={int(a): b for a, b in t['dist6'].items()}, div2=t['div2']))

out.sort(key=lambda x: (-x['rank'], x['field'], x['games']))
draws = open('data/tattslot.txt').read().strip()
games = open('data/number.lot', 'rb').read().decode().replace('\n\r', '\n').replace('\r', '').strip()
tpl = open('tools/template.html').read()
html = (tpl.replace('/*__DRAWS__*/', draws)
           .replace('/*__WHEELS__*/', json.dumps(out, separators=(',', ':')))
           .replace('/*__GAMES__*/', games))
open('docs/index.html', 'w').write(html)
print(f"docs/index.html: {len(html):,} bytes, {len(out)} wheels, {draws.count(chr(10))} draws")
