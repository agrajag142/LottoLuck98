import json, re
w = json.load(open('tools/wheels.json'))
out = []
names = {'1st':'Division 1','2nd':'Division 2','3rd':'Division 3','4th':'Division 4'}
for k, v in w.items():
    if k == 'TenOfFive': continue
    t = ' '.join(v['text']).replace("'", '')
    m = re.search(r'(1st|2nd|3rd|4th)', t, re.I)
    prize = names[m.group(1).lower()] if m else 'Division 1'
    combos = [v['combos'][i] for i in sorted(v['combos'], key=int)]
    assert len(combos) == v['games']
    text = f"These {v['games']} combination{'s' if v['games']>1 else ''} of a field of {v['nums']} numbers guarantee at least a {prize} win if all six numbers drawn fall within your field."
    out.append(dict(id=k, text=text, field=v['nums'], games=v['games'], combos=combos, prize=prize))
order = ['Division 1','Division 2','Division 3','Division 4']
out.sort(key=lambda x: (order.index(x['prize']), x['field'], x['games']))
draws = open('data/tattslot.txt').read().strip()
games = open('data/number.lot', 'rb').read().decode().replace('\n\r', '\n').replace('\r','').strip()
tpl = open('tools/template.html').read()
html = tpl.replace('/*__DRAWS__*/', draws).replace('/*__WHEELS__*/', json.dumps(out, separators=(',',':'))).replace('/*__GAMES__*/', games)
open('web/index.html', 'w').write(html)
print(len(html), len(out), draws.count('\n'))
