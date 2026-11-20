"""Exhaustively tests every wheel in tools/wheels.json.

For each wheel and for m = 3..6 (how many of the six winning numbers land inside your
field), it tries every possible set of m numbers from the field and records the worst
case: the most numbers that the best game of the wheel is sure to match. It also records,
for the case where all six winning numbers are in the field, how often the best game
matches 6, 5, 4 or 3, and whether Division 2 (5 + supp) is guaranteed when the supp is
in the field too.

Writes tools/wheel_tests.json, which build.py reads.  Needs numpy.
Run from the repo root:  python3 tools/test_wheels.py
"""
import itertools, json, re
import numpy as np

wheels = json.load(open('tools/wheels.json'))
POP = np.array([bin(i).count('1') for i in range(1 << 16)], dtype=np.int8)


def popcount(a):
    return POP[a & 0xFFFF] + POP[(a >> 16) & 0xFFFF]


def masks(nums, size):
    """every size-subset of range(nums) as a bitmask array"""
    out = np.zeros(sum(1 for _ in itertools.combinations(range(nums), size)), dtype=np.int64)
    for i, c in enumerate(itertools.combinations(range(nums), size)):
        out[i] = sum(1 << b for b in c)
    return out


results = {}
for name, w in wheels.items():
    n = w['nums']
    combos = [w['combos'][i] for i in sorted(w['combos'], key=int)]
    problems = []
    if len(combos) != w['games']:
        problems.append(f"{len(combos)} games in the table but the wheel says {w['games']}")
    games = []
    for c in combos:
        pos = [int(x) for x in re.findall(r'\d\d', c)]
        if len(set(pos)) != 6 or min(pos) < 1 or max(pos) > n:
            problems.append(f"game {c} uses positions outside a field of {n}")
        games.append(sum(1 << (p - 1) for p in set(pos)))
    if len(set(games)) != len(games):
        problems.append('duplicate games')
    if problems:
        results[name] = {'ok': False, 'problems': problems}
        print(f"{name:18} BROKEN: {'; '.join(problems)}")
        continue

    G = np.array(games, dtype=np.int64)
    worst, dist = {}, {}
    for m in range(3, 7):
        best = np.array([popcount(G & d).max() for d in masks(n, m)])
        worst[m] = int(best.min())
        if m == 6:
            dist = {int(k): round(float((best == k).mean()), 5) for k in range(3, 7) if (best == k).any()}

    # Division 2: six winning numbers + one supp all in the field -> a game with 5 + supp?
    div2 = True
    for sub in itertools.combinations(range(n), 7):
        for s in sub:
            main = sum(1 << b for b in sub if b != s)
            h = popcount(G & main)
            if not ((h == 6).any() or ((h == 5) & (((G >> s) & 1) == 1)).any()):
                div2 = False
                break
        if not div2:
            break

    results[name] = {'ok': True, 'field': n, 'games': len(G), 'worst': worst, 'dist6': dist, 'div2': div2}
    print(f"{name:18} field {n:2} games {len(G):4} worst-case match if 6/5/4/3 in field: "
          f"{worst[6]}/{worst[5]}/{worst[4]}/{worst[3]}  div2 {div2}  {dist}")

json.dump(results, open('tools/wheel_tests.json', 'w'), indent=1)
