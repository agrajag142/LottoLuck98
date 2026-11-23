# LottoLuck98 (Lottoluck '98)

A Saturday Lotto (Tattslotto) checker, grapher, game designer and wheeling system, first written in Visual Basic 5 in 1997–98 and brought back here as source code and a web version.

## What's here

| Folder | Contents |
| --- | --- |
| `docs/` | `index.html`: the web version. It is a single page with everything built in and runs in any browser, with no server needed. |
| `vb6/` | The Visual Basic project, rebuilt from the original compiled program. Open `Xlotto.vbp` in VB6 or twinBASIC. |
| `data/` | `tattslot.txt` (the full Saturday Lotto draw history) and `number.lot` (the default game file). |
| `tools/` | `build.py`, `template.html` and `wheels.json`, used to regenerate `docs/index.html` after the data changes. `test_wheels.py` tests every wheel and writes `wheel_tests.json`. |

## Web version

The current web version number is shown in the page header, and **Version history** at the bottom of the page lists every release. The same history is in [CHANGELOG.md](CHANGELOG.md), and each release is tagged in git (`v1.0.0` … `v1.5.0`). To make a release, update `tools/version.json` and the top of `CHANGELOG.md`, then run `python3 tools/build.py`.

Open `docs/index.html` in a browser. To put it online, open the repo's Settings → Pages and choose to deploy from the `main` branch, `/docs` folder. It has the same five screens as the original:

- **Costs.** Uses Lott's Saturday Lotto prices: exact pack prices for 6, 10, 12, 18, 25, 36 and 50 games, and the average of those packs (about 95.7c a game) for any other number of games. You can change the price per game on the Check winnings screen. Costs are shown per draw and per year on the Game designer and Wheeling screens, and in the Wheel list.
- **Check winnings.** Checks your games against any number of past draws and shows counts for all six of today's divisions (Division 6 is 3 numbers), expected wins, the outlay at your price per game, and a list of wins. It warns about duplicate games, and it can find every draw that contained a set of numbers. Approximate winnings use average payouts for Divisions 2–6, and are shown with the cost, the net result and the return per dollar.
- **Graphing.** Shows how often each ball has been drawn, weeks since each ball was last drawn, draw sums with a trend line, your wins per draw, which balls are drawn most often with a chosen ball, and how many odd numbers each draw had.
- **Game designer.** Builds games from up to five key numbers that go in every game. These can be filled with the coldest, hottest, most overdue or power/due numbers, counted over the last 20, 52 or 100 draws or since 45 balls. You can also limit games to a number pool and give a list of numbers never to use. Numbers can be spread evenly across the games. Filters cover the sum range (one click sets it to the middle 80% of past sums), odd/even, low/high, consecutive runs, same last digit, numbers repeated from the last draw, how many numbers two games may share, and skipping past Division 1 winners. It shows how many past draws would have passed your filters and how often each number is used in your list.
- **Wheeling.** Holds 18 wheels, from 6 numbers in 1 game up to 28 numbers in 36 games, including a 12-number, 132-game Division 2 wheel (the Steiner system S(5,6,12)) added in 2026. Each one shows its tested guarantee and how often your chosen numbers have contained 3, 4, 5 or 6 of the winning numbers.
- **Add draw.** Adds new results, or loads a newer `tattslot.txt` and takes the draws it doesn't have. Draws are saved in your browser, and the full history can be copied or saved for the VB6 program.

To build the page again after updating `data/tattslot.txt`:

```
python3 tools/build.py
```

## VB6 project

The original source code was lost. This project was reconstructed from the compiled VB5 executable:

- **Recovered exactly:** the forms and their controls, menus, string constants and the wheel tables.
- **Rebuilt from the machine code:** the procedure bodies. Lines marked `RECON:` are places where a detail couldn't be read back exactly and an equivalent was written instead.

The draw arrays have been enlarged from 1,500 draws to 5,000 draws so that the current history fits.

The project needs these components, which all ship with VB6: COMCTL32.OCX, COMDLG32.OCX, SYSINFO.OCX, MCI32.OCX, MSINET.OCX and MSCHRT20.OCX.

The win sounds (`sound1.wav`, `sound2.wav` and `sound3.wav`) are not included. Put any WAV files with those names next to the program to hear them, or turn sound off in the menu.

## Wheel test results

`tools/test_wheels.py` tries every possible draw inside each wheel's field and finds the worst case. All 17 working wheels produce valid games. However, seven of them guaranteed less than the 1998 program said:

| Wheel | Games | 1998 text said | Tested guarantee (all 6 in field) |
| --- | --- | --- | --- |
| 6–12 numbers, full wheels | 1–924 | Division 1 | Division 1 ✓ |
| 9 numbers | 12 | Division 3 | Division 3 ✓ (Division 2 if a supp is in the field too) |
| 10 numbers | 10 | Division 4 | Division 4 ✓ |
| 11 numbers | 5 | Division 4 | Division 4 ✓ |
| 12 numbers | 6 | Division 3 | **Division 4** |
| 12 numbers | 18 | Division 3 | **Division 4** |
| 12 numbers | 42 | Division 2 | **Division 3** |
| 12 numbers (new, web only) | 132 | — | Division 2 if a supp is in the field too, otherwise Division 3. Also Division 3 when only 5 of the winning numbers are in the field. |
| 14 numbers | 14 | Division 3 | **Division 4** |
| 17 numbers | 11 | Division 4 | **3 numbers (Division 6)** |
| 18 numbers | 42 | Division 3 | **Division 4** |
| 28 numbers | 36 | Division 3 | **3 numbers (Division 6)** |

The `TenOfFive` wheel is broken in the original program. Its table has only 5 of its 10 games, and it uses positions up to 18 while asking for 5 numbers, so it has been disabled. The wheel descriptions in both versions now state the tested guarantee.

## Fixes made to the VB6 code

- **Newest draw ignored.** The history loader stopped one draw short, so the newest draw was never checked or graphed. Add Draw then offered the next draw number after the second-newest draw, which already existed.
- **Wheeled games broken.** Numbers chosen on the wheeling grid weren't padded to two digits, so any number under 10 broke the 12-digit games.
- **Picking stuck.** After changing wheels, or using Quick Pick on a full field, the selected count wasn't reset, so no numbers could be picked.
- **Overflow crashes.** Big checks (lots of games over all draws) overflowed the Integer counters (error 6), and game files over 1,250 games crashed.
- **Startup error.** Removed `DrawLine(GameIndex) = File`, which assigned the File menu to a string (error 438 in VB6).
- **Game designer hangs.** It could hang forever when a preferred number was also excluded (7 or 13) or when the sum range couldn't be met. It also accepted non-numbers in the preferred boxes and invalid hand-typed games.
- **Add Draw.** Next week's date didn't roll into the next month (day 28 + 7 gave day 35). Numbers are now checked before they're written to `tattslot.txt`.
- **Random numbers.** `Randomize` is called once at start. The old "stir" loop also overwrote a variable used by the wheeling screen.
- **Smaller fixes:**
  - The outlay shown on the 4, 8 and 52 draw buttons was worked out for 5, 10 and 50 draws.
  - The sum graph is now drawn oldest to newest, so its trend line reads the right way.
  - The win sounds are optional.
  - The dead Tattersalls download now explains where to get new draws.

The VB6 changes are marked `FIX:` in the code.

## Data format

`tattslot.txt` has one draw per line:

```
Draw,YYYYMMDD,Drawn1,Drawn2,Drawn3,Drawn4,Drawn5,Drawn6,Drawn7,Drawn8
B4727,20261107,3,5,7,17,24,29,19,45
```

The first six numbers are the winning numbers and the last two are the supplementaries. Draws before B0413, when there was only one supplementary, have a 0 in the last column.

`number.lot` has one game per line, written as 12 digits (six two-digit numbers).
