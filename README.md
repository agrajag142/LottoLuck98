# LottoLuck98 (Lottoluck '98)

A Saturday Lotto (Tattslotto) checker, grapher, game designer and wheeling system, first written in Visual Basic 5 in 1997–98 and brought back here as source code and a web version.

## What's here

| Folder | Contents |
| --- | --- |
| `docs/` | `index.html`: the web version. It is a single page with everything built in and runs in any browser, with no server needed. |
| `vb6/` | The Visual Basic project, rebuilt from the original compiled program. Open `Xlotto.vbp` in VB6 or twinBASIC. |
| `data/` | `tattslot.txt` (the full Saturday Lotto draw history) and `number.lot` (the default game file). |
| `tools/` | `build.py`, `template.html` and `wheels.json`, used to regenerate `docs/index.html` after the data changes. |

## Web version

Open `docs/index.html` in a browser. To put it online, open the repo's Settings → Pages and choose to deploy from the `main` branch, `/docs` folder. It has the same five screens as the original:

- **Check winnings.** Checks your games against any number of past draws and shows division counts, expected wins and a list of wins.
- **Graphing.** Shows how often each ball has been drawn, weeks since each ball was last drawn, draw sums with a trend line, and your wins per draw.
- **Game designer.** Makes random games with a sum range, preferred numbers (coldest, hottest or most overdue) and the option to leave out 7 or 13.
- **Wheeling.** Holds 17 wheels, from 6 numbers in 1 game up to 28 numbers in 36 games.
- **Add draw.** Adds new results. They are saved in your browser, and the full `tattslot.txt` can be copied out.

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

## Data format

`tattslot.txt` has one draw per line:

```
Draw,YYYYMMDD,Drawn1,Drawn2,Drawn3,Drawn4,Drawn5,Drawn6,Drawn7,Drawn8
B4727,20261107,3,5,7,17,24,29,19,45
```

The first six numbers are the winning numbers and the last two are the supplementaries. Draws before B0413, when there was only one supplementary, have a 0 in the last column.

`number.lot` has one game per line, written as 12 digits (six two-digit numbers).
