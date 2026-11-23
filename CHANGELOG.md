# Changelog

Version numbers apply to the web version (`docs/index.html`). Each release is tagged in git as `vX.Y.Z`. The VB6 project in `vb6/` stays at the original Version 0.93h beta with the November 2026 bug fixes.

To release a new version, change the number and date in `tools/version.json`, add a section to the top of this file, and run `python3 tools/build.py`. The page reads both files, so the version and history shown in the app always match.

## 1.5.0 — 2026-11-24
- Check winnings shows approximate winnings using average payouts for Divisions 2–6, along with the cost, the net result and the return per dollar. Division 1 wins are counted but the jackpot is not included.
- Version numbering: the version is shown in the header and footer, and the footer has a version history.

## 1.4.0 — 2026-11-23
- Game designer: Power/due now compares recent draws with each ball's average since 45 balls. The old all-time baseline included the 40-ball draws, so 41–45 never came up as due.
- New Most overdue fill, a choice of counting window (last 20, 52 or 100 draws, or since 45 balls) and a Count supps option.
- Number pool, a never-use list, and even spreading of numbers across games.
- New filters: low/high, same last digit, numbers repeated from the last draw, the most numbers any two games may share, and a one-click typical sum range. Shows how many past draws would have passed the filters.
- Chart of how often each number is used in the game list.

## 1.3.0 — 2026-11-20
- Cost of play using Lott's Saturday Lotto pack prices (6–50 games) and their average per game, shown on Check winnings, Game designer and Wheeling. The price per game can be changed.

## 1.2.0 — 2026-11-20
- New 12-number, 132-game Division 2 wheel (Steiner system S(5,6,12)), tested exhaustively.

## 1.1.0 — 2026-11-20
- Every wheel tested against every possible draw. Each wheel now shows its true guarantee. Seven wheels claimed more than they deliver, and the broken TenOfFive wheel was removed.
- Division 6 (3 numbers) added. Outlay, a duplicate game warning and a search for draws containing chosen numbers.
- Graphs of balls drawn together and of odd numbers per draw.
- Game designer filters for odd numbers, runs and past Division 1 winners.
- Saving .lot files and tattslot.txt, and loading a newer tattslot.txt. Copied games use Windows line endings so the VB6 program can read them.

## 1.0.0 — 2026-11-11
- First web version of Lottoluck '98, with Check winnings, Graphing, Game designer, Wheeling and Add draw, and the full draw history to B4727.
