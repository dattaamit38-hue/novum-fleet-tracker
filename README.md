# Novum fleet tracker

Interactive fleet map: vessel schedules from the position list, projected positions, live AIS (AISStream), weather overlays and a GIF export.

- Open `index.html` in a browser, or use the GitHub Pages link for this repository.
- To update for a new position list: put the workbook at `data/positions.xlsx`, update `SCHEDULE` in `build.py`, then run `python3 build.py` and copy `fleet_map.html` to `index.html`.

The position list workbook is not stored in this repository (`data/` is ignored).
