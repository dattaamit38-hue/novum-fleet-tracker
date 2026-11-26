# Novum fleet tracker

Interactive fleet map: vessel schedules from the position list, projected positions, live AIS (AISStream), weather overlays and a GIF export.

- Open `index.html` in a browser, or use the GitHub Pages link for this repository.
- To update for a new position list: put the workbook at `data/positions.xlsx`, update `SCHEDULE` in `build.py`, then run `python3 build.py` and copy `fleet_map.html` to `index.html`.

The position list workbook is not stored in this repository (`data/` is ignored).

## Live AIS feed

`.github/workflows/ais.yml` runs `ais_relay.py` every 10 minutes. It listens to AISStream.io for the MMSIs in
`ais_ships.json` and saves the latest positions to `ais.json` on the `ais-data` branch, which the page reads.
The AISStream key is stored as the repository secret `AISSTREAM_API_KEY` (Settings → Secrets and variables → Actions).
