"""Collect AIS positions for the fleet from AISStream.io and merge them into ais.json.

Run by .github/workflows/ais.yml every 10 minutes. The web page reads ais.json from the
`ais-data` branch, so viewers need no API key and the key never reaches the browser.

Usage: AISSTREAM_API_KEY=... python ais_relay.py path/to/ais.json
"""
import asyncio, datetime as dt, json, os, re, sys, time
from pathlib import Path

import websockets

OUT = Path(sys.argv[1] if len(sys.argv) > 1 else "ais.json")
SHIPS = json.loads((Path(__file__).parent / "ais_ships.json").read_text())   # MMSI -> vessel name
LISTEN = int(os.environ.get("LISTEN_SECONDS", "240"))
KEEP_DAYS = 14
NAV = ["Under way using engine", "At anchor", "Not under command", "Restricted manoeuvrability",
       "Constrained by draught", "Moored", "Aground", "Engaged in fishing", "Under way sailing",
       "", "", "", "", "", "AIS-SART", "Not defined"]
MON = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()


def utc(s):
    """'2026-11-26 14:58:17.123456 +0000 UTC' -> ISO 8601 UTC"""
    m = re.match(r"(\d{4}-\d{2}-\d{2})[ T](\d{2}:\d{2}:\d{2})", s or "")
    if not m:
        return dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    return f"{m[1]}T{m[2]}Z"


async def listen(key, ships):
    seen = 0
    async with websockets.connect("wss://stream.aisstream.io/v0/stream", open_timeout=30, ping_interval=20) as ws:
        await ws.send(json.dumps({
            "APIKey": key,
            "BoundingBoxes": [[[-90, -180], [90, 180]]],
            "FiltersShipMMSI": list(SHIPS),
            "FilterMessageTypes": ["PositionReport", "StandardClassBPositionReport", "ShipStaticData"],
        }))
        end = time.time() + LISTEN
        while (left := end - time.time()) > 0:
            try:
                raw = await asyncio.wait_for(ws.recv(), timeout=left)
            except asyncio.TimeoutError:
                break
            m = json.loads(raw)
            if "error" in m:
                raise SystemExit(f"AISStream error: {m['error']}")
            md = m.get("MetaData", {})
            name = SHIPS.get(str(md.get("MMSI")))
            if not name:
                continue
            rec = ships.setdefault(name, {"mmsi": str(md.get("MMSI"))})
            msg = m.get("Message", {})
            pr = msg.get("PositionReport") or msg.get("StandardClassBPositionReport")
            if pr:
                t = utc(md.get("time_utc"))
                if t >= rec.get("time", ""):
                    hdg = pr.get("TrueHeading")
                    rec.update(lat=pr.get("Latitude", md.get("latitude")), lon=pr.get("Longitude", md.get("longitude")),
                               sog=pr.get("Sog"), cog=pr.get("Cog"), hdg=None if hdg == 511 else hdg,
                               nav=NAV[pr.get("NavigationalStatus", 15)] if pr.get("NavigationalStatus", 15) < 16 else "",
                               time=t)
                    seen += 1
            sd = msg.get("ShipStaticData")
            if sd:
                rec["dest"] = (sd.get("Destination") or "").strip()
                rec["draught"] = sd.get("MaximumStaticDraught")
                if sd.get("ImoNumber"):
                    rec["imo"] = sd["ImoNumber"]
                e = sd.get("Eta") or {}
                if e.get("Month"):
                    rec["eta"] = f"{e['Day']:02d} {MON[e['Month'] - 1]} {e['Hour']:02d}:{e['Minute']:02d} UTC"
    return seen


async def stream_check(key, seconds=15):
    """Count all AIS messages near Houston / Sabine Pass for a few seconds: proves the key and feed work."""
    n = 0
    async with websockets.connect("wss://stream.aisstream.io/v0/stream", open_timeout=30) as ws:
        await ws.send(json.dumps({"APIKey": key, "BoundingBoxes": [[[28.5, -95.6], [30.2, -93.4]]],
                                  "FilterMessageTypes": ["PositionReport"]}))
        end = time.time() + seconds
        while (left := end - time.time()) > 0:
            try:
                m = json.loads(await asyncio.wait_for(ws.recv(), timeout=left))
            except asyncio.TimeoutError:
                break
            if "error" in m:
                raise SystemExit(f"AISStream error: {m['error']}")
            n += 1
    return n


def main():
    key = os.environ.get("AISSTREAM_API_KEY", "").strip()
    try:
        data = json.loads(OUT.read_text())
    except Exception:
        data = {}
    ships = data.get("ships", {})
    now = dt.datetime.now(dt.timezone.utc)
    if not key:
        print("AISSTREAM_API_KEY secret is not set; nothing to do.")
        status = "no_key"
        seen = 0; check = None
    else:
        try:
            check = asyncio.run(stream_check(key))
            print(f"Feed check: {check} messages from all ships near Houston/Sabine in 15 s")
        except BaseException as e:
            check = f"failed: {e!r}"; print(f"Feed check {check}")
        try:
            seen = asyncio.run(listen(key, ships))
            status = "ok"
        except SystemExit as e:
            print(e); status = str(e); seen = 0
        except Exception as e:   # network trouble: keep the old data, report it
            print(f"Connection problem: {e!r}"); status = f"connection problem: {type(e).__name__}"; seen = 0
    cutoff = (now - dt.timedelta(days=KEEP_DAYS)).strftime("%Y-%m-%dT%H:%M:%SZ")
    ships = {k: v for k, v in ships.items() if v.get("time", "") >= cutoff and v.get("lat") is not None}
    out = {"updated": now.strftime("%Y-%m-%dT%H:%M:%SZ"), "status": status, "positions_this_run": seen,
           "listen_seconds": LISTEN, "feed_check_houston_15s": check, "ships": ships}
    OUT.write_text(json.dumps(out, indent=1))
    print(f"{status}: {seen} position reports this run; {len(ships)} ships on file")


if __name__ == "__main__":
    main()
