"""Build the fleet schedule map (fleet_map.html) from the position list workbook.

Usage:  python3 build.py [path/to/positions.xlsx]

What comes from where:
  * Vessel specs, remarks, SIRE dates and charter dates are read straight from the workbook.
  * The port-call schedule is free text in the "Remarks" column, so it is transcribed by hand
    in SCHEDULE below (each entry keeps the remark text it came from). Update SCHEDULE when a
    new position list arrives, then re-run this script.
  * Sea routes are approximate waypoint paths (ROUTES) so lines follow open water instead of
    crossing land. They are for display, not navigation.
"""
import json, math, re, sys, datetime as dt
from pathlib import Path
import openpyxl

HERE = Path(__file__).parent
XLSX = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / "data" / "positions.xlsx"
SPEED_KN = 12.0  # assumed speed when a departure/arrival time is missing

# ----------------------------------------------------------------------------- ports
# lat, lon. Dongying is east of the antimeridian; it is stored "unwrapped" (lon - 360)
# so the Pacific crossing draws as one continuous line.
PORTS = {
    "Guayanilla":    ("Guayanilla, PR", 17.98, -66.77),
    "Rotterdam":     ("Rotterdam, NL", 51.95, 4.05),
    "NYH":           ("New York Harbor", 40.66, -74.06),
    "Itacoatiara":   ("Itacoatiara, BR", -3.145, -58.43),
    "Manaus":        ("Manaus, BR", -3.14, -60.02),
    "BarraNorte":    ("Barra Norte (Amazon bar)", 1.55, -49.75),
    "Bayovar":       ("Bayovar, PE", -5.82, -81.06),
    "Dongying":      ("Dongying, CN", 38.07, 118.98 - 360),
    "PortArthur":    ("Port Arthur, TX", 29.87, -93.93),
    "Nederland":     ("Nederland, TX", 29.97, -93.99),
    "StCharles":     ("St. Charles, LA", 29.96, -90.36),
    "Palenque":      ("Palenque, DR", 18.22, -70.17),
    "StCroix":       ("St. Croix, USVI", 17.69, -64.74),
    "Barcelona":     ("Barcelona, ES", 41.34, 2.16),
    "TTSTS":         ("T&T offshore (STS)", 10.55, -61.72),
    "PortOfSpain":   ("Port of Spain anchorage, TT", 10.65, -61.52),
    "Macapa":        ("Macapá, BR", -0.05, -51.10),
    "Valencia":      ("Valencia, ES", 39.44, -0.31),
    "Georgetown":    ("Georgetown, GY", 6.81, -58.16),
    "VreedEnHoop":   ("Vreed en Hoop, GY", 6.81, -58.19),
}

# ----------------------------------------------------------------------------- route pieces
AMAZON = [  # Manaus -> Barra Norte
    (-3.14, -60.02), (-3.2, -59.3), (-3.145, -58.43), (-2.95, -57.55), (-2.63, -56.74),
    (-2.2, -56.1), (-1.92, -55.52), (-2.35, -54.85), (-2.15, -54.2), (-1.85, -53.3),
    (-1.55, -52.6), (-1.35, -51.9), (-1.1, -51.45), (-0.5, -51.25), (-0.05, -51.1),
    (0.5, -50.45), (1.0, -50.0), (1.55, -49.75)]
AMAZON_ITA = AMAZON[2:]            # Itacoatiara -> Barra Norte
GUIANA = [(1.55, -49.75), (3.5, -49.3), (5.5, -51.0), (7.5, -54.5), (9.0, -57.5), (10.9, -59.6)]
GATE = [(10.9, -59.6), (11.6, -60.3), (11.85, -61.4), (11.9, -62.2)]   # Atlantic -> Caribbean S of Grenada
CARIB_TO_GUAY = [(11.9, -62.2), (14.5, -64.0), (16.8, -65.8), (17.98, -66.77)]
CARIB_TO_YUC = [(11.9, -62.2), (12.6, -66.0), (14.5, -73.0), (17.0, -79.5), (20.0, -84.5),
                (21.8, -85.6), (24.0, -86.5)]
GULF_PA = [(29.87, -93.93), (29.68, -93.85), (28.8, -93.4), (26.5, -90.0), (24.0, -86.5)]
MISS = [(28.88, -89.45), (29.15, -89.25), (29.37, -89.53), (29.55, -89.75), (29.8, -89.95),
        (29.95, -90.06), (29.93, -90.2), (29.96, -90.36)]                # SW Pass -> St. Charles
NED_TO_SWP = [(29.97, -93.99), (29.88, -93.93), (29.68, -93.85), (29.0, -92.5), (28.6, -90.5),
              (28.75, -89.6), (28.88, -89.45)]
SWP_TO_PAL = [(28.88, -89.45), (28.6, -89.3), (24.0, -86.5), (21.8, -85.6), (20.0, -84.5),
              (19.0, -82.0), (18.6, -80.0), (17.3, -77.0), (17.4, -74.5), (17.2, -71.6),
              (17.9, -70.5), (18.22, -70.17)]
GUAY_NYH = [(17.98, -66.77), (17.85, -67.3), (18.25, -67.55), (18.8, -67.7), (20.0, -68.3),
            (25.0, -70.5), (33.0, -73.0), (38.5, -73.5), (40.42, -73.85), (40.55, -74.03),
            (40.66, -74.06)]
PACIFIC = [(-5.82, -81.06), (-5.9, -81.5), (-3.0, -90.0), (5.0, -120.0), (18.0, -160.0),
           (27.0, -190.0), (28.0, -215.0), (29.0, -221.0), (30.4, -227.0), (30.92, -228.7),
           (30.92, -229.7), (31.0, -231.0), (33.0, -235.0), (36.5, -236.7), (37.6, -237.0),
           (38.3, -238.4), (38.3, -239.1), (38.1, -240.4), (38.07, 118.98 - 360)]
STX_BCN = [(17.69, -64.74), (18.3, -63.8), (20.0, -61.0), (28.0, -45.0), (34.0, -25.0),
           (36.0, -10.0), (35.95, -5.6), (36.2, -4.0), (36.6, -1.6), (38.3, 0.6),
           (38.85, 0.85), (40.3, 1.4), (41.34, 2.16)]
STX_STS = [(17.69, -64.74), (16.0, -63.6), (13.5, -62.3), (12.0, -62.05), (11.0, -61.9),
           (10.70, -61.82), (10.62, -61.72), (10.55, -61.72)]
STS_GEO = [(10.55, -61.72), (10.62, -61.72), (10.70, -61.82), (10.85, -61.6), (10.9, -61.0),
           (10.98, -60.75), (10.5, -59.8), (8.5, -58.6), (7.2, -58.2), (6.85, -58.17), (6.81, -58.16)]
GEO_POS = list(reversed(STS_GEO))[:-3] + [(10.70, -61.82), (10.62, -61.72), (10.65, -61.52)]
GEO_POS[0] = (6.81, -58.19)   # Monarca 1 sails from Vreed en Hoop
POS_STS = [(10.65, -61.52), (10.6, -61.62), (10.55, -61.72)]


def rev(p): return list(reversed(p))
def join(*parts):
    out = []
    for p in parts:
        out.extend(p if not out else p[1:] if out[-1] == p[0] else p)
    return out


# Rotterdam -> Guayanilla, following Asopos's Signal AIS track across Biscay and the Atlantic
ROTT_GUAY = [(51.95, 4.05), (51.98, 3.6), (51.6, 2.3), (51.0, 1.45), (50.3, -0.5), (49.9, -3.0), (49.2, -5.5),
             (48.0, -7.9), (45.35, -13.1), (42.58, -17.42), (39.76, -21.93), (37.96, -24.55), (31.0, -35.0),
             (23.0, -52.0), (19.2, -62.0), (18.25, -63.9), (17.98, -64.9), (17.8, -66.0), (17.98, -66.77)]

MACAPA_IN = [(1.55, -49.75), (1.0, -50.0), (0.5, -50.45), (-0.05, -51.1)]     # Barra Norte -> Macapá
_ita_up = rev(AMAZON_ITA)                                                      # Barra Norte -> Itacoatiara
MACAPA_ITA = _ita_up[_ita_up.index((-0.05, -51.1)):]
STX_VLC = STX_BCN[:STX_BCN.index((36.6, -1.6)) + 1] + [(37.5, -0.3), (38.5, 0.4), (39.2, 0.0), (39.44, -0.31)]
VLC_BCN = [(39.44, -0.31), (39.6, 0.3), (40.3, 1.2), (41.0, 1.9), (41.34, 2.16)]

ROUTES = {
    ("PortArthur", "Macapa"): join(GULF_PA, rev(CARIB_TO_YUC), rev(GATE), rev(GUIANA), MACAPA_IN),
    ("Macapa", "Itacoatiara"): MACAPA_ITA,
    ("StCroix", "Valencia"): STX_VLC,
    ("Valencia", "Barcelona"): VLC_BCN,
    ("PortOfSpain", "TTSTS"): [(10.65, -61.52), (10.6, -61.62), (10.55, -61.72)],
    ("Rotterdam", "Guayanilla"): ROTT_GUAY,
    ("Guayanilla", "NYH"): GUAY_NYH,
    ("NYH", "Guayanilla"): rev(GUAY_NYH),
    ("BarraNorte", "Itacoatiara"): rev(AMAZON_ITA),
    ("Itacoatiara", "Guayanilla"): join(AMAZON_ITA, GUIANA, GATE, CARIB_TO_GUAY),
    ("Bayovar", "Dongying"): PACIFIC,
    ("PortArthur", "BarraNorte"): join(GULF_PA, rev(CARIB_TO_YUC), rev(GATE), rev(GUIANA)),
    ("BarraNorte", "Itacoatiara"): rev(AMAZON_ITA),
    ("Itacoatiara", "Manaus"): rev(AMAZON[:3]),
    ("Manaus", "Itacoatiara"): AMAZON[:3],
    ("Nederland", "StCharles"): join(NED_TO_SWP, MISS),
    ("StCharles", "Palenque"): join(rev(MISS), SWP_TO_PAL),
    ("StCroix", "Barcelona"): STX_BCN,
    ("StCroix", "TTSTS"): STX_STS,
    ("TTSTS", "Georgetown"): STS_GEO,
    ("VreedEnHoop", "PortOfSpain"): GEO_POS,
    ("PortOfSpain", "TTSTS"): POS_STS,
}

# ----------------------------------------------------------------------------- schedule
# Times are local time as written in the report (report date 25 Nov 2026).
# AM=08:00, PM=16:00, LPM=18:00, EAM=04:00, Noon=12:00 when no hour is given.
# etb = estimated time of berthing where the remarks give one (etb=None: ETB TBC).
# "est": True marks a time not in the report (worked out at SPEED_KN or from the open date).
REPORT_DATE = "2026-11-27"

SCHEDULE = {
    # From the position list of 27 Nov 2026. alongside=True: the list says the ship is at berth now, so an
    # older stopped-off-port AIS position does not hold it at the anchorage.
    "ASOPOS": dict(
        # Rotterdam sailing from the voyage screen the user shared; positions come from the Signal export.
        status="At sea", origin="Rotterdam", origin_dep="2026-11-21T11:51",
        stops=[
            dict(port="Guayanilla", arr="2026-12-05T14:00", dep="2026-12-07T12:00", act="Discharge", etb="2026-12-05T18:00",
                 voy="26DPP141", tbc=True, note="ETA 05/1400LT; ETB 05/1800 (TBC); open 7 Dec (TBC)"),
            dict(port="NYH", arr="2026-12-12T18:00", dep=None, act="Load", voy="Next voyage",
                 tbc=True, note="Jones Act waiver obtained; loading NYH 12/LPM (TBC)"),
            dict(port="Guayanilla", arr="2026-12-20T08:00", dep=None, act="Discharge", voy="Next voyage",
                 tbc=False, note="ETA Guayanilla 20 AM"),
        ]),
    "VALLE BIANCA": dict(
        status="At anchorage (loading)", origin=None,
        stops=[
            dict(port="Itacoatiara", arr="2026-11-26T21:00", dep="2026-12-02T12:00", act="Load", alongside=True,
                 voy="26DPP132", tbc=False, note="At Ita anchorage loading (arrived about 26 Nov 21:00 per AIS); ETS 02 Dec"),
            dict(port="Guayanilla", arr="2026-12-10T12:00", dep="2026-12-12T12:00", act="Discharge",
                 voy="26DPP132", tbc=True, note="ETA 10 Dec (TBC); open 12 Dec (TBC)"),
        ]),
    "VALVERDE": dict(
        status="At sea", origin=None,
        stops=[
            dict(port="Bayovar", arr=None, dep="2026-11-27T03:00", act="Load", voy="26DPP142",
                 tbc=False, note="Sailed Bayovar 27/0300LT"),
            dict(port="Dongying", arr="2026-12-27T07:00", dep="2026-12-29T12:00", act="Discharge",
                 voy="26DPP142", tbc=True, note="ETA 27 Dec 0700LT; open 29 Dec (TBC)"),
        ]),
    "XANTHOS": dict(
        status="At sea", origin="PortArthur", origin_dep="2026-11-19T18:00",   # sailing per the Signal AIS track
        stops=[
            dict(port="Macapa", arr="2026-12-02T15:00", dep="2026-12-02T15:00", act="Transit",
                 voy="26DPP129", tbc=False, note="ETA Macapá 02/1500LT"),
            dict(port="Itacoatiara", arr="2026-12-04T08:00", dep="2026-12-05T16:00", act="Discharge",
                 voy="26DPP129", tbc=False, note="ETA Ita 04 AM; ETC Ita 05 PM; then shifting to Manaus"),
            dict(port="Manaus", arr=None, dep="2026-12-09T16:00", act="Discharge", etb=None,
                 voy="26DPP129", tbc=False, note="ETB Manaus 05 PM (list); ETC 08 PM; ETS 09 PM"),
            dict(port="Itacoatiara", arr="2026-12-10T08:00", dep="2026-12-16T08:00", act="Load",
                 voy="26DPP150", tbc=False, note="Wash and backload Bretana; ETA 10 AM; ETS 16 AM"),
            dict(port="Guayanilla", arr="2026-12-23T08:00", dep="2026-12-25T12:00", act="Discharge",
                 voy="26DPP150", tbc=True, note="ETA Guayanilla (TBC) 23 AM; open 25 Dec (TBC)"),
        ]),
    "VICTORY": dict(
        status="At berth (loading)", origin=None,
        stops=[
            dict(port="Nederland", arr="2026-11-26T11:00", dep="2026-11-28T11:00", act="Load", etb=None, alongside=True,
                 voy="", tbc=False, note="At berth, loading; ETS earliest 28/1100LT"),
            dict(port="StCharles", arr="2026-11-29T20:00", dep="2026-12-01T08:00", act="Port call",
                 voy="", tbc=False, note="ETA 29/2000LT; ETS 01 AM"),
            dict(port="Palenque", arr="2026-12-06T12:00", dep="2026-12-08T16:00", act="Discharge",
                 voy="", tbc=True, note="ETA 06 Dec noon (TBC); open 8 Dec PM"),
        ]),
    "HIGH CHALLENGE": dict(
        status="At sea", origin=None,
        stops=[
            dict(port="StCroix", arr="2026-11-25T16:00", dep="2026-11-26T14:00", act="Port call", etb="2026-11-25T20:00",
                 voy="26DPP151", tbc=False, note="Departed St Croix 26/1400LT; max speed to Valencia"),
            dict(port="Valencia", arr="2026-12-08T21:00", dep="2026-12-10T08:00", act="Discharge", etb=None,
                 voy="26DPP151", tbc=True, note="Part discharge; ETA 08/2100LT; ETB 08 LPM (TBC); ETS 10 AM (TBC)"),
            dict(port="Barcelona", arr="2026-12-10T18:00", dep="2026-12-12T08:00", act="Discharge",
                 voy="26DPP151", tbc=True, note="ETA 10 LPM (TBC); open 12 Dec AM (TBC); next voyage: load Barcelona, discharge Skikda"),
        ]),
    "ATA VOYAGER": dict(
        status="At berth", origin=None,
        stops=[
            dict(port="StCroix", arr="2026-11-26T21:24", dep="2026-11-28T16:00", act="Port call", etb="2026-11-26T21:24", alongside=True,
                 voy="", tbc=True, note="Berthed 26/2124LT; ETS 28 PM (TBC)"),
            dict(port="TTSTS", arr="2026-12-01T08:00", dep="2026-12-02T08:00", act="STS",
                 voy="", tbc=False, note="ETA T&T offshore for STS 01 AM; ETS 02 AM"),
            dict(port="Georgetown", arr="2026-12-03T16:00", dep="2026-12-05T12:00", act="Discharge",
                 voy="", tbc=True, note="ETA Georgetown 03 PM; open 5 Dec (TBC)"),
        ]),
    "MONARCA 1": dict(
        status="At anchorage", origin=None,
        stops=[
            dict(port="PortOfSpain", arr=None, dep=None, act="Port call", voy="",
                 tbc=False, note="At Port of Spain anchorage"),
            dict(port="TTSTS", arr="2026-12-01T08:00", dep="2026-12-03T16:00", act="STS", voy="",
                 tbc=False, note="STS planned 01 AM; ETS T&T 03 PM"),
            dict(port="Georgetown", arr="2026-12-05T08:00", dep="2026-12-07T12:00", act="Discharge",
                 voy="", tbc=True, note="ETA Georgetown 05 AM (TBC); open 7 Dec (TBC)"),
        ]),
    "PRIMA N": dict(
        status="Delivering", origin=None, region=("East Atlantic", 24.0, -28.0, 900),
        window=dict(label="Delivery to Novum", start="2026-12-21T00:00", end="2026-12-23T23:59"),
        stops=[]),
    "ALMA MARINE": dict(
        status="Held in port", origin=None, region=("West Africa", 4.0, 3.0, 600),
        window=dict(label="Open St. Croix", start="2026-12-15T00:00", end="2026-12-30T23:59"),
        stops=[]),
}

# Things in the workbook that look inconsistent; shown under "Data notes" in the page.
NOTES = [
    dict(vessel="XANTHOS", text="Xanthos: the list gives ETC Itacoatiara 05 PM and ETB Manaus 05 PM, which can't both hold (about 110 nm apart). The page keeps ETC Ita and works out the Manaus arrival."),
    dict(vessel="HIGH CHALLENGE", text="High Challenge: ETB Valencia (08 LPM) is before its ETA (08/2100LT), so the ETB is treated as TBC. Open position 'Balcelona' is read as Barcelona, Spain."),
    dict(vessel="VICTORY", text="Victory: 'Palanque' is read as Palenque, Dominican Republic."),
    dict(vessel="PRIMA N", text="Prima N: no particulars (year, DWT, size) in the position list."),
    dict(vessel="MONARCA 1", text="Monarca 1 is in the position list but not on the charter sheet."),
]

# ----------------------------------------------------------------------------- helpers
def nm(a, b):
    la1, lo1, la2, lo2 = map(math.radians, (a[0], a[1], b[0], b[1]))
    h = math.sin((la2 - la1) / 2) ** 2 + math.cos(la1) * math.cos(la2) * math.sin((lo2 - lo1) / 2) ** 2
    return 2 * 3440.065 * math.asin(math.sqrt(h))

def path_nm(p): return sum(nm(p[i], p[i + 1]) for i in range(len(p) - 1))
def t(s): return dt.datetime.fromisoformat(s)
def iso(d): return d.strftime("%Y-%m-%dT%H:%M")

def route(a, b):
    if (a, b) in ROUTES:
        return [list(x) for x in ROUTES[(a, b)]]
    pa, pb = PORTS[a], PORTS[b]
    print(f"  ! no waypoint route for {a} -> {b}; drawing a straight line")
    return [[pa[1], pa[2]], [pb[1], pb[2]]]

def norm(name): return re.sub(r"[^A-Z0-9]", "", str(name or "").upper())

def num(v):
    return v if isinstance(v, (int, float)) else None

# ----------------------------------------------------------------------------- read workbook
def read_workbook():
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb["Position List"]
    report_date = ws["A1"].value
    vessels, segment = [], None
    for row in ws.iter_rows(min_row=1, values_only=True):
        name = row[0]
        if not isinstance(name, str) or not name.strip():
            continue
        if name.strip() == "Vessel Name" or name.startswith("Novum"):
            continue
        if all(v is None or str(v).strip() == "" for v in row[1:12]):
            segment = name.strip().replace(" Vessels", "")
            continue
        vessels.append(dict(
            name=name.strip(), key=norm(name), segment=segment,
            year=num(row[1]), dwt=num(row[2]), cubics=num(row[3]), draft=num(row[4]),
            beam=num(row[5]), loa=num(row[6]), cargoes=row[7], open_pos=row[8],
            open_date=str(row[9]).strip() if row[9] else None, current=row[10],
            remarks=re.sub(r"\s{2,}", "  ", str(row[11] or "")).strip()))
    sire = {}
    for row in wb["SIRE dates"].iter_rows(min_row=5, values_only=True):
        if isinstance(row[0], str) and row[0] != "Vessel Name" and not row[0].endswith("Vessels"):
            sire[norm(row[0])] = dict(
                last=row[1].date().isoformat() if isinstance(row[1], dt.datetime) else None,
                place=row[2],
                due=row[3].date().isoformat() if isinstance(row[3], dt.datetime) else None,
                remark=(row[6] or "").strip() or None)
    charter = {}
    if "Sheet2" in wb.sheetnames:
        for row in wb["Sheet2"].iter_rows(min_row=2, values_only=True):
            if row[1]:
                charter[norm(row[1])] = dict(delivery=row[2], earliest=row[3], latest=row[4])
    return report_date, vessels, sire, charter


def read_signal_positions():
    """Daily AIS positions exported from Signal (data/signal_positions_*.xlsx; newest file wins).
    Each row is one vessel-day: average lat/lon, speed, heading, loading state, AIS destination and ETA."""
    files = sorted((HERE / "data").glob("signal_positions_*.xlsx"))
    if not files:
        return None, {}
    wb = openpyxl.load_workbook(files[-1], data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    head = [str(h).strip() if h else "" for h in rows[0]]
    col = {h: i for i, h in enumerate(head)}
    need = ["DayDate", "VesselName", "AverageLatitude", "AverageLongitude"]
    if any(n not in col for n in need):
        print(f"  ! {files[-1].name}: expected columns {need}")
        return files[-1].name, {}
    out = {}
    for r in rows[1:]:
        if not r or not r[col["VesselName"]] or r[col["AverageLatitude"]] is None:
            continue
        g = lambda k: r[col[k]] if k in col else None
        day = g("DayDate"); eta = g("EtaToDestination")
        out.setdefault(norm(g("VesselName")), []).append(dict(
            day=day.date().isoformat() if isinstance(day, dt.datetime) else str(day)[:10],
            lat=round(float(g("AverageLatitude")), 4), lon=round(float(g("AverageLongitude")), 4),
            sog=g("Speed"), heading=g("HeadingDirection"), state=g("LoadingState"), move=g("MovementState"),
            dest=g("AisDestinationPortName"), dest_country=g("AisDestinationCountryName"),
            eta=eta.strftime("%Y-%m-%dT%H:%M") if isinstance(eta, dt.datetime) else None,
            draught=g("AverageDraught"), cargo=g("CargoType") or None, days_since_ais=g("DaysSinceLastAis"),
            voyage=g("VoyageNumber")))
    for k in out:
        out[k].sort(key=lambda x: x["day"])
    print(f"Signal positions: {files[-1].name}, {sum(len(v) for v in out.values())} vessel-days for {len(out)} vessels")
    return files[-1].name, out


def parse_charter_date(s):
    """'Nov 10 '26/ 1400LT; 1900Z' -> '2026-11-10'"""
    if not isinstance(s, str):
        return None
    m = re.match(r"([A-Za-z]+)\s+(\d{1,2})\s+'(\d{2})", s.strip())
    if not m:
        return None
    try:
        return dt.datetime.strptime(f"{m[1][:3]} {m[2]} 20{m[3]}", "%b %d %Y").date().isoformat()
    except ValueError:
        return None


# ----------------------------------------------------------------------------- build
def build():
    report_date, vessels, sire, charter = read_workbook()
    sig_file, sig = read_signal_positions()
    rd = report_date.date() if isinstance(report_date, dt.datetime) else dt.date.fromisoformat(REPORT_DATE)
    print(f"Report date {rd}; {len(vessels)} vessels in workbook")
    out = []
    for v in vessels:
        sched = SCHEDULE.get(v["key"].replace("1", " 1").strip()) or next(
            (s for k, s in SCHEDULE.items() if norm(k) == v["key"]), None)
        if sched is None:
            print(f"  ! {v['name']}: no schedule entry, shown in the list only")
            sched = dict(status=v["current"] or "Unknown", origin=None, stops=[])
        stops = [dict(s) for s in sched["stops"]]
        for s in stops:
            s["name"], s["lat"], s["lon"] = PORTS[s["port"]]
            s["est_dep"] = s["est_arr"] = False
        legs = []
        for i in range(len(stops) - 1):
            a, b = stops[i], stops[i + 1]
            p = route(a["port"], b["port"])
            d = path_nm(p)
            if a["dep"] is None and b["arr"]:
                a["dep"], a["est_dep"] = iso(t(b["arr"]) - dt.timedelta(hours=d / SPEED_KN)), True
            if b["arr"] is None and a["dep"]:
                b["arr"], b["est_arr"] = iso(t(a["dep"]) + dt.timedelta(hours=d / SPEED_KN)), True
            legs.append(dict(path=p, nm=round(d)))
        origin = None
        if sched.get("origin") and stops:
            p = route(sched["origin"], stops[0]["port"])
            d = path_nm(p)
            origin = dict(port=sched["origin"], name=PORTS[sched["origin"]][0],
                          dep=sched.get("origin_dep") or iso(t(stops[0]["arr"]) - dt.timedelta(hours=d / SPEED_KN)),
                          est=not sched.get("origin_dep"), path=p, nm=round(d))
        s = sire.get(v["key"], {})
        c = charter.get(v["key"], {})
        out.append(dict(
            **{k: v[k] for k in ("name", "segment", "year", "dwt", "cubics", "draft", "beam", "loa",
                                 "cargoes", "open_pos", "open_date", "current", "remarks")},
            status=sched["status"], stops=stops, legs=legs, origin=origin,
            region=dict(zip(("label", "lat", "lon", "km"), sched["region"])) if sched.get("region") else None,
            window=sched.get("window"),
            signal=sig.get(v["key"], []),
            charter=dict(delivery=c.get("delivery"), earliest=parse_charter_date(c.get("earliest")),
                         latest=parse_charter_date(c.get("latest")),
                         earliest_raw=c.get("earliest"), latest_raw=c.get("latest")) if c else None))
    data = dict(report_date=rd.isoformat(), speed_kn=SPEED_KN, source=XLSX.name, notes=NOTES, signal_file=sig_file,
                generated=dt.datetime.now().strftime("%Y-%m-%d %H:%M"), vessels=out)
    html = (HERE / "template.html").read_text()
    html = html.replace("/*__DATA__*/null", json.dumps(data, separators=(",", ":"), default=str))
    (HERE / "fleet_map.html").write_text(html)
    print(f"Wrote {HERE / 'fleet_map.html'} ({len(html) // 1024} KB)")


if __name__ == "__main__":
    build()
