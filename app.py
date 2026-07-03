import streamlit as st
import pandas as pd
import glob
import os
import json
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Evergreen Lake Michelle — Meter Hierarchy",
    page_icon="⚡",
    layout="wide",
)

# ---------- Styling (same palette as the Sitari app) ----------
st.markdown("""
<style>
.block-container {padding-top: 1.2rem; max-width: 1400px;}
.stMetric, div[data-testid="stMetric"] {
  background: #FBF9F3 !important;
  border: 1px solid #DCD6C4;
  border-radius: 10px;
  padding: 10px 14px;
}
div[data-testid="stMetric"] * {
  color: #152B45 !important;
}
div[data-testid="stMetricLabel"], div[data-testid="stMetricLabel"] * {
  font-size: 12px !important;
  text-transform: uppercase;
  letter-spacing: .05em;
  color: #3E5066 !important;
}
div[data-testid="stMetricValue"], div[data-testid="stMetricValue"] * {
  color: #152B45 !important;
}
div[data-testid="stMetricDelta"], div[data-testid="stMetricDelta"] * {
  color: #3F7D5C !important;
  fill: #3F7D5C !important;
}
h1, h2, h3 {color: #152B45;}
</style>
""", unsafe_allow_html=True)

SITE_NAME = "Evergreen Lake Michelle"
ELEC_SHEET_CANDIDATES = ["Elec", "Elec Meters", "Electrical Meters"]
FILE_PATTERN = "EVG_Lake_Michelle_Meter_Hierarchy_*.xlsx"

# =====================================================================
# SITE RETICULATION CONFIG — transcribed from the as-built SLDs
# MS → feeder → daisy-chained kiosks. "cable_in" is the cable feeding
# that kiosk (from the MS for the first kiosk, else from the previous
# kiosk in the chain). "extras" are non-metered items in the kiosk.
# Meters themselves still come from the spreadsheet, grouped by kiosk.
# =====================================================================
RETICULATION = [
    {
        "ms_name": "MS LM-C",
        "serial": "98820081",
        "rating": "315 kVA (500 A)",
        "ct": "CT Ratio 500:5",
        "note": "As built 04.2021",
        "feeders": [
            {
                "name": "Feeder 1",
                "chain": [
                    {"kiosk": "LM1C", "cable_in": "250A · 185×4mm² · 48m"},
                    {"kiosk": "LM2C", "cable_in": "120mm² · 39m"},
                    {"kiosk": "LM3C", "cable_in": "120mm² · 51m",
                     "extras": [{"name": "SL Switchgear", "detail": "3x60A · RWB · no meter"}]},
                ],
            },
            {
                "name": "Feeder 2",
                "chain": [
                    {"kiosk": "LM4C", "cable_in": "250A · 185×4mm² · 145m"},
                    {"kiosk": "LM5C", "cable_in": "120mm² · 57m"},
                ],
            },
            {
                "name": "Feeder 3",
                "chain": [
                    {"kiosk": "LM6C", "cable_in": "250A · 300×4mm² · 191m"},
                    {"kiosk": "LM7C", "cable_in": "300mm² · 56m"},
                ],
            },
        ],
    },
    {
        "ms_name": "MS LM-D",
        "serial": "98820079",
        "rating": "315 kVA (500 A)",
        "ct": "CT Ratio 500:5",
        "note": "",
        "feeders": [
            {
                "name": "Feeder 1",
                "chain": [
                    {"kiosk": "LM1D", "cable_in": "250A · 185×4mm² · 239m"},
                    {"kiosk": "LM2D", "cable_in": "185mm² · 64m"},
                ],
            },
            {
                "name": "Feeder 2",
                "chain": [
                    {"kiosk": "LM3D", "cable_in": "250A · 185×4mm² · 148m"},
                    {"kiosk": "LM4D", "cable_in": "185mm² · 31m"},
                    {"kiosk": "LM5D", "cable_in": "185mm² · 65m"},
                ],
            },
            {
                "name": "Feeder 3",
                "chain": [
                    {"kiosk": "LM6D", "cable_in": "250A · 185×4mm² · 75m"},
                    {"kiosk": "LM7D", "cable_in": "185mm² · 74m"},
                    {"kiosk": "LM8D", "cable_in": "185mm² · 57m"},
                ],
            },
            {
                "name": "Feeder 4",
                "chain": [
                    {"kiosk": "LM9D", "cable_in": "250A · 185×4mm² · 31m",
                     "extras": [{"name": "SL Switchgear", "detail": "3x60A · RWB · no meter"}]},
                    {"kiosk": "LM10D", "cable_in": "185mm² · 64m"},
                    {"kiosk": "LM11D", "cable_in": "120mm² · 60m"},
                ],
            },
        ],
    },
]

# ---------- Helpers ----------
def find_sheet(xls, candidates):
    for name in xls.sheet_names:
        if name.strip().lower() in [c.lower() for c in candidates]:
            return name
    return None


def find_data_file():
    """The spreadsheet lives in this same repo folder — push an updated copy
    whenever it changes, and this picks up the most recently modified match."""
    matches = glob.glob(os.path.join(os.path.dirname(__file__), FILE_PATTERN))
    if not matches:
        return None
    return max(matches, key=os.path.getmtime)


def find_kml_file():
    """Auto-detect a KML file in the same folder as app.py (added later)."""
    matches = glob.glob(os.path.join(os.path.dirname(__file__), "*.kml"))
    if not matches:
        return None
    return max(matches, key=os.path.getmtime)


def fmt_serial(v):
    """Convert float/int serials like 98820079.0 → '98820079', blanking NaN."""
    if pd.isna(v):
        return ""
    try:
        return str(int(float(v)))
    except (ValueError, OverflowError):
        return str(v).strip()


@st.cache_data(show_spinner=False)
def load_data(file_path, _mtime):
    xls = pd.ExcelFile(file_path)
    sheet = find_sheet(xls, ELEC_SHEET_CANDIDATES)
    if sheet is None:
        return None

    raw = xls.parse(sheet)
    raw.columns = [str(c).strip() for c in raw.columns]

    df = pd.DataFrame()
    df["stand"] = raw["Stand Number"].astype(str).str.strip()
    df["serial"] = raw["Meter Serial"].apply(fmt_serial)
    df["parent"] = raw["Parent Meter"].apply(fmt_serial)
    df["manufacturer"] = raw.get("Manufacturer", pd.Series(dtype=str)).astype(str).str.strip()
    df["model"] = raw.get("Meter Model", pd.Series(dtype=str)).astype(str).str.strip()
    df["connection"] = raw.get("Connection size", pd.Series(dtype=str)).astype(str).str.strip()
    df["phase"] = raw.get("Phase connected", pd.Series(dtype=str)).astype(str).str.strip()
    df["kiosk"] = raw["Kiosk Number"].astype(str).str.strip()

    # Readings hook: today this is the opening reading; when live/latest
    # readings are added to the sheet (e.g. a "Latest reading [kWh]" column),
    # map them here and the kiosk/feeder/minisub roll-ups light up downstream.
    df["reading"] = pd.to_numeric(
        raw.get("Opening reading [kWh]", pd.Series(dtype=float)), errors="coerce"
    ).fillna(0)
    latest = None
    for cand in ["Latest reading [kWh]", "Latest Reading [kWh]", "Current reading [kWh]"]:
        if cand in raw.columns:
            latest = cand
            break
    df["latest_reading"] = (
        pd.to_numeric(raw[latest], errors="coerce") if latest else pd.Series([None] * len(raw))
    )

    dates = raw.get("Meter Commissioning Date")
    if dates is not None and pd.api.types.is_numeric_dtype(dates):
        df["commissioned"] = pd.to_datetime(dates, unit="D", origin="1899-12-30", errors="coerce")
    else:
        df["commissioned"] = pd.to_datetime(dates, errors="coerce")

    amr = raw.get("AMR Installed", pd.Series(dtype=object))
    df["amr"] = amr.map(lambda v: str(v).strip().upper() == "TRUE" if not isinstance(v, bool) else v).fillna(False)
    df["amr_port"] = raw.get("AMR Port", pd.Series(dtype=object)).apply(
        lambda v: "" if pd.isna(v) else fmt_serial(v)
    )

    # Classify rows: MUNIC bulk is its own parent; minisubs are parented to it.
    munic_mask = df["serial"] == df["parent"]
    munic_serials = set(df.loc[munic_mask, "serial"])

    def level(r):
        if r["serial"] in munic_serials:
            return "bulk"
        if r["parent"] in munic_serials:
            return "minisub"
        return "meter"

    df["level"] = df.apply(level, axis=1)
    return df


def meter_dict(r):
    return {
        "stand": r["stand"],
        "serial": r["serial"],
        "model": r["model"],
        "connection": r["connection"],
        "phase": r["phase"],
        "amr": bool(r["amr"]),
        "port": r["amr_port"],
        "date": r["commissioned"].strftime("%Y-%m-%d") if pd.notna(r["commissioned"]) else "",
        "reading": float(r["reading"]) if pd.notna(r["reading"]) else 0.0,
        "latest": float(r["latest_reading"]) if pd.notna(r["latest_reading"]) else None,
    }


def phase_split(kdf):
    counts = {"R": 0, "W": 0, "B": 0}
    for p in kdf["phase"]:
        p = str(p).strip().upper()
        if p in counts:
            counts[p] += 1
    return counts


def build_hierarchy(df):
    """Build MUNIC → MS → feeders → kiosk chains from RETICULATION config,
    attaching the spreadsheet's meters to each kiosk. Returns (munic, tree,
    unassigned_kiosks, parent_mismatches)."""
    bulk = df[df["level"] == "bulk"]
    ms_rows = {r["stand"]: r for _, r in df[df["level"] == "minisub"].iterrows()}
    meters = df[df["level"] == "meter"]
    by_kiosk = {k: g.sort_values("stand") for k, g in meters.groupby("kiosk")}

    configured_kiosks = set()
    mismatches = []
    tree = []

    for ms_cfg in RETICULATION:
        ms_row = ms_rows.get(ms_cfg["ms_name"])
        ms = {
            "ms_name": ms_cfg["ms_name"],
            "serial": ms_cfg["serial"],
            "rating": ms_cfg["rating"],
            "ct": ms_cfg["ct"],
            "note": ms_cfg["note"],
            "amr": bool(ms_row["amr"]) if ms_row is not None else False,
            "model": ms_row["model"] if ms_row is not None else "",
            "feeders": [],
        }

        for f_cfg in ms_cfg["feeders"]:
            feeder = {"name": f_cfg["name"], "chain": []}
            for pos, node in enumerate(f_cfg["chain"]):
                kname = node["kiosk"]
                configured_kiosks.add(kname)
                kdf = by_kiosk.get(kname, pd.DataFrame(columns=meters.columns))
                for _, r in kdf.iterrows():
                    if r["parent"] and r["parent"] != ms_cfg["serial"]:
                        mismatches.append({
                            "Stand": r["stand"], "Serial": r["serial"], "Kiosk": kname,
                            "Sheet Parent": r["parent"],
                            "Reticulation Parent": f"{ms_cfg['serial']} ({ms_cfg['ms_name']})",
                        })
                ph = phase_split(kdf)
                feeder["chain"].append({
                    "kiosk": kname,
                    "cable_in": node["cable_in"],
                    "eol": pos == len(f_cfg["chain"]) - 1,
                    "extras": node.get("extras", []),
                    "total": len(kdf),
                    "amr_count": int(kdf["amr"].sum()) if len(kdf) else 0,
                    "phase": ph,
                    "kwh": float(kdf["reading"].sum()) if len(kdf) else 0.0,
                    "latest_kwh": float(kdf["latest_reading"].sum()) if len(kdf) and kdf["latest_reading"].notna().any() else None,
                    "meters": [meter_dict(r) for _, r in kdf.iterrows()],
                })
            feeder["total"] = sum(k["total"] for k in feeder["chain"])
            feeder["amr_count"] = sum(k["amr_count"] for k in feeder["chain"])
            feeder["kwh"] = sum(k["kwh"] for k in feeder["chain"])
            ms["feeders"].append(feeder)

        ms["total"] = sum(f["total"] for f in ms["feeders"])
        ms["amr_count"] = sum(f["amr_count"] for f in ms["feeders"])
        ms["kwh"] = sum(f["kwh"] for f in ms["feeders"])
        allm = pd.concat([by_kiosk[k["kiosk"]] for f in ms["feeders"] for k in f["chain"] if k["kiosk"] in by_kiosk]) \
            if any(k["kiosk"] in by_kiosk for f in ms["feeders"] for k in f["chain"]) else pd.DataFrame(columns=meters.columns)
        ms["phase"] = phase_split(allm)
        tree.append(ms)

    unassigned = sorted(set(by_kiosk.keys()) - configured_kiosks)

    munic = None
    if not bulk.empty:
        b = bulk.iloc[0]
        munic = {
            "name": b["stand"], "serial": b["serial"],
            "manufacturer": b["manufacturer"], "connection": b["connection"],
            "amr": bool(b["amr"]),
        }
    return munic, tree, unassigned, mismatches


def kiosk_to_ms_feeder():
    """Lookup: kiosk name → (ms_name, ms_serial, feeder name, position label)."""
    lookup = {}
    for ms in RETICULATION:
        for f in ms["feeders"]:
            n = len(f["chain"])
            for pos, node in enumerate(f["chain"]):
                pos_label = "EOL" if pos == n - 1 else f"{pos + 1} of {n}"
                lookup[node["kiosk"]] = (ms["ms_name"], ms["serial"], f["name"], pos_label)
    return lookup


def show_table(view_df, columns, rename, sort_col=None, ascending=True):
    out = view_df[columns].rename(columns=rename)
    if sort_col:
        out = out.sort_values(sort_col, ascending=ascending)
    st.dataframe(out, use_container_width=True, hide_index=True)


# ---------- Load data ----------
data_path = find_data_file()
if data_path is None:
    st.error(f"No spreadsheet matching `{FILE_PATTERN}` found next to app.py. Push the file to the repo and redeploy.")
    st.stop()

mtime = os.path.getmtime(data_path)
df = load_data(data_path, mtime)
if df is None:
    st.error("Could not find an 'Elec' sheet in the spreadsheet.")
    st.stop()

munic, tree, unassigned_kiosks, parent_mismatches = build_hierarchy(df)
KIOSK_LOOKUP = kiosk_to_ms_feeder()

meters_df = df[df["level"] == "meter"]
bulk_df = df[df["level"] != "meter"]
has_readings = meters_df["latest_reading"].notna().any()

st.title(f"⚡ {SITE_NAME} — Site Hierarchy & Smart Metering")
st.caption(
    f"Source: `{os.path.basename(data_path)}` · Reticulation per as-built SLDs (MS LM-C / MS LM-D) · "
    f"All meters installed — tracking outstanding smart metering (AMR)."
)

# ---------- KPI strip ----------
total_meters = len(meters_df)
amr_done = int(meters_df["amr"].sum())
amr_outstanding = total_meters - amr_done
amr_pct = round(amr_done / total_meters * 100) if total_meters else 0
total_feeders = sum(len(ms["feeders"]) for ms in tree)

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Unit / common meters", total_meters)
c2.metric("Feeders", total_feeders, f"{meters_df['kiosk'].nunique()} kiosks")
c3.metric("Smart metering done", amr_done, f"{amr_pct}%")
c4.metric("Smart metering outstanding", amr_outstanding)
c5.metric("Bulk & check meters", len(bulk_df))

st.divider()

# ---------- Tabs ----------
tab_hierarchy, tab_amr, tab_all, tab_kiosks, tab_map = st.tabs(
    ["🗼 Site Hierarchy", "📡 Smart Metering (AMR)", "📋 All Meters", "🧰 Kiosk & Feeder Summary", "🗺️ Estate Map"]
)

# =====================================================================
# SITE HIERARCHY TAB — MUNIC → Minisubs → Feeders → daisy-chained kiosks
# =====================================================================
with tab_hierarchy:
    st.subheader("⚡ Electrical Reticulation — Single Line Diagram")
    st.caption("Municipal supply → MUNIC bulk meter → Minisubs → Feeders → daisy-chained kiosks (cable specs on each leg). "
               "Click a kiosk to expand its meters; click a meter chip for full details.")

    search = st.text_input(
        "🔍 Search by meter serial or stand number",
        placeholder="Type part of a serial or stand number to highlight it in the diagram",
        key="hier_search",
    )

    highlight_serials = []
    if search.strip():
        s = search.strip()
        match = meters_df[
            meters_df["serial"].str.contains(s, case=False, na=False)
            | meters_df["stand"].str.contains(s, case=False, na=False)
        ]
        if match.empty:
            st.warning(f"No meter found matching **{s}**.")
        else:
            highlight_serials = match["serial"].tolist()
            for _, row in match.iterrows():
                ms_name, _, feeder, pos = KIOSK_LOOKUP.get(row["kiosk"], ("?", "", "?", ""))
                amr_txt = f"AMR ✓ (port {row['amr_port']})" if row["amr"] else "AMR outstanding"
                st.success(
                    f"Stand **{row['stand']}** · Serial `{row['serial']}` · {row['model']} · "
                    f"Kiosk **{row['kiosk']}** ({pos}) · **{ms_name} / {feeder}** · {amr_txt} "
                    f"— highlighted below."
                )

    if not has_readings:
        st.caption("💡 Readings roll-up is built in: once a *Latest reading [kWh]* column is added to the sheet, "
                   "kiosk totals per feeder will show here and be compared against each sub-bulk check meter.")

    retic_json = json.dumps(tree)
    munic_json = json.dumps(munic)
    highlight_json = json.dumps(highlight_serials)
    has_readings_json = json.dumps(bool(has_readings))

    html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'IBM Plex Mono', 'Courier New', monospace; background: #0e1117; color: #e0e0e0; padding: 16px; }}

  .legend-bar {{
    display: flex; gap: 14px; margin-bottom: 14px; flex-wrap: wrap;
    font-size: 10px; align-items: center; padding: 8px 12px;
    background: #0d1520; border: 1px solid #1e3050; border-radius: 6px;
  }}
  .legend-item {{ display: flex; align-items: center; gap: 5px; }}
  .legend-swatch {{ width: 12px; height: 12px; border-radius: 3px; flex-shrink: 0; }}

  .supply-bus {{ display: flex; align-items: center; justify-content: center; }}
  .supply-box {{
    background: #1a2535; border: 2px solid #E69138; border-radius: 8px;
    padding: 10px 28px; font-size: 13px; font-weight: 700;
    color: #E69138; letter-spacing: .1em; text-transform: uppercase;
  }}
  .bus-line {{ height: 4px; background: #E69138; flex: 1; max-width: 300px; }}

  .munic-wrap {{ display: flex; flex-direction: column; align-items: center; }}
  .munic-vline {{ width: 3px; height: 24px; background: #E69138; }}
  .munic-box {{
    background: #2b1f10; border: 2px solid #E69138; border-radius: 10px;
    padding: 10px 24px; text-align: center; min-width: 320px;
  }}
  .munic-box .m-label {{ font-size: 11px; color: #d4a35c; letter-spacing: .08em; }}
  .munic-box .m-title {{ font-size: 15px; font-weight: 700; color: #fff; margin-top: 2px; }}
  .munic-box .m-serial {{ font-size: 10px; color: #b08a4f; margin-top: 3px; }}

  /* ---- Minisub section ---- */
  .ms-section {{
    border: 1px solid #1e3050; border-radius: 12px; background: #0d1520;
    padding: 18px 16px 20px 16px; margin-top: 26px;
  }}
  .ms-box {{
    background: #1F3F66; border: 2px solid #5B86B3; border-radius: 10px;
    padding: 12px 22px; text-align: center; max-width: 460px; margin: 0 auto;
  }}
  .ms-box .ms-label {{ font-size: 11px; color: #9FB0C2; letter-spacing: .08em; margin-bottom: 3px; }}
  .ms-box .ms-title {{ font-size: 16px; font-weight: 700; color: #FFFFFF; }}
  .ms-box .ms-serial {{ font-size: 10px; color: #7A96B2; margin-top: 3px; }}
  .ms-progress {{ margin-top: 8px; }}
  .progress-track {{ height: 5px; background: #2a3f55; border-radius: 3px; overflow: hidden; }}
  .progress-fill {{ height: 100%; background: #3F7D5C; border-radius: 3px; }}
  .ms-counts {{ font-size: 10px; color: #9FB0C2; margin-top: 4px; }}
  .amr-ok-count {{ color: #6eb88a; }}
  .amr-miss-count {{ color: #d4902a; }}
  .phase-pills {{ display: flex; justify-content: center; gap: 5px; margin-top: 6px; }}
  .phase-pill {{ font-size: 9px; font-weight: 700; border-radius: 3px; padding: 1px 7px; }}
  .phase-pill.r {{ background: #b3282833; color: #e07060; border: 1px solid #b3282866; }}
  .phase-pill.w {{ background: #d0d0d022; color: #cfcfcf; border: 1px solid #88888866; }}
  .phase-pill.b {{ background: #2d7dd233; color: #6aaef0; border: 1px solid #2d7dd266; }}

  .feeders-row {{
    display: flex; justify-content: center; gap: 22px;
    align-items: flex-start; margin-top: 0; flex-wrap: wrap;
  }}
  .feeder-col {{
    display: flex; flex-direction: column; align-items: center;
    min-width: 235px; max-width: 300px; flex: 1;
  }}
  .feeder-drop {{ width: 3px; height: 22px; background: #5B86B3; }}
  .feeder-head {{
    width: 100%; text-align: center; background: #14243a;
    border: 1.5px dashed #5B86B3; border-radius: 8px; padding: 6px 10px;
  }}
  .feeder-head .f-name {{ font-size: 11px; font-weight: 700; color: #9FC1E4; letter-spacing: .06em; }}
  .feeder-head .f-counts {{ font-size: 9px; color: #7A96B2; margin-top: 2px; }}
  .feeder-head .f-kwh {{ font-size: 9px; color: #6eb88a; margin-top: 2px; }}

  .cable-leg {{ display: flex; flex-direction: column; align-items: center; width: 100%; }}
  .cable-line {{ width: 3px; height: 16px; background: #5B86B3; }}
  .cable-label {{
    font-size: 8.5px; color: #7A96B2; background: #0e1117;
    border: 1px solid #23364f; border-radius: 4px; padding: 1px 7px; margin: 1px 0;
    white-space: nowrap;
  }}

  .kiosk-node {{
    width: 100%; border-radius: 8px; border: 1.5px solid #334d6e;
    background: #131c2b; cursor: pointer; padding: 9px 12px;
    transition: border-color .15s, background .15s;
  }}
  .kiosk-node:hover {{ border-color: #5B86B3; background: #1a2840; }}
  .kiosk-node.expanded {{ border-color: #E69138; background: #1e2b3a; }}
  .kiosk-node.all-amr {{ border-color: #3F7D5C; }}

  .kiosk-header {{ display: flex; align-items: center; gap: 8px; }}
  .kiosk-id {{ font-size: 12px; font-weight: 700; color: #c8d8eb; }}
  .kiosk-bar-wrap {{ flex: 1; }}
  .kiosk-mini-bar {{ height: 4px; border-radius: 2px; background: #2a3f55; overflow: hidden; }}
  .kiosk-mini-fill {{ height: 100%; border-radius: 2px; }}
  .kiosk-counts {{ font-size: 10px; color: #7A96B2; white-space: nowrap; }}
  .kiosk-chevron {{ font-size: 10px; color: #5B86B3; transition: transform .2s; }}
  .kiosk-chevron.open {{ transform: rotate(180deg); }}
  .kiosk-amr-line {{ font-size: 9px; color: #7A96B2; margin-top: 3px; display: flex; gap: 8px; flex-wrap: wrap; }}

  .eol-cap {{
    font-size: 8.5px; font-weight: 700; letter-spacing: .12em; color: #5B86B3;
    border: 1px solid #23364f; border-radius: 10px; padding: 1px 10px; margin-top: 6px;
    background: #0e1117;
  }}

  .kiosk-detail {{
    display: none; width: 100%; background: #0d1520; border: 1px solid #1e3050;
    border-top: none; border-radius: 0 0 8px 8px; padding: 8px 10px; font-size: 10px;
  }}
  .kiosk-detail.open {{ display: block; }}
  .stand-grid {{ display: flex; flex-wrap: wrap; gap: 4px; margin-top: 4px; }}
  .stand-chip {{
    padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 600;
    background: #3F7D5C22; color: #6eb88a; border: 1px solid #3F7D5C55;
    display: inline-flex; align-items: center; gap: 3px; cursor: pointer;
  }}
  .stand-chip.no-amr {{ background: #E6913822; color: #d4902a; border: 1px solid #E6913866; }}
  .stand-chip.extra {{
    background: #2a2f3a55; color: #8a93a3; border: 1px dashed #4a5567; cursor: default;
  }}
  .stand-chip.highlight {{ outline: 2px solid #E69138; box-shadow: 0 0 6px #E6913888; }}
  .amr-dot {{ display: inline-block; width: 6px; height: 6px; border-radius: 50%; flex-shrink: 0; }}
  .amr-dot.ok {{ background: #3F7D5C; }}
  .amr-dot.missing {{ background: #E69138; }}

  .serial-popup {{
    position: fixed; z-index: 9999;
    background: #152B45; border: 1px solid #5B86B3; border-radius: 8px;
    padding: 10px 14px; font-size: 11px; min-width: 230px;
    box-shadow: 0 6px 24px rgba(0,0,0,.5);
    display: none; pointer-events: none;
  }}
  .serial-popup.visible {{ display: block; }}
  .serial-popup .sp-stand {{ font-size: 13px; font-weight: 700; color: #fff; margin-bottom: 6px; }}
  .serial-popup .sp-row {{ display: flex; justify-content: space-between; gap: 16px; margin-bottom: 3px; }}
  .serial-popup .sp-label {{ color: #7A96B2; font-size: 10px; text-transform: uppercase; letter-spacing: .05em; }}
  .serial-popup .sp-val {{ color: #E0E8F0; font-family: monospace; font-size: 11px; }}
  .serial-popup .sp-amr-ok {{ color: #6eb88a; }}
  .serial-popup .sp-amr-miss {{ color: #d4902a; }}
</style>
</head>
<body>

<div class="serial-popup" id="serialPopup"></div>

<div class="legend-bar">
  <strong style="color:#9FB0C2;font-size:10px;letter-spacing:.06em;">METER STATUS:</strong>
  <div class="legend-item"><span class="legend-swatch" style="background:#3F7D5C33;border:1px solid #3F7D5C66"></span><span style="color:#6eb88a">Installed · Smart metering ✓</span></div>
  <div class="legend-item"><span class="legend-swatch" style="background:#E6913833;border:1px solid #E6913866"></span><span style="color:#d4902a">Installed · Smart metering outstanding</span></div>
  <div class="legend-item"><span class="legend-swatch" style="background:#2a2f3a55;border:1px dashed #4a5567"></span><span style="color:#8a93a3">Switchgear / no meter</span></div>
  <div class="legend-item"><span style="color:#7A96B2">— Cable specs shown on each feeder leg</span></div>
</div>

<div class="supply-bus">
  <div class="bus-line"></div>
  <div class="supply-box">⚡ Municipal Supply (MV)</div>
  <div class="bus-line"></div>
</div>

<div class="munic-wrap" id="munic"></div>

<div id="diagramRoot"></div>

<script>
const data = {retic_json};
const munic = {munic_json};
const highlightSerials = new Set({highlight_json});
const hasReadings = {has_readings_json};
const popup = document.getElementById('serialPopup');

function pct(a, b) {{ return b > 0 ? Math.round(a / b * 100) : 0; }}
function barColor(p) {{
  if (p >= 100) return '#3F7D5C';
  if (p >= 60) return '#5B86B3';
  if (p >= 30) return '#E69138';
  return '#BD4B2C';
}}
function fmtKwh(v) {{ return v.toLocaleString(undefined, {{maximumFractionDigits: 0}}) + ' kWh'; }}

function showPopup(e, m) {{
  const amrHtml = m.amr
    ? `<span class="sp-amr-ok">✓ Commissioned${{m.port ? ' · port ' + m.port : ''}}</span>`
    : `<span class="sp-amr-miss">⚠ Outstanding</span>`;
  const readingHtml = m.latest !== null && m.latest !== undefined
    ? `<div class="sp-row"><span class="sp-label">Latest reading</span><span class="sp-val">${{fmtKwh(m.latest)}}</span></div>`
    : `<div class="sp-row"><span class="sp-label">Opening reading</span><span class="sp-val">${{fmtKwh(m.reading)}}</span></div>`;
  popup.innerHTML = `
    <div class="sp-stand">Stand ${{m.stand}}</div>
    <div class="sp-row"><span class="sp-label">Meter serial</span><span class="sp-val">${{m.serial}}</span></div>
    <div class="sp-row"><span class="sp-label">Model</span><span class="sp-val">${{m.model}} · ${{m.connection}}</span></div>
    <div class="sp-row"><span class="sp-label">Phase</span><span class="sp-val">${{m.phase}}</span></div>
    <div class="sp-row"><span class="sp-label">Commissioned</span><span class="sp-val">${{m.date || '—'}}</span></div>
    ${{readingHtml}}
    <div class="sp-row"><span class="sp-label">Smart metering</span>${{amrHtml}}</div>
  `;
  popup.classList.add('visible');
  popup.style.left = Math.min(e.clientX + 12, window.innerWidth - 260) + 'px';
  popup.style.top = Math.min(e.clientY + 12, window.innerHeight - 190) + 'px';
  e.stopPropagation();
}}
document.addEventListener('click', () => popup.classList.remove('visible'));

if (munic) {{
  document.getElementById('munic').innerHTML = `
    <div class="munic-vline"></div>
    <div class="munic-box">
      <div class="m-label">Bulk Supply Meter</div>
      <div class="m-title">${{munic.name}}</div>
      <div class="m-serial">Serial: ${{munic.serial}} · ${{munic.manufacturer}} · ${{munic.connection}}</div>
    </div>
  `;
}}

function phasePills(ph) {{
  return `<div class="phase-pills">
    <span class="phase-pill r">R ${{ph.R}}</span>
    <span class="phase-pill w">W ${{ph.W}}</span>
    <span class="phase-pill b">B ${{ph.B}}</span>
  </div>`;
}}

function buildDiagram() {{
  const root = document.getElementById('diagramRoot');

  data.forEach(ms => {{
    const section = document.createElement('div');
    section.className = 'ms-section';

    const p = pct(ms.amr_count, ms.total);
    const kwhLine = hasReadings
      ? `<div class="ms-counts amr-ok-count">Kiosk totals: ${{fmtKwh(ms.kwh)}} · check vs bulk ${{ms.serial}}</div>`
      : '';
    section.innerHTML = `
      <div class="ms-box">
        <div class="ms-label">Minisub · ${{ms.rating}} · ${{ms.ct}} ${{ms.amr ? '· AMR ✓' : ''}}</div>
        <div class="ms-title">${{ms.ms_name}}</div>
        <div class="ms-serial">Bulk meter SN: ${{ms.serial}} ${{ms.model ? '· ' + ms.model : ''}} ${{ms.note ? '· ' + ms.note : ''}}</div>
        <div class="ms-progress">
          <div class="progress-track"><div class="progress-fill" style="width:${{p}}%; background:${{barColor(p)}}"></div></div>
          <div class="ms-counts">
            ${{ms.feeders.length}} feeders · ${{ms.total}} meters ·
            <span class="amr-ok-count">✓ ${{ms.amr_count}} smart</span>
            ${{(ms.total - ms.amr_count) > 0 ? ' &nbsp;<span class="amr-miss-count">⚠ ' + (ms.total - ms.amr_count) + ' outstanding</span>' : ''}}
          </div>
          ${{kwhLine}}
        </div>
        ${{phasePills(ms.phase)}}
      </div>
      <div style="display:flex;justify-content:center;"><div class="feeder-drop"></div></div>
    `;

    const row = document.createElement('div');
    row.className = 'feeders-row';

    ms.feeders.forEach(f => {{
      const col = document.createElement('div');
      col.className = 'feeder-col';

      const fKwh = hasReadings ? `<div class="f-kwh">${{fmtKwh(f.kwh)}}</div>` : '';
      col.innerHTML = `
        <div class="feeder-head">
          <div class="f-name">${{f.name.toUpperCase()}}</div>
          <div class="f-counts">${{f.chain.length}} kiosks · ${{f.total}} meters ·
            <span class="amr-ok-count">✓ ${{f.amr_count}}</span>
            ${{(f.total - f.amr_count) > 0 ? ' <span class="amr-miss-count">⚠ ' + (f.total - f.amr_count) + '</span>' : ''}}
          </div>
          ${{fKwh}}
        </div>
      `;

      f.chain.forEach(k => {{
        const leg = document.createElement('div');
        leg.className = 'cable-leg';
        leg.innerHTML = `
          <div class="cable-line"></div>
          <div class="cable-label">${{k.cable_in}}</div>
          <div class="cable-line"></div>
        `;
        col.appendChild(leg);

        const kp = pct(k.amr_count, k.total);
        const allAmr = k.amr_count >= k.total && k.total > 0;
        const kid = ms.serial + '-' + k.kiosk.replace(/\\s+/g, '_');

        const node = document.createElement('div');
        node.className = 'kiosk-node' + (allAmr ? ' all-amr' : '');
        const kwhTag = hasReadings ? `<span class="amr-ok-count">${{fmtKwh(k.kwh)}}</span>` : '';
        node.innerHTML = `
          <div class="kiosk-header">
            <span class="kiosk-id">${{k.kiosk}}</span>
            <div class="kiosk-bar-wrap">
              <div class="kiosk-mini-bar">
                <div class="kiosk-mini-fill" style="width:${{kp}}%; background:${{barColor(kp)}}"></div>
              </div>
            </div>
            <span class="kiosk-counts">${{k.amr_count}}/${{k.total}} smart</span>
            <span class="kiosk-chevron" id="chev-${{kid}}">▾</span>
          </div>
          <div class="kiosk-amr-line">
            <span>${{k.total}} meters</span>
            <span class="amr-ok-count">✓ ${{k.amr_count}}</span>
            ${{(k.total - k.amr_count) > 0 ? '<span class="amr-miss-count">⚠ ' + (k.total - k.amr_count) + '</span>' : ''}}
            <span>R${{k.phase.R}} W${{k.phase.W}} B${{k.phase.B}}</span>
            ${{kwhTag}}
          </div>
        `;

        const detail = document.createElement('div');
        detail.className = 'kiosk-detail';
        detail.id = 'detail-' + kid;

        const chipsHtml = k.meters.map(m => {{
          let cls = m.amr ? 'stand-chip' : 'stand-chip no-amr';
          if (highlightSerials.has(m.serial)) cls += ' highlight';
          const dot = `<span class="amr-dot ${{m.amr ? 'ok' : 'missing'}}"></span>`;
          const title = `Stand ${{m.stand}} · Serial: ${{m.serial}} · AMR: ${{m.amr ? '✓' + (m.port ? ' port ' + m.port : '') : 'outstanding'}}`;
          return `<span class="${{cls}}" data-m='${{JSON.stringify(m).replace(/'/g, "&#39;")}}' title="${{title}}">${{dot}}${{m.stand}}</span>`;
        }}).join('');

        const extrasHtml = (k.extras || []).map(x =>
          `<span class="stand-chip extra" title="${{x.detail}}">⚙ ${{x.name}}</span>`
        ).join('');

        detail.innerHTML = `
          <div style="font-size:9px;color:#5B86B3;margin-bottom:4px;">
            METERS (${{k.total}})
            &nbsp;·&nbsp; <span class="amr-ok-count">Smart done: ${{k.amr_count}}</span>
            ${{(k.total - k.amr_count) > 0 ? '&nbsp;·&nbsp; <span class="amr-miss-count">Outstanding: ' + (k.total - k.amr_count) + '</span>' : ''}}
          </div>
          <div class="stand-grid">${{chipsHtml}}${{extrasHtml}}</div>
        `;

        detail.querySelectorAll('.stand-chip:not(.extra)').forEach(chip => {{
          chip.addEventListener('click', function(e) {{
            showPopup(e, JSON.parse(chip.dataset.m));
          }});
        }});

        node.addEventListener('click', function() {{
          const open = detail.classList.toggle('open');
          node.classList.toggle('expanded', open);
          const chev = document.getElementById('chev-' + kid);
          if (chev) chev.classList.toggle('open', open);
        }});

        col.appendChild(node);
        col.appendChild(detail);

        if (k.eol) {{
          const cap = document.createElement('div');
          cap.className = 'eol-cap';
          cap.textContent = 'EOL';
          col.appendChild(cap);
        }}
      }});

      row.appendChild(col);
    }});

    section.appendChild(row);
    root.appendChild(section);
  }});
}}

buildDiagram();

// Auto-expand kiosks containing highlighted serials
if (highlightSerials.size > 0) {{
  data.forEach(ms => {{
    ms.feeders.forEach(f => {{
      f.chain.forEach(k => {{
        if (k.meters.some(m => highlightSerials.has(m.serial))) {{
          const kid = ms.serial + '-' + k.kiosk.replace(/\\s+/g, '_');
          const detail = document.getElementById('detail-' + kid);
          const chev = document.getElementById('chev-' + kid);
          if (detail) {{
            detail.classList.add('open');
            detail.previousElementSibling.classList.add('expanded');
          }}
          if (chev) chev.classList.add('open');
        }}
      }});
    }});
  }});
}}
</script>
</body>
</html>
"""

    components.html(html, height=1250, scrolling=True)

    # ---- Data checks: sheet vs as-built reticulation ----
    issues = len(parent_mismatches) + len(unassigned_kiosks)
    with st.expander(f"🔎 Data checks — sheet vs as-built reticulation ({issues} flag{'s' if issues != 1 else ''})"):
        if unassigned_kiosks:
            st.warning(f"Kiosks in the sheet but not in the configured reticulation: {', '.join(unassigned_kiosks)}")
        if parent_mismatches:
            st.warning(
                f"{len(parent_mismatches)} meters where the sheet's **Parent Meter** column disagrees with the "
                "as-built reticulation (the diagram above follows the as-builts). "
                "Worth aligning the sheet when you do the LM7D relabelling update:"
            )
            st.dataframe(pd.DataFrame(parent_mismatches), use_container_width=True, hide_index=True)
        if not issues:
            st.success("Sheet parent references match the as-built reticulation. ✅")

# =====================================================================
# SMART METERING (AMR) TAB
# =====================================================================
with tab_amr:
    st.subheader("📡 Smart Metering Status")
    st.caption("All meters on the estate are installed. This tab tracks which units still need AMR / smart metering commissioned.")

    fc1, fc2, fc3 = st.columns(3)
    ms_options = ["All minisubs"] + [ms["ms_name"] for ms in tree]
    sel_ms = fc1.selectbox("Minisub", ms_options, key="amr_ms")

    feeder_options = ["All feeders"]
    if sel_ms != "All minisubs":
        feeder_options += [f["name"] for ms in tree if ms["ms_name"] == sel_ms for f in ms["feeders"]]
    sel_feeder = fc2.selectbox("Feeder", feeder_options, key="amr_feeder")

    kiosk_options = ["All kiosks"] + sorted(meters_df["kiosk"].unique(), key=lambda k: (len(k), k))
    sel_kiosk = fc3.selectbox("Kiosk", kiosk_options, key="amr_kiosk")

    view = meters_df.copy()
    view = view.assign(
        ms_name=view["kiosk"].map(lambda k: KIOSK_LOOKUP.get(k, ("—",))[0]),
        feeder=view["kiosk"].map(lambda k: KIOSK_LOOKUP.get(k, ("—", "", "—"))[2]),
    )
    if sel_ms != "All minisubs":
        view = view[view["ms_name"] == sel_ms]
    if sel_feeder != "All feeders":
        view = view[view["feeder"] == sel_feeder]
    if sel_kiosk != "All kiosks":
        view = view[view["kiosk"] == sel_kiosk]

    out_view = view[~view["amr"]]
    done_view = view[view["amr"]]

    m1, m2, m3 = st.columns(3)
    m1.metric("Meters in view", len(view))
    m2.metric("Smart metering done", len(done_view))
    m3.metric("Outstanding", len(out_view))

    st.markdown("#### ⚠️ Outstanding smart metering")
    if out_view.empty:
        st.success("No outstanding smart metering in this view. 🎉")
    else:
        show_table(
            out_view,
            ["stand", "serial", "kiosk", "feeder", "ms_name", "model", "connection", "phase"],
            {"stand": "Stand", "serial": "Meter Serial", "kiosk": "Kiosk", "feeder": "Feeder",
             "ms_name": "Minisub", "model": "Model", "connection": "Connection", "phase": "Phase"},
            sort_col="Kiosk",
        )
        st.download_button(
            "⬇️ Download outstanding list (CSV)",
            out_view[["stand", "serial", "kiosk", "feeder", "ms_name", "model", "connection", "phase"]].to_csv(index=False),
            file_name="lake_michelle_amr_outstanding.csv",
            mime="text/csv",
        )

    with st.expander(f"✅ Smart metering completed ({len(done_view)})"):
        show_table(
            done_view,
            ["stand", "serial", "kiosk", "feeder", "ms_name", "amr_port", "model", "phase"],
            {"stand": "Stand", "serial": "Meter Serial", "kiosk": "Kiosk", "feeder": "Feeder",
             "ms_name": "Minisub", "amr_port": "AMR Port", "model": "Model", "phase": "Phase"},
            sort_col="Kiosk",
        )

# =====================================================================
# ALL METERS TAB
# =====================================================================
with tab_all:
    st.subheader("📋 All Meters")

    q = st.text_input("Search stand / serial / kiosk", key="all_search", placeholder="e.g. 057, 91546282, LM3D, Pump")
    view = df.copy()
    if q.strip():
        s = q.strip()
        view = view[
            view["stand"].str.contains(s, case=False, na=False)
            | view["serial"].str.contains(s, case=False, na=False)
            | view["kiosk"].str.contains(s, case=False, na=False)
        ]

    level_labels = {"bulk": "Bulk (MUNIC)", "minisub": "Minisub check meter", "meter": "Unit / common meter"}
    view = view.assign(
        level_label=view["level"].map(level_labels),
        amr_label=view["amr"].map({True: "✓ Done", False: "⚠ Outstanding"}),
        feeder=view["kiosk"].map(lambda k: KIOSK_LOOKUP.get(k, ("", "", "—"))[2]),
    )

    show_table(
        view,
        ["stand", "serial", "kiosk", "feeder", "level_label", "manufacturer", "model", "connection",
         "phase", "commissioned", "amr_label", "amr_port"],
        {"stand": "Stand", "serial": "Meter Serial", "kiosk": "Kiosk", "feeder": "Feeder",
         "level_label": "Level", "manufacturer": "Manufacturer", "model": "Model",
         "connection": "Connection", "phase": "Phase", "commissioned": "Commissioned",
         "amr_label": "Smart Metering", "amr_port": "AMR Port"},
        sort_col="Kiosk",
    )
    st.caption(f"{len(view)} of {len(df)} rows shown.")

# =====================================================================
# KIOSK & FEEDER SUMMARY TAB
# =====================================================================
with tab_kiosks:
    st.subheader("🧰 Kiosk & Feeder Summary")

    rows = []
    for ms in tree:
        for f in ms["feeders"]:
            for k in f["chain"]:
                _, _, _, pos = KIOSK_LOOKUP[k["kiosk"]]
                rows.append({
                    "Minisub": ms["ms_name"],
                    "Feeder": f["name"],
                    "Kiosk": k["kiosk"],
                    "Position": pos,
                    "Cable In": k["cable_in"],
                    "Meters": k["total"],
                    "Smart Done": k["amr_count"],
                    "Outstanding": k["total"] - k["amr_count"],
                    "Smart %": round(k["amr_count"] / k["total"] * 100) if k["total"] else 0,
                })
    kiosk_df = pd.DataFrame(rows)

    st.dataframe(
        kiosk_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Smart %": st.column_config.ProgressColumn("Smart %", min_value=0, max_value=100, format="%d%%"),
        },
    )

    st.markdown("#### Smart metering per feeder")
    feeder_rows = []
    for ms in tree:
        for f in ms["feeders"]:
            feeder_rows.append({
                "Feeder": f"{ms['ms_name']} · {f['name']}",
                "Smart Metering Done": f["amr_count"],
                "Outstanding": f["total"] - f["amr_count"],
            })
    st.bar_chart(pd.DataFrame(feeder_rows).set_index("Feeder"), color=["#E69138", "#3F7D5C"], horizontal=True)

    st.markdown("#### Smart metering by kiosk")
    chart_df = kiosk_df.set_index("Kiosk")[["Smart Done", "Outstanding"]]
    st.bar_chart(chart_df, color=["#E69138", "#3F7D5C"])

    if has_readings:
        st.markdown("#### ⚖️ Kiosk totals vs sub-bulk check meters")
        check_rows = []
        for ms in tree:
            check_rows.append({
                "Minisub": ms["ms_name"],
                "Sum of kiosk meters [kWh]": ms["kwh"],
                "Bulk check meter SN": ms["serial"],
            })
        st.dataframe(pd.DataFrame(check_rows), use_container_width=True, hide_index=True)
    else:
        st.caption("⚖️ Once meter readings are available in the sheet, this tab will total usage per kiosk "
                   "and feeder and compare each minisub's downstream sum against its bulk check meter.")

# =====================================================================
# ESTATE MAP TAB — placeholder until KML is available
# =====================================================================
with tab_map:
    st.subheader("🗺️ Estate Map")
    kml_path = find_kml_file()
    if kml_path is None:
        st.info(
            "📍 **No KML data for Lake Michelle yet.**\n\n"
            "Once the estate KML is ready, push it to this repo folder (any `*.kml` file) — "
            "the app auto-detects it and this tab can then be wired up with the estate map, "
            "the same way the Sitari app works."
        )
    else:
        st.success(f"KML file detected: `{os.path.basename(kml_path)}`. Map rendering to be wired up next.")