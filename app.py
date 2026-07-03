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
.block-container {padding-top: 1.2rem; max-width: 1300px;}
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
.badge {display:inline-block; padding:2px 9px; border-radius:14px; font-size:11px; font-weight:600; font-family: monospace;}
.badge-amr {background:#E3EEE6; color:#3F7D5C;}
.badge-noamr {background:#FCEFDD; color:#B96E1E;}
.badge-bulk {background:#E7EEF5; color:#1F3F66;}
</style>
""", unsafe_allow_html=True)

SITE_NAME = "Evergreen Lake Michelle"
ELEC_SHEET_CANDIDATES = ["Elec", "Elec Meters", "Electrical Meters"]
FILE_PATTERN = "EVG_Lake_Michelle_Meter_Hierarchy_*.xlsx"

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

    df["opening_reading"] = pd.to_numeric(
        raw.get("Opening reading [kWh]", pd.Series(dtype=float)), errors="coerce"
    )

    dates = raw.get("Meter Commissioning Date")
    if dates is not None and pd.api.types.is_numeric_dtype(dates):
        # Excel serial dates
        df["commissioned"] = pd.to_datetime(dates, unit="D", origin="1899-12-30", errors="coerce")
    else:
        df["commissioned"] = pd.to_datetime(dates, errors="coerce")

    amr = raw.get("AMR Installed", pd.Series(dtype=object))
    df["amr"] = amr.map(lambda v: str(v).strip().upper() == "TRUE" if not isinstance(v, bool) else v).fillna(False)
    df["amr_port"] = raw.get("AMR Port", pd.Series(dtype=object)).apply(
        lambda v: "" if pd.isna(v) else fmt_serial(v)
    )

    # ---- Classify each row in the hierarchy ----
    # The MUNIC bulk meter is its own parent (top of the tree).
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


def build_hierarchy(df):
    """MUNIC → minisubs → kiosks → meters, driven purely by the Parent Meter column."""
    bulk = df[df["level"] == "bulk"]
    minisubs = df[df["level"] == "minisub"].sort_values("stand")
    meters = df[df["level"] == "meter"]

    tree = []
    for _, ms in minisubs.iterrows():
        children = meters[meters["parent"] == ms["serial"]]
        kiosks = []
        for kiosk_name, kdf in children.groupby("kiosk"):
            kdf = kdf.sort_values("stand")
            kiosks.append({
                "kiosk": kiosk_name,
                "total": len(kdf),
                "amr_count": int(kdf["amr"].sum()),
                "meters": [
                    {
                        "stand": r["stand"],
                        "serial": r["serial"],
                        "model": r["model"],
                        "connection": r["connection"],
                        "phase": r["phase"],
                        "amr": bool(r["amr"]),
                        "port": r["amr_port"],
                        "date": r["commissioned"].strftime("%Y-%m-%d") if pd.notna(r["commissioned"]) else "",
                    }
                    for _, r in kdf.iterrows()
                ],
            })
        kiosks.sort(key=lambda k: (len(k["kiosk"]), k["kiosk"]))
        tree.append({
            "ms_name": ms["stand"],
            "serial": ms["serial"],
            "kiosk_label": ms["kiosk"],
            "amr": bool(ms["amr"]),
            "model": ms["model"],
            "connection": ms["connection"],
            "total": sum(k["total"] for k in kiosks),
            "amr_count": sum(k["amr_count"] for k in kiosks),
            "kiosks": kiosks,
        })

    munic = None
    if not bulk.empty:
        b = bulk.iloc[0]
        munic = {
            "name": b["stand"],
            "serial": b["serial"],
            "manufacturer": b["manufacturer"],
            "connection": b["connection"],
            "amr": bool(b["amr"]),
        }
    return munic, tree


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

munic, tree = build_hierarchy(df)

meters_df = df[df["level"] == "meter"]
bulk_df = df[df["level"] != "meter"]

st.title(f"⚡ {SITE_NAME} — Site Hierarchy & Smart Metering")
st.caption(
    f"Source: `{os.path.basename(data_path)}` · "
    f"All meters on the estate are installed — tracking outstanding smart metering (AMR)."
)

# ---------- KPI strip ----------
total_meters = len(meters_df)
amr_done = int(meters_df["amr"].sum())
amr_outstanding = total_meters - amr_done
amr_pct = round(amr_done / total_meters * 100) if total_meters else 0

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Unit / common meters", total_meters)
c2.metric("Bulk & check meters", len(bulk_df))
c3.metric("Smart metering done", amr_done, f"{amr_pct}%")
c4.metric("Smart metering outstanding", amr_outstanding)
c5.metric("Kiosks", meters_df["kiosk"].nunique())

st.divider()

# ---------- Tabs ----------
tab_hierarchy, tab_amr, tab_all, tab_kiosks, tab_map = st.tabs(
    ["🗼 Site Hierarchy", "📡 Smart Metering (AMR)", "📋 All Meters", "🧰 Kiosk Summary", "🗺️ Estate Map"]
)

# =====================================================================
# SITE HIERARCHY TAB — single-line diagram: MUNIC → Minisubs → Kiosks → Meters
# =====================================================================
with tab_hierarchy:
    st.subheader("⚡ Electrical Hierarchy — Single Line Diagram")
    st.caption("Municipal supply → MUNIC bulk meter → Minisub check meters → Kiosks → Unit meters. "
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
                ms_label = ""
                for ms in tree:
                    if any(m["serial"] == row["serial"] for k in ms["kiosks"] for m in k["meters"]):
                        ms_label = f" · fed from **{ms['ms_name']}**"
                amr_txt = f"AMR ✓ (port {row['amr_port']})" if row["amr"] else "AMR outstanding"
                st.success(
                    f"Stand **{row['stand']}** · Serial `{row['serial']}` · {row['model']} · "
                    f"Kiosk **{row['kiosk']}**{ms_label} · {amr_txt} — the matching chip is highlighted below."
                )

    st.divider()

    diagram_json = json.dumps(tree)
    munic_json = json.dumps(munic)
    highlight_json = json.dumps(highlight_serials)

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
  .munic-vline {{ width: 3px; height: 26px; background: #E69138; }}
  .munic-box {{
    background: #2b1f10; border: 2px solid #E69138; border-radius: 10px;
    padding: 10px 24px; text-align: center; min-width: 300px;
  }}
  .munic-box .m-label {{ font-size: 11px; color: #d4a35c; letter-spacing: .08em; }}
  .munic-box .m-title {{ font-size: 15px; font-weight: 700; color: #fff; margin-top: 2px; }}
  .munic-box .m-serial {{ font-size: 10px; color: #b08a4f; margin-top: 3px; }}

  .minisubs-row {{ display: flex; justify-content: center; gap: 32px; align-items: flex-start; }}

  .ms-col {{ display: flex; flex-direction: column; align-items: center; min-width: 260px; max-width: 420px; flex: 1; }}
  .ms-vert-line {{ width: 3px; height: 28px; background: #E69138; }}
  .ms-box {{
    background: #1F3F66; border: 2px solid #5B86B3; border-radius: 10px;
    padding: 12px 18px; text-align: center; width: 100%; position: relative;
  }}
  .ms-box .ms-label {{ font-size: 12px; color: #9FB0C2; letter-spacing: .08em; margin-bottom: 4px; }}
  .ms-box .ms-title {{ font-size: 15px; font-weight: 700; color: #FFFFFF; }}
  .ms-box .ms-serial {{ font-size: 10px; color: #7A96B2; margin-top: 3px; }}
  .ms-progress {{ margin-top: 8px; }}
  .progress-track {{ height: 5px; background: #2a3f55; border-radius: 3px; overflow: hidden; }}
  .progress-fill {{ height: 100%; background: #3F7D5C; border-radius: 3px; }}
  .ms-counts {{ font-size: 10px; color: #9FB0C2; margin-top: 4px; }}
  .amr-ok-count {{ color: #6eb88a; }}
  .amr-miss-count {{ color: #d4902a; }}

  .kiosk-connector {{ width: 3px; height: 20px; background: #5B86B3; }}
  .kiosk-grid {{ display: flex; flex-direction: column; width: 100%; }}
  .kiosk-entry {{ display: flex; flex-direction: column; align-items: center; width: 100%; }}
  .kiosk-drop-line {{ width: 3px; height: 18px; background: #5B86B3; }}

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
  .kiosk-amr-line {{ font-size: 9px; color: #7A96B2; margin-top: 3px; display: flex; gap: 8px; }}

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
  .stand-chip.highlight {{ outline: 2px solid #E69138; box-shadow: 0 0 6px #E6913888; }}
  .amr-dot {{ display: inline-block; width: 6px; height: 6px; border-radius: 50%; flex-shrink: 0; }}
  .amr-dot.ok {{ background: #3F7D5C; }}
  .amr-dot.missing {{ background: #E69138; }}

  .serial-popup {{
    position: fixed; z-index: 9999;
    background: #152B45; border: 1px solid #5B86B3; border-radius: 8px;
    padding: 10px 14px; font-size: 11px; min-width: 220px;
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
</div>

<div class="supply-bus">
  <div class="bus-line"></div>
  <div class="supply-box">⚡ Municipal Supply (MV)</div>
  <div class="bus-line"></div>
</div>

<div class="munic-wrap" id="munic"></div>

<div class="minisubs-row" id="diagramRoot"></div>

<script>
const data = {diagram_json};
const munic = {munic_json};
const highlightSerials = new Set({highlight_json});
const popup = document.getElementById('serialPopup');

function pct(a, b) {{ return b > 0 ? Math.round(a / b * 100) : 0; }}
function barColor(p) {{
  if (p >= 100) return '#3F7D5C';
  if (p >= 60) return '#5B86B3';
  if (p >= 30) return '#E69138';
  return '#BD4B2C';
}}

function showPopup(e, m) {{
  const amrHtml = m.amr
    ? `<span class="sp-amr-ok">✓ Commissioned${{m.port ? ' · port ' + m.port : ''}}</span>`
    : `<span class="sp-amr-miss">⚠ Outstanding</span>`;
  popup.innerHTML = `
    <div class="sp-stand">Stand ${{m.stand}}</div>
    <div class="sp-row"><span class="sp-label">Meter serial</span><span class="sp-val">${{m.serial}}</span></div>
    <div class="sp-row"><span class="sp-label">Model</span><span class="sp-val">${{m.model}} · ${{m.connection}}</span></div>
    <div class="sp-row"><span class="sp-label">Phase</span><span class="sp-val">${{m.phase}}</span></div>
    <div class="sp-row"><span class="sp-label">Commissioned</span><span class="sp-val">${{m.date || '—'}}</span></div>
    <div class="sp-row"><span class="sp-label">Smart metering</span>${{amrHtml}}</div>
  `;
  popup.classList.add('visible');
  popup.style.left = Math.min(e.clientX + 12, window.innerWidth - 250) + 'px';
  popup.style.top = Math.min(e.clientY + 12, window.innerHeight - 170) + 'px';
  e.stopPropagation();
}}
document.addEventListener('click', () => popup.classList.remove('visible'));

// ---- MUNIC bulk meter box ----
if (munic) {{
  document.getElementById('munic').innerHTML = `
    <div class="munic-vline"></div>
    <div class="munic-box">
      <div class="m-label">Bulk Supply Meter</div>
      <div class="m-title">${{munic.name}}</div>
      <div class="m-serial">Serial: ${{munic.serial}} · ${{munic.manufacturer}} · ${{munic.connection}}</div>
    </div>
    <div class="munic-vline"></div>
  `;
}}

function buildDiagram() {{
  const root = document.getElementById('diagramRoot');

  data.forEach(ms => {{
    const col = document.createElement('div');
    col.className = 'ms-col';

    const vline = document.createElement('div');
    vline.className = 'ms-vert-line';
    col.appendChild(vline);

    const p = pct(ms.amr_count, ms.total);
    const msBox = document.createElement('div');
    msBox.className = 'ms-box';
    msBox.innerHTML = `
      <div class="ms-label">Minisub Check Meter ${{ms.amr ? '· AMR ✓' : ''}}</div>
      <div class="ms-title">${{ms.ms_name}}</div>
      <div class="ms-serial">Serial: ${{ms.serial}} · ${{ms.model}} · ${{ms.connection}}</div>
      <div class="ms-progress">
        <div class="progress-track"><div class="progress-fill" style="width:${{p}}%; background:${{barColor(p)}}"></div></div>
        <div class="ms-counts">Smart metering: ${{ms.amr_count}} / ${{ms.total}} (${{p}}%)</div>
        <div class="ms-counts" style="margin-top:2px;">
          <span class="amr-ok-count">✓ ${{ms.amr_count}} done</span>
          ${{(ms.total - ms.amr_count) > 0 ? ' &nbsp;<span class="amr-miss-count">⚠ ' + (ms.total - ms.amr_count) + ' outstanding</span>' : ''}}
        </div>
      </div>
    `;
    col.appendChild(msBox);

    const conn = document.createElement('div');
    conn.className = 'kiosk-connector';
    col.appendChild(conn);

    const grid = document.createElement('div');
    grid.className = 'kiosk-grid';

    ms.kiosks.forEach(k => {{
      const entry = document.createElement('div');
      entry.className = 'kiosk-entry';

      const dropLine = document.createElement('div');
      dropLine.className = 'kiosk-drop-line';
      entry.appendChild(dropLine);

      const kp = pct(k.amr_count, k.total);
      const allAmr = k.amr_count >= k.total && k.total > 0;
      const kid = ms.serial + '-' + k.kiosk.replace(/\\s+/g, '_');

      const node = document.createElement('div');
      node.className = 'kiosk-node' + (allAmr ? ' all-amr' : '');
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
          <span class="amr-ok-count">✓ ${{k.amr_count}} AMR</span>
          ${{(k.total - k.amr_count) > 0 ? '<span class="amr-miss-count">⚠ ' + (k.total - k.amr_count) + ' outstanding</span>' : ''}}
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

      detail.innerHTML = `
        <div style="font-size:9px;color:#5B86B3;margin-bottom:4px;">
          METERS (${{k.total}})
          &nbsp;·&nbsp; <span class="amr-ok-count">Smart metering done: ${{k.amr_count}}</span>
          ${{(k.total - k.amr_count) > 0 ? '&nbsp;·&nbsp; <span class="amr-miss-count">Outstanding: ' + (k.total - k.amr_count) + '</span>' : ''}}
        </div>
        <div class="stand-grid">${{chipsHtml}}</div>
      `;

      detail.querySelectorAll('.stand-chip').forEach(chip => {{
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

      entry.appendChild(node);
      entry.appendChild(detail);
      grid.appendChild(entry);
    }});

    col.appendChild(grid);
    root.appendChild(col);
  }});
}}

buildDiagram();

// Auto-expand kiosks containing highlighted serials
if (highlightSerials.size > 0) {{
  data.forEach(ms => {{
    ms.kiosks.forEach(k => {{
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
}}
</script>
</body>
</html>
"""

    components.html(html, height=950, scrolling=True)

    st.info(
        "ℹ️ Kiosks are grouped under the minisub recorded in the **Parent Meter** column of the sheet. "
        "A kiosk whose meters are split across both minisubs (e.g. LM7D / LM8D) will appear under each, "
        "showing only the meters it feeds from that minisub."
    )

# =====================================================================
# SMART METERING (AMR) TAB
# =====================================================================
with tab_amr:
    st.subheader("📡 Smart Metering Status")
    st.caption("All meters on the estate are installed. This tab tracks which units still need AMR / smart metering commissioned.")

    fc1, fc2, fc3 = st.columns([1.2, 1.2, 2])
    kiosk_options = ["All kiosks"] + sorted(meters_df["kiosk"].unique(), key=lambda k: (len(k), k))
    ms_options = ["All minisubs"] + [ms["ms_name"] for ms in tree]
    sel_kiosk = fc1.selectbox("Kiosk", kiosk_options, key="amr_kiosk")
    sel_ms = fc2.selectbox("Fed from minisub", ms_options, key="amr_ms")

    view = meters_df.copy()
    if sel_kiosk != "All kiosks":
        view = view[view["kiosk"] == sel_kiosk]
    if sel_ms != "All minisubs":
        ms_serial = next(ms["serial"] for ms in tree if ms["ms_name"] == sel_ms)
        view = view[view["parent"] == ms_serial]

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
            ["stand", "serial", "kiosk", "model", "connection", "phase", "commissioned"],
            {"stand": "Stand", "serial": "Meter Serial", "kiosk": "Kiosk", "model": "Model",
             "connection": "Connection", "phase": "Phase", "commissioned": "Meter Commissioned"},
            sort_col="Kiosk",
        )
        st.download_button(
            "⬇️ Download outstanding list (CSV)",
            out_view[["stand", "serial", "kiosk", "model", "connection", "phase"]].to_csv(index=False),
            file_name="lake_michelle_amr_outstanding.csv",
            mime="text/csv",
        )

    with st.expander(f"✅ Smart metering completed ({len(done_view)})"):
        show_table(
            done_view,
            ["stand", "serial", "kiosk", "amr_port", "model", "phase", "commissioned"],
            {"stand": "Stand", "serial": "Meter Serial", "kiosk": "Kiosk", "amr_port": "AMR Port",
             "model": "Model", "phase": "Phase", "commissioned": "Meter Commissioned"},
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
    view = view.assign(level_label=view["level"].map(level_labels))
    view = view.assign(amr_label=view["amr"].map({True: "✓ Done", False: "⚠ Outstanding"}))

    show_table(
        view,
        ["stand", "serial", "kiosk", "level_label", "manufacturer", "model", "connection",
         "phase", "commissioned", "amr_label", "amr_port"],
        {"stand": "Stand", "serial": "Meter Serial", "kiosk": "Kiosk", "level_label": "Level",
         "manufacturer": "Manufacturer", "model": "Model", "connection": "Connection",
         "phase": "Phase", "commissioned": "Commissioned", "amr_label": "Smart Metering", "amr_port": "AMR Port"},
        sort_col="Kiosk",
    )
    st.caption(f"{len(view)} of {len(df)} rows shown.")

# =====================================================================
# KIOSK SUMMARY TAB
# =====================================================================
with tab_kiosks:
    st.subheader("🧰 Kiosk Summary")

    rows = []
    for ms in tree:
        for k in ms["kiosks"]:
            rows.append({
                "Kiosk": k["kiosk"],
                "Fed From": ms["ms_name"],
                "Meters": k["total"],
                "Smart Metering Done": k["amr_count"],
                "Outstanding": k["total"] - k["amr_count"],
                "Smart %": round(k["amr_count"] / k["total"] * 100) if k["total"] else 0,
            })
    kiosk_df = pd.DataFrame(rows).sort_values(["Fed From", "Kiosk"], key=lambda c: c.map(lambda v: (len(str(v)), str(v))))

    st.dataframe(
        kiosk_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Smart %": st.column_config.ProgressColumn("Smart %", min_value=0, max_value=100, format="%d%%"),
        },
    )

    st.markdown("#### Smart metering by kiosk")
    chart_df = kiosk_df.set_index("Kiosk")[["Smart Metering Done", "Outstanding"]]
    st.bar_chart(chart_df, color=["#3F7D5C", "#E69138"])

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