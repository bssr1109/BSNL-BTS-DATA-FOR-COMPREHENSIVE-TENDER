import os
from io import BytesIO
from datetime import datetime, timezone

import pandas as pd
import streamlit as st
from supabase import create_client


st.set_page_config(page_title="BSNL BTS Field Survey", layout="centered")

st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 5rem; max-width: 650px;}
[data-testid="stSidebar"] {display: none;}
div.stButton > button {width: 100%; min-height: 48px; font-size: 16px; border-radius: 10px;}
.stSelectbox label, .stTextInput label, .stTextArea label, .stNumberInput label {font-weight: 700;}
.site-card {border:1px solid #ddd;border-radius:14px;padding:14px;margin:10px 0;background:#fafafa;}
.small-muted {color:#666;font-size:0.9rem;}
.big-title {font-size:1.25rem;font-weight:800;}
.ok {padding:10px;border-radius:10px;background:#e8f7ee;border:1px solid #b8e2c7;}
.warn {padding:10px;border-radius:10px;background:#fff6e5;border:1px solid #f1d396;}
</style>
""", unsafe_allow_html=True)

st.title("BSNL BTS Field Data submission for Comprehensive Mtce Tender")

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
MASTER_PASSWORD = os.getenv("MASTER_PASSWORD", "1234")

if not SUPABASE_URL or not SUPABASE_KEY:
    st.error("SUPABASE_URL / SUPABASE_KEY missing in Render Environment Variables.")
    st.stop()

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

FALLBACK_OPTIONS = {
    "tower_type": ["GBT", "RTT", "RTP", "GBM"],
    "site_type": ["Building", "Shelter", "Cage", "Open"],
    "sites_having_acs": ["Yes", "No"],
    "status_of_ac": ["Working", "Faulty", "Not Available"],
    "ac_pm_cm": ["Required", "Not Required"],
    "free_cooling_available": ["Yes", "No"],
    "site_load_gt_10kw": ["Yes", "No"],
}

DISPLAY_COLUMNS = [
    "si_no", "ba_name", "oa_name", "sdca_name", "bts_type", "tower_type",
    "bts_name", "rp_id", "site_type", "sites_having_acs", "acs_count",
    "make", "ac_capacity", "status_of_ac", "ac_pm_cm",
    "free_cooling_available", "site_load_gt_10kw", "remarks",
    "sde_name", "is_completed", "last_updated_by", "last_updated_at"
]

LABELS = {
    "si_no": "SI NO",
    "ba_name": "BA Name",
    "oa_name": "OA Name",
    "sdca_name": "SDCA Name",
    "bts_type": "BTS Type",
    "tower_type": "Tower Type",
    "bts_name": "BTS Name",
    "rp_id": "RP ID",
    "site_type": "Site Type",
    "sites_having_acs": "Sites Having ACS",
    "acs_count": "ACS Count",
    "make": "Make",
    "ac_capacity": "AC Capacity",
    "status_of_ac": "Status of AC",
    "ac_pm_cm": "AC PM/CM",
    "free_cooling_available": "Free Cooling Available",
    "site_load_gt_10kw": "Site Load >10KW",
    "remarks": "Remarks",
    "sde_name": "Name of SDE CM/CFA",
    "is_completed": "Completed",
    "last_updated_by": "Last Updated By",
    "last_updated_at": "Last Updated At",
}


@st.cache_data(ttl=15)
def load_sites():
    data = supabase.table("sites").select("*").order("si_no").execute().data
    return pd.DataFrame(data)


@st.cache_data(ttl=300)
def load_options():
    try:
        data = supabase.table("dropdown_options").select("*").order("sort_order").execute().data
    except Exception:
        data = []

    options = {}
    for r in data:
        options.setdefault(r["field_name"], []).append(r["option_value"])

    for k, v in FALLBACK_OPTIONS.items():
        options.setdefault(k, v)

    return options


def is_blank(x):
    if x is None:
        return True
    try:
        if pd.isna(x):
            return True
    except Exception:
        pass
    return str(x).strip() == "" or str(x).strip().lower() == "nan"


def clean_value(x):
    return None if is_blank(x) else str(x).strip()


def safe_int(x, default=0):
    try:
        if pd.isna(x):
            return default
        return int(float(x))
    except Exception:
        return default


def selected_value(options, current):
    if is_blank(current):
        return ""
    current = str(current).strip()
    return current if current in options else ""


def compute_completed(row):
    if is_blank(row.get("tower_type")):
        return False
    if is_blank(row.get("site_type")):
        return False
    if is_blank(row.get("sites_having_acs")):
        return False
    if is_blank(row.get("free_cooling_available")):
        return False
    if is_blank(row.get("site_load_gt_10kw")):
        return False

    if str(row.get("sites_having_acs")).strip() == "No":
        return True

    ac_required = ["acs_count", "make", "ac_capacity", "status_of_ac", "ac_pm_cm"]
    return all(not is_blank(row.get(c)) for c in ac_required)


def save_site(row_id, payload):
    payload["is_completed"] = compute_completed(payload)
    payload["last_updated_by"] = payload.get("sde_name")
    payload["last_updated_at"] = datetime.now(timezone.utc).isoformat()

    supabase.table("sites").update(payload).eq("id", row_id).execute()
    st.cache_data.clear()


def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.rename(columns=LABELS).to_excel(writer, index=False, sheet_name="Master_Report")
    return output.getvalue()


options = load_options()
mode = st.radio("Open as", ["Field User", "Master Login"], horizontal=True)

if mode == "Field User":
    df_all = load_sites()

    if df_all.empty:
        st.error("No data found. Import seed_sites.csv into Supabase sites table.")
        st.stop()

    if "field_started" not in st.session_state:
        st.session_state.field_started = False

    if not st.session_state.field_started:
        sdca_list = sorted([x for x in df_all["sdca_name"].dropna().unique() if str(x).strip()])
        sdca = st.selectbox("Select SDCA", sdca_list)
        sde_name = st.text_input("Name of SDE CM/CFA")

        if st.button("Continue", disabled=not bool(sde_name.strip())):
            df_sdca = df_all[df_all["sdca_name"] == sdca].copy().reset_index(drop=True)
            pending = df_sdca[~df_sdca.apply(compute_completed, axis=1)]

            st.session_state.sdca = sdca
            st.session_state.sde_name = sde_name.strip()
            st.session_state.site_pos = int(pending.index[0]) if not pending.empty else 0
            st.session_state.field_started = True
            st.rerun()

        st.stop()

    sdca = st.session_state.sdca
    sde_name = st.session_state.sde_name

    df = df_all[df_all["sdca_name"] == sdca].copy().sort_values("si_no").reset_index(drop=True)
    total = len(df)

    if total == 0:
        st.error("No sites found for selected SDCA.")
        st.stop()

    completed_count = int(df.apply(compute_completed, axis=1).sum())

    st.markdown(
        f"<div class='small-muted'>SDCA: <b>{sdca}</b> | SDE: <b>{sde_name}</b></div>",
        unsafe_allow_html=True,
    )
    st.progress(completed_count / total, text=f"Completed {completed_count}/{total}")

    search = st.text_input("Search BTS / RP ID", placeholder="Optional")

    if search.strip():
        mask = (
            df["bts_name"].astype(str).str.contains(search, case=False, na=False)
            | df["rp_id"].astype(str).str.contains(search, case=False, na=False)
        )
        matches = df[mask]

        if not matches.empty:
            match_labels = [
                f"{safe_int(r.get('si_no'), i + 1)}. {r.get('bts_name')} - {r.get('rp_id')}"
                for i, r in matches.iterrows()
            ]
            selected_match = st.selectbox("Matching sites", match_labels)

            if st.button("Open Selected Site"):
                st.session_state.site_pos = matches.index[match_labels.index(selected_match)]
                st.rerun()
        else:
            st.warning("No matching site found.")

    pos = max(0, min(st.session_state.get("site_pos", 0), total - 1))
    row = df.iloc[pos].to_dict()

    st.markdown(
        f"""
        <div class='site-card'>
            <div class='big-title'>Site {pos + 1} of {total}</div>
            <b>BTS:</b> {row.get("bts_name", "")}<br>
            <b>RP ID:</b> {row.get("rp_id", "")}<br>
            <span class='small-muted'>
            BA: {row.get("ba_name", "")} | OA: {row.get("oa_name", "")} | BTS Type: {row.get("bts_type", "")}
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    tower_type = st.selectbox(
        "Tower Type *",
        [""] + options["tower_type"],
        index=([""] + options["tower_type"]).index(selected_value(options["tower_type"], row.get("tower_type"))),
        key=f"tower_{row['id']}",
    )

    site_type = st.selectbox(
        "Site Type *",
        [""] + options["site_type"],
        index=([""] + options["site_type"]).index(selected_value(options["site_type"], row.get("site_type"))),
        key=f"site_type_{row['id']}",
    )

    sites_having_acs = st.selectbox(
        "Sites Having ACS *",
        [""] + options["sites_having_acs"],
        index=([""] + options["sites_having_acs"]).index(selected_value(options["sites_having_acs"], row.get("sites_having_acs"))),
        key=f"acs_yesno_{row['id']}",
    )

    if sites_having_acs == "No":
        st.info("ACS not available. AC fields are auto-set as N/A.")
        acs_count = 0
        make = "N/A"
        ac_capacity = "N/A"
        status_of_ac = "Not Available"
        ac_pm_cm = "Not Required"
    else:
        acs_count = st.number_input(
            "ACS Count *",
            min_value=0,
            value=safe_int(row.get("acs_count")),
            step=1,
            key=f"acs_count_{row['id']}",
        )

        make = st.text_input(
            "Make *",
            value="" if is_blank(row.get("make")) else str(row.get("make")),
            key=f"make_{row['id']}",
        )

        ac_capacity = st.text_input(
            "AC Capacity *",
            value="" if is_blank(row.get("ac_capacity")) else str(row.get("ac_capacity")),
            key=f"capacity_{row['id']}",
        )

        status_of_ac = st.selectbox(
            "Status of AC *",
            [""] + options["status_of_ac"],
            index=([""] + options["status_of_ac"]).index(selected_value(options["status_of_ac"], row.get("status_of_ac"))),
            key=f"status_{row['id']}",
        )

        ac_pm_cm = st.selectbox(
            "AC PM/CM *",
            [""] + options["ac_pm_cm"],
            index=([""] + options["ac_pm_cm"]).index(selected_value(options["ac_pm_cm"], row.get("ac_pm_cm"))),
            key=f"pmcm_{row['id']}",
        )

    free_cooling_available = st.selectbox(
        "Free Cooling Available *",
        [""] + options["free_cooling_available"],
        index=([""] + options["free_cooling_available"]).index(selected_value(options["free_cooling_available"], row.get("free_cooling_available"))),
        key=f"freecool_{row['id']}",
    )

    site_load_gt_10kw = st.selectbox(
        "Site Load >10KW *",
        [""] + options["site_load_gt_10kw"],
        index=([""] + options["site_load_gt_10kw"]).index(selected_value(options["site_load_gt_10kw"], row.get("site_load_gt_10kw"))),
        key=f"load_{row['id']}",
    )

    remarks = st.text_area(
        "Remarks",
        value="" if is_blank(row.get("remarks")) else str(row.get("remarks")),
        height=90,
        key=f"remarks_{row['id']}",
    )

    payload = {
        "tower_type": clean_value(tower_type),
        "site_type": clean_value(site_type),
        "sites_having_acs": clean_value(sites_having_acs),
        "acs_count": acs_count,
        "make": clean_value(make),
        "ac_capacity": clean_value(ac_capacity),
        "status_of_ac": clean_value(status_of_ac),
        "ac_pm_cm": clean_value(ac_pm_cm),
        "free_cooling_available": clean_value(free_cooling_available),
        "site_load_gt_10kw": clean_value(site_load_gt_10kw),
        "remarks": clean_value(remarks),
        "sde_name": sde_name,
    }

    current_complete = compute_completed(payload)

    if current_complete:
        st.markdown("<div class='ok'>This site is ready to save.</div>", unsafe_allow_html=True)
    else:
        st.markdown("<div class='warn'>Fill all fields marked * to enable Save & Next.</div>", unsafe_allow_html=True)

    c1, c2 = st.columns(2)

    with c1:
        if st.button("← Previous"):
            st.session_state.site_pos = max(0, pos - 1)
            st.rerun()

    with c2:
        if st.button("Save & Next", disabled=not current_complete):
            save_site(row["id"], payload)
            if pos < total - 1:
                st.session_state.site_pos = pos + 1
            st.success("Saved successfully.")
            st.rerun()

    c1, c2, c3 = st.columns(3)

    with c1:
        if st.button("First"):
            st.session_state.site_pos = 0
            st.rerun()

    with c2:
        if st.button("Pending"):
            fresh = load_sites()
            fresh_sdca = fresh[fresh["sdca_name"] == sdca].copy().sort_values("si_no").reset_index(drop=True)
            pending = fresh_sdca[~fresh_sdca.apply(compute_completed, axis=1)]

            if not pending.empty:
                st.session_state.site_pos = int(pending.index[0])
            else:
                st.session_state.site_pos = pos

            st.rerun()

    with c3:
        if st.button("Logout"):
            for k in ["field_started", "sdca", "sde_name", "site_pos"]:
                st.session_state.pop(k, None)
            st.rerun()

    if completed_count == total:
        st.success("All sites in this SDCA are completed.")

        if st.button("Submit Final Report"):
            now = datetime.now(timezone.utc).isoformat()
            supabase.table("sites").update({
                "last_updated_by": sde_name,
                "last_updated_at": now
            }).eq("sdca_name", sdca).execute()
            st.success("Final report submitted successfully.")

else:
    password = st.text_input("Master Password", type="password")

    if password != MASTER_PASSWORD:
        st.stop()

    df_all = load_sites()

    if df_all.empty:
        st.error("No site data found.")
        st.stop()

    st.subheader("Master Combined Report")

    summary = df_all.groupby("sdca_name", dropna=False).agg(
        total_sites=("rp_id", "count"),
        completed=("is_completed", "sum"),
    ).reset_index()

    summary["completed"] = summary["completed"].fillna(0).astype(int)
    summary["pending"] = summary["total_sites"] - summary["completed"]
    summary["completion_%"] = (summary["completed"] / summary["total_sites"] * 100).round(1)

    st.dataframe(summary, use_container_width=True, hide_index=True)

    sdca_filter = st.multiselect("Filter SDCA", sorted(df_all["sdca_name"].dropna().unique()))
    view = df_all.copy()

    if sdca_filter:
        view = view[view["sdca_name"].isin(sdca_filter)]

    search = st.text_input("Search BTS / RP ID / SDE")

    if search.strip():
        m = pd.Series(False, index=view.index)
        for col in ["bts_name", "rp_id", "sde_name", "sdca_name"]:
            if col in view.columns:
                m = m | view[col].astype(str).str.contains(search, case=False, na=False)
        view = view[m]

    cols = [c for c in DISPLAY_COLUMNS if c in view.columns]
    st.dataframe(view[cols].rename(columns=LABELS), use_container_width=True, hide_index=True)

    st.download_button(
        "Download Combined Excel",
        data=to_excel(view[cols]),
        file_name="BTS_Master_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
