import re
from datetime import datetime, timedelta

import pandas as pd
import requests
import streamlit as st

# ──────────────────────────────────────────────────────────────
# Page config
# ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Access Review Dashboard",
    page_icon="🔐",
    layout="wide",
)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
REMINDER_DAYS_BEFORE_END = 7  # assumed Entra ID default reminder lead-time


# ──────────────────────────────────────────────────────────────
# Authentication
# ──────────────────────────────────────────────────────────────
def get_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    resp = requests.post(
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
        data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


# ──────────────────────────────────────────────────────────────
# Graph API helpers
# ──────────────────────────────────────────────────────────────
def _graph_get_all(token: str, url: str, params: dict = None) -> list:
    """Follow @odata.nextLink pagination and return all items."""
    headers = {"Authorization": f"Bearer {token}"}
    results = []
    while url:
        resp = requests.get(url, headers=headers, params=params, timeout=30)
        resp.raise_for_status()
        body = resp.json()
        results.extend(body.get("value", []))
        url = body.get("@odata.nextLink")
        params = None
    return results


@st.cache_data(ttl=300, show_spinner=False)
def _fetch_definitions(token: str) -> list:
    return _graph_get_all(token, f"{GRAPH_BASE}/identityGovernance/accessReviews/definitions")


@st.cache_data(ttl=300, show_spinner=False)
def _fetch_instances(token: str, def_id: str) -> list:
    return _graph_get_all(
        token,
        f"{GRAPH_BASE}/identityGovernance/accessReviews/definitions/{def_id}/instances",
    )


@st.cache_data(ttl=300, show_spinner=False)
def _fetch_decisions(token: str, def_id: str, inst_id: str) -> list:
    return _graph_get_all(
        token,
        f"{GRAPH_BASE}/identityGovernance/accessReviews/definitions/{def_id}/instances/{inst_id}/decisions",
    )


# ──────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────
def _parse_dt(s) -> datetime | None:
    if not s:
        return None
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except (ValueError, AttributeError):
        return None


def _fmt_date(dt: datetime | None) -> str:
    return dt.strftime("%Y-%m-%d") if dt else "—"


def _is_sponsor(name: str) -> bool:
    return "GuestReview" in name and "No Sponsor" not in name


def _extract_sponsor(name: str) -> str:
    """Extract sponsor name from 'GuestReview - SponsorName (Month)'."""
    m = re.match(r"GuestReview\s*-\s*(.+?)\s*\(", name)
    return m.group(1).strip() if m else name


# ──────────────────────────────────────────────────────────────
# Data loading
# ──────────────────────────────────────────────────────────────
_INST_COLS = [
    "def_id", "def_name", "inst_id", "status", "start_dt", "end_dt",
    "apply_dt", "reminder_dt", "reminders_enabled", "auto_apply", "type", "sponsor",
]
_DEC_COLS = [
    "def_id", "def_name", "inst_id", "decision", "principal", "reviewed_by", "type", "sponsor",
]


def load_data(token: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Load all GuestReview definitions and their instances.
    For InProgress instances also load decisions.
    Returns (instances_df, decisions_df).
    """
    try:
        all_defs = _fetch_definitions(token)
    except requests.HTTPError as exc:
        st.error(f"Failed to fetch definitions: {exc}")
        return pd.DataFrame(columns=_INST_COLS), pd.DataFrame(columns=_DEC_COLS)

    guest_defs = [d for d in all_defs if "GuestReview" in d.get("displayName", "")]
    if not guest_defs:
        return pd.DataFrame(columns=_INST_COLS), pd.DataFrame(columns=_DEC_COLS)

    inst_rows, dec_rows = [], []
    n = len(guest_defs)
    prog = st.progress(0.0, text="Loading access reviews…")

    for i, defn in enumerate(guest_defs):
        prog.progress((i + 1) / n, text=defn.get("displayName", "?"))
        name = defn.get("displayName", "")
        settings = defn.get("settings") or {}
        auto_apply = settings.get("autoApplyDecisionsEnabled", False)
        reminders = settings.get("reminderNotificationsEnabled", False)
        review_type = "Sponsor" if _is_sponsor(name) else "Self-Review"
        sponsor = _extract_sponsor(name) if _is_sponsor(name) else "No Sponsor"

        try:
            instances = _fetch_instances(token, defn["id"])
        except requests.HTTPError:
            instances = []

        for inst in instances:
            start = _parse_dt(inst.get("startDateTime"))
            end = _parse_dt(inst.get("endDateTime"))
            inst_rows.append({
                "def_id": defn["id"],
                "def_name": name,
                "inst_id": inst["id"],
                "status": inst.get("status", ""),
                "start_dt": start,
                "end_dt": end,
                "apply_dt": end if auto_apply else None,
                "reminder_dt": (end - timedelta(days=REMINDER_DAYS_BEFORE_END)) if reminders and end else None,
                "reminders_enabled": reminders,
                "auto_apply": auto_apply,
                "type": review_type,
                "sponsor": sponsor,
            })

            # Only load decisions for currently active instances
            if inst.get("status") == "InProgress":
                try:
                    for dec in _fetch_decisions(token, defn["id"], inst["id"]):
                        dec_rows.append({
                            "def_id": defn["id"],
                            "def_name": name,
                            "inst_id": inst["id"],
                            "decision": dec.get("decision", ""),
                            "principal": (dec.get("principal") or {}).get("displayName", ""),
                            "reviewed_by": (dec.get("reviewedBy") or {}).get("displayName", ""),
                            "type": review_type,
                            "sponsor": sponsor,
                        })
                except requests.HTTPError:
                    pass  # instance may not have decisions yet

    prog.empty()

    inst_df = pd.DataFrame(inst_rows, columns=_INST_COLS) if inst_rows else pd.DataFrame(columns=_INST_COLS)
    dec_df = pd.DataFrame(dec_rows, columns=_DEC_COLS) if dec_rows else pd.DataFrame(columns=_DEC_COLS)
    return inst_df, dec_df


# ──────────────────────────────────────────────────────────────
# Tab: Overview
# ──────────────────────────────────────────────────────────────
def tab_overview(inst_df: pd.DataFrame, dec_df: pd.DataFrame) -> None:
    st.header("Overview")

    if inst_df.empty:
        st.info("No GuestReview access reviews found.")
        return

    active_df = inst_df[inst_df["status"] == "InProgress"]
    pending_count = int((dec_df["decision"].isin(["NotReviewed", "NotNotified"])).sum()) if not dec_df.empty else 0

    # KPI metrics
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Active Reviews", len(active_df))
    c2.metric("Pending Decisions", pending_count)
    c3.metric("Sponsor Reviews (active)", int((active_df["type"] == "Sponsor").sum()))
    c4.metric("Self-Reviews (active)", int((active_df["type"] == "Self-Review").sum()))

    st.divider()

    # Active reviews table with key dates
    st.subheader("Active Reviews")
    if active_df.empty:
        st.info("No reviews are currently in progress.")
    else:
        rows = []
        for _, r in active_df.iterrows():
            rows.append({
                "Review Name": r["def_name"],
                "Type": r["type"],
                "Sponsor / Owner": r["sponsor"],
                "Status": r["status"],
                "Start Date": _fmt_date(r["start_dt"]),
                "End Date": _fmt_date(r["end_dt"]),
                "Results Applied": _fmt_date(r["apply_dt"]) if r["apply_dt"] else "Manual",
                "Reminder Date": (
                    _fmt_date(r["reminder_dt"]) if r["reminder_dt"]
                    else ("Disabled" if not r["reminders_enabled"] else "—")
                ),
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    st.divider()

    # All instances by status
    st.subheader("All Instances by Status")
    by_status = inst_df.groupby("status").size().reset_index(name="Count")
    st.dataframe(by_status, use_container_width=True, hide_index=True)


# ──────────────────────────────────────────────────────────────
# Tab: Sponsor-Based Reviews
# ──────────────────────────────────────────────────────────────
def _decision_summary(dec_df: pd.DataFrame, id_col: str) -> pd.DataFrame:
    """Return per-review counts for Approve, Deny, and pending decisions."""
    if dec_df.empty:
        return pd.DataFrame(columns=[id_col, "Not Reviewed", "Approved", "Denied"])

    def _count(decision_values):
        return (
            dec_df[dec_df["decision"].isin(decision_values)]
            .groupby(id_col)
            .size()
            .reset_index(name="_count")
        )

    pending_df = _count(["NotReviewed", "NotNotified"]).rename(columns={"_count": "Not Reviewed"})
    approved_df = _count(["Approve"]).rename(columns={"_count": "Approved"})
    denied_df = _count(["Deny"]).rename(columns={"_count": "Denied"})

    result = pending_df.merge(approved_df, on=id_col, how="outer")
    result = result.merge(denied_df, on=id_col, how="outer")
    result[["Not Reviewed", "Approved", "Denied"]] = (
        result[["Not Reviewed", "Approved", "Denied"]].fillna(0).astype(int)
    )
    return result


def tab_sponsor_reviews(inst_df: pd.DataFrame, dec_df: pd.DataFrame) -> None:
    st.header("Sponsor-Based Access Reviews")
    st.caption("Pattern: **GuestReview - {SponsorName} ({Month})**")

    sponsor_inst = inst_df[(inst_df["type"] == "Sponsor") & (inst_df["status"] == "InProgress")]
    sponsor_dec = dec_df[dec_df["type"] == "Sponsor"] if not dec_df.empty else pd.DataFrame(columns=_DEC_COLS)

    # Aggregate metrics
    not_reviewed = int(sponsor_dec["decision"].isin(["NotReviewed", "NotNotified"]).sum())
    approved = int((sponsor_dec["decision"] == "Approve").sum())
    denied = int((sponsor_dec["decision"] == "Deny").sum())

    c1, c2, c3 = st.columns(3)
    c1.metric("🟡 Not Reviewed", not_reviewed)
    c2.metric("🟢 Approved", approved)
    c3.metric("🔴 Denied", denied)

    st.divider()

    st.subheader("Reviews with Pending Decisions (by Sponsor)")

    if sponsor_inst.empty:
        st.info("No active sponsor-based reviews found.")
        return

    summary_counts = _decision_summary(sponsor_dec, "def_id")
    base = (
        sponsor_inst[["def_id", "def_name", "sponsor", "end_dt"]]
        .drop_duplicates("def_id")
        .merge(summary_counts, on="def_id", how="left")
    )
    base[["Not Reviewed", "Approved", "Denied"]] = (
        base[["Not Reviewed", "Approved", "Denied"]].fillna(0).astype(int)
    )
    base = base.sort_values("Not Reviewed", ascending=False)

    display = base.rename(columns={"def_name": "Review Name", "sponsor": "Sponsor", "end_dt": "End Date"})
    display["End Date"] = display["End Date"].apply(_fmt_date)
    display = display[["Review Name", "Sponsor", "End Date", "Not Reviewed", "Approved", "Denied"]]

    with_pending = display[display["Not Reviewed"] > 0]
    without_pending = display[display["Not Reviewed"] == 0]

    if not with_pending.empty:
        st.dataframe(with_pending, use_container_width=True, hide_index=True)
    if not without_pending.empty:
        with st.expander(f"Reviews with no pending decisions ({len(without_pending)})"):
            st.dataframe(without_pending, use_container_width=True, hide_index=True)


# ──────────────────────────────────────────────────────────────
# Tab: Self-Reviews
# ──────────────────────────────────────────────────────────────
def tab_self_reviews(inst_df: pd.DataFrame, dec_df: pd.DataFrame) -> None:
    st.header("Self-Review Access Reviews")
    st.caption("Pattern: **GuestReview - No Sponsor ({Month})**")

    self_inst = inst_df[(inst_df["type"] == "Self-Review") & (inst_df["status"] == "InProgress")]
    self_dec = dec_df[dec_df["type"] == "Self-Review"] if not dec_df.empty else pd.DataFrame(columns=_DEC_COLS)

    # Aggregate metrics
    pending = int(self_dec["decision"].isin(["NotReviewed", "NotNotified"]).sum())
    approved = int((self_dec["decision"] == "Approve").sum())
    denied = int((self_dec["decision"] == "Deny").sum())

    c1, c2, c3 = st.columns(3)
    c1.metric("🟡 Pending", pending)
    c2.metric("🟢 Approved", approved)
    c3.metric("🔴 Denied", denied)

    st.divider()

    st.subheader("Active Self-Reviews")

    if self_inst.empty:
        st.info("No active self-reviews found.")
        return

    summary_counts = _decision_summary(self_dec, "def_id")
    base = (
        self_inst[["def_id", "def_name", "end_dt"]]
        .drop_duplicates("def_id")
        .merge(summary_counts, on="def_id", how="left")
    )
    base[["Not Reviewed", "Approved", "Denied"]] = (
        base[["Not Reviewed", "Approved", "Denied"]].fillna(0).astype(int)
    )
    base = base.sort_values("Not Reviewed", ascending=False)

    display = base.rename(columns={"def_name": "Review Name", "end_dt": "End Date"})
    display["End Date"] = display["End Date"].apply(_fmt_date)
    display = display[["Review Name", "End Date", "Not Reviewed", "Approved", "Denied"]]
    st.dataframe(display, use_container_width=True, hide_index=True)


# ──────────────────────────────────────────────────────────────
# Sidebar
# ──────────────────────────────────────────────────────────────
def render_sidebar() -> str | None:
    with st.sidebar:
        st.title("🔐 Configuration")
        st.divider()

        tenant_id = st.text_input("Tenant ID", key="tenant_id")
        client_id = st.text_input("Client ID", key="client_id")
        client_secret = st.text_input("Client Secret", key="client_secret", type="password")

        st.divider()
        connect = st.button("Connect", use_container_width=True, type="primary")
        if st.button("Refresh Data", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

        if connect:
            if not all([tenant_id, client_id, client_secret]):
                st.error("All credential fields are required.")
            else:
                with st.spinner("Authenticating…"):
                    try:
                        token = get_token(tenant_id, client_id, client_secret)
                        st.session_state["token"] = token
                        st.success("Connected ✓")
                    except requests.HTTPError as exc:
                        st.error(f"Authentication failed: {exc}")
                        st.session_state.pop("token", None)
                    except Exception as exc:
                        st.error(f"Unexpected error: {exc}")
                        st.session_state.pop("token", None)

        if "token" in st.session_state:
            st.success("Status: Connected")
        else:
            st.warning("Not connected")

    return st.session_state.get("token")


# ──────────────────────────────────────────────────────────────
# Entry point
# ──────────────────────────────────────────────────────────────
def main() -> None:
    st.title("Access Review Dashboard")
    st.caption("Azure Entra Identity Governance — GuestReview KPIs")

    token = render_sidebar()

    if not token:
        st.info("Enter your service principal credentials in the sidebar to connect.")
        st.markdown(
            "**Required Graph API permission:** `AccessReview.Read.All` (Application)"
        )
        return

    inst_df, dec_df = load_data(token)

    if inst_df.empty:
        st.warning(
            "No GuestReview access reviews found. "
            "Verify the service principal has `AccessReview.Read.All` permission."
        )
        return

    tab1, tab2, tab3 = st.tabs(["📊 Overview", "👤 Sponsor Reviews", "🧑 Self-Reviews"])
    with tab1:
        tab_overview(inst_df, dec_df)
    with tab2:
        tab_sponsor_reviews(inst_df, dec_df)
    with tab3:
        tab_self_reviews(inst_df, dec_df)


if __name__ == "__main__":
    main()
