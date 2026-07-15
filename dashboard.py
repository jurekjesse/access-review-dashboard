import logging
import re
import time
import pickle
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

# ──────────────────────────────────────────────────────────────
# Logging setup
# ──────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format='%(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# ──────────────────────────────────────────────────────────────
# Page config
# ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Access Review Dashboard",
    page_icon="🔐",
    layout="wide",
)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
REMINDER_DAYS_BEFORE_END = 25  # assumed Entra ID default reminder lead-time
API_RATE_LIMIT_SLEEP = 2  # Sleep 2 seconds between definition fetches to avoid 429 errors

# Cache configuration
CACHE_DIR = Path.home() / ".access_review_cache"
CACHE_FILE = CACHE_DIR / "access_review_data.pkl"

# Fallback cache in project directory (more reliable in codespace)
PROJECT_CACHE_DIR = Path(__file__).parent / ".cache"
PROJECT_CACHE_FILE = PROJECT_CACHE_DIR / "access_review_data.pkl"


def _get_cache_file() -> Path:
    """Get the best available cache file location."""
    # Try home directory first
    try:
        CACHE_DIR.mkdir(exist_ok=True, parents=True)
        CACHE_FILE.touch()  # Test write access
        logger.debug(f"Using home cache: {CACHE_FILE}")
        return CACHE_FILE
    except (OSError, PermissionError):
        pass
    
    # Fallback to project directory
    try:
        PROJECT_CACHE_DIR.mkdir(exist_ok=True, parents=True)
        PROJECT_CACHE_FILE.touch()  # Test write access
        logger.debug(f"Using project cache: {PROJECT_CACHE_FILE}")
        return PROJECT_CACHE_FILE
    except (OSError, PermissionError) as e:
        logger.error(f"No writable cache location found: {e}")
        raise


# ──────────────────────────────────────────────────────────────
# Cache management
# ──────────────────────────────────────────────────────────────
def _save_cache(inst_df: pd.DataFrame, dec_df: pd.DataFrame) -> bool:
    """Save dataframes to persistent cache file. Returns True if successful."""
    try:
        cache_file = _get_cache_file()
        with open(cache_file, "wb") as f:
            pickle.dump({"instances": inst_df, "decisions": dec_df, "timestamp": datetime.now()}, f)
        logger.info(f"✓ Cache saved: {len(inst_df)} instances, {len(dec_df)} decisions")
        return True
    except Exception as e:
        logger.error(f"✗ Failed to save cache: {e}")
        st.error(f"⚠️ Cache save failed: {e}")
        return False


def _load_cache() -> tuple[pd.DataFrame, pd.DataFrame] | None:
    """Load dataframes from persistent cache file."""
    try:
        # Try project cache first, then home cache
        for cache_file in [PROJECT_CACHE_FILE, CACHE_FILE]:
            if cache_file.exists():
                logger.info(f"Attempting to load from: {cache_file}")
                with open(cache_file, "rb") as f:
                    data = pickle.load(f)
                timestamp = data.get("timestamp", "unknown")
                inst_df = data["instances"]
                dec_df = data["decisions"]
                # Validate required columns exist
                required_inst_cols = {"def_id", "def_name", "inst_id", "status", "start_dt", "end_dt"}
                if not required_inst_cols.issubset(inst_df.columns):
                    missing = required_inst_cols - set(inst_df.columns)
                    logger.warning(f"Stale cache (missing columns: {missing}) — deleting")
                    cache_file.unlink(missing_ok=True)
                    continue
                logger.info(f"✓ Cache loaded from {timestamp}: {len(inst_df)} instances, {len(dec_df)} decisions")
                return inst_df, dec_df
    except Exception as e:
        logger.error(f"✗ Failed to load cache: {e}", exc_info=True)
    
    logger.warning("No cache file found")
    return None


def _clear_cache() -> None:
    """Clear all persistent cache files."""
    cleared = False
    for cache_file in [CACHE_FILE, PROJECT_CACHE_FILE]:
        try:
            if cache_file.exists():
                cache_file.unlink()
                logger.info(f"✓ Cache cleared: {cache_file}")
                cleared = True
        except Exception as e:
            logger.error(f"✗ Failed to clear cache: {e}")
    if cleared:
        st.success("Cache cleared")


def _get_cache_info() -> str:
    """Get info about current cache."""
    for cache_file in [PROJECT_CACHE_FILE, CACHE_FILE]:
        if cache_file.exists():
            try:
                mtime = datetime.fromtimestamp(cache_file.stat().st_mtime)
                return f"✓ Cached ({mtime.strftime('%Y-%m-%d %H:%M:%S')})"
            except Exception:
                pass
    return "No cache"


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
    token = resp.json()["access_token"]
    return token


# ──────────────────────────────────────────────────────────────
# Graph API helpers
# ──────────────────────────────────────────────────────────────
def _graph_get_all(token: str, url: str, params: dict = None) -> list:
    """Follow @odata.nextLink pagination and return all items."""
    headers = {"Authorization": f"Bearer {token}"}
    results = []
    page_count = 0
    while url:
        page_count += 1
        try:
            resp = requests.get(url, headers=headers, params=params, timeout=30)
            resp.raise_for_status()
            body = resp.json()
            page_results = body.get("value", [])
            results.extend(page_results)
            url = body.get("@odata.nextLink")
            params = None
        except requests.HTTPError as e:
            logger.error(f"HTTP Error: {e.response.status_code} - {e.response.text}")
            raise
        except Exception as e:
            logger.error(f"Unexpected error: {e}")
            raise
    return results


@st.cache_data(ttl=300, show_spinner=False)
def _fetch_definitions(token: str) -> list:
    result = _graph_get_all(token, f"{GRAPH_BASE}/identityGovernance/accessReviews/definitions")
    logger.info(f"Fetched {len(result)} access review definitions")
    return result


@st.cache_data(ttl=300, show_spinner=False)
def _fetch_instances(token: str, def_id: str) -> list:
    result = _graph_get_all(
        token,
        f"{GRAPH_BASE}/identityGovernance/accessReviews/definitions/{def_id}/instances",
    )
    return result


@st.cache_data(ttl=300, show_spinner=False)
def _fetch_decisions(token: str, def_id: str, inst_id: str) -> list:
    result = _graph_get_all(
        token,
        f"{GRAPH_BASE}/identityGovernance/accessReviews/definitions/{def_id}/instances/{inst_id}/decisions",
    )
    return result


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
    return "Guest Recertification" in name and not name.endswith("No Sponsor")


def _extract_sponsor(name: str) -> str:
    """Extract sponsor name from 'Guest Recertification - SponsorName (Month)'."""
    m = re.match(r"Guest Recertification\s*-\s*(.+?)\s*\(", name)
    return m.group(1).strip() if m else name


# ──────────────────────────────────────────────────────────────
# Data loading
# ──────────────────────────────────────────────────────────────
_INST_COLS = [
    "def_id", "def_name", "inst_id", "status", "start_dt", "end_dt",
    "apply_dt", "reminder_dt", "reminders_enabled", "auto_apply", "type", "sponsor",
]
_DEC_COLS = [
    "def_id", "def_name", "inst_id", "decision", "principal", "reviewed_by", "justification", "type", "sponsor",
]


def load_data(token: str, use_cache: bool = True) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Load all Guest Recertification definitions and their instances.
    For InProgress instances also load decisions.
    Returns (instances_df, decisions_df).
    """
    # Try to load from cache first
    if use_cache:
        logger.info("Attempting to load from cache...")
        cached_data = _load_cache()
        if cached_data:
            inst_df, dec_df = cached_data
            logger.info(f"✓ Using cached data: {len(inst_df)} instances, {len(dec_df)} decisions")
            return inst_df, dec_df
        else:
            logger.warning("Cache load failed, fetching from API...")
    else:
        logger.info("Cache bypass requested, fetching from API...")
    
    # Fetch from API if cache not available
    try:
        logger.info("Fetching from Graph API...")
        all_defs = _fetch_definitions(token)
        logger.info(f"Fetched {len(all_defs)} access review definitions")
    except requests.HTTPError as exc:
        logger.error(f"Failed to fetch definitions: {exc}")
        logger.error(f"Response status: {exc.response.status_code if exc.response else 'N/A'}")
        logger.error(f"Response text: {exc.response.text if exc.response else 'N/A'}")
        st.error(f"Failed to fetch definitions: {exc}")
        return pd.DataFrame(columns=_INST_COLS), pd.DataFrame(columns=_DEC_COLS)
    except Exception as exc:
        logger.error(f"Unexpected error fetching definitions: {exc}", exc_info=True)
        st.error(f"Unexpected error: {exc}")
        return pd.DataFrame(columns=_INST_COLS), pd.DataFrame(columns=_DEC_COLS)

    guest_defs = [d for d in all_defs if "Guest Recertification" in d.get("displayName", "")]
    logger.info(f"Found {len(guest_defs)} Guest Recertification access reviews")
    
    if not guest_defs:
        logger.warning("No Guest Recertification access reviews found.")
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
        except requests.HTTPError as e:
            logger.error(f"Failed to fetch instances for {name}: {e}")
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
                    decisions = _fetch_decisions(token, defn["id"], inst["id"])
                    for dec in decisions:
                        dec_rows.append({
                            "def_id": defn["id"],
                            "def_name": name,
                            "inst_id": inst["id"],
                            "decision": dec.get("decision", ""),
                            "principal": (dec.get("principal") or {}).get("displayName", ""),
                            "reviewed_by": (dec.get("reviewedBy") or {}).get("displayName", ""),
                            "justification": dec.get("justification", ""),
                            "type": review_type,
                            "sponsor": sponsor,
                        })
                except requests.HTTPError as e:
                    logger.warning(f"Failed to fetch decisions for instance {inst['id']}: {e}")
                except Exception as e:
                    logger.warning(f"Unexpected error fetching decisions: {e}")

        # Rate limit: sleep between definition fetches to avoid 429 errors
        if i < n - 1:  # Don't sleep after the last definition
            time.sleep(API_RATE_LIMIT_SLEEP)

    prog.empty()

    inst_df = pd.DataFrame(inst_rows, columns=_INST_COLS) if inst_rows else pd.DataFrame(columns=_INST_COLS)
    dec_df = pd.DataFrame(dec_rows, columns=_DEC_COLS) if dec_rows else pd.DataFrame(columns=_DEC_COLS)
    logger.info(f"Loaded {len(inst_df)} instances, {len(dec_df)} decisions")
    
    # Save to cache (critical for avoiding refetch on error)
    cache_saved = _save_cache(inst_df, dec_df)
    if not cache_saved:
        logger.warning("⚠️ Cache save failed - data will need to be refetched if the dashboard crashes")
    
    return inst_df, dec_df


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
    st.caption("Pattern: **Guest Recertification - {SponsorName}**")

    sponsor_inst = inst_df[(inst_df["type"] == "Sponsor") & (inst_df["status"] == "InProgress")]
    sponsor_dec = dec_df[dec_df["type"] == "Sponsor"] if not dec_df.empty else pd.DataFrame(columns=_DEC_COLS)

    # Aggregate metrics
    not_reviewed = int(sponsor_dec["decision"].isin(["NotReviewed", "NotNotified"]).sum())
    approved = int((sponsor_dec["decision"] == "Approve").sum())
    denied = int((sponsor_dec["decision"] == "Deny").sum())
    
    # Calculate "finished" reviews (0 pending decisions)
    if not sponsor_inst.empty:
        pending_by_review = sponsor_dec[sponsor_dec["decision"].isin(["NotReviewed", "NotNotified"])].groupby("def_id").size()
        finished_count = len(sponsor_inst) - len(pending_by_review)
        finished_pct = (finished_count / len(sponsor_inst) * 100) if len(sponsor_inst) > 0 else 0
    else:
        finished_pct = 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🟡 Pending Decisions", not_reviewed)
    c2.metric("🟢 Approved", approved)
    c3.metric("🔴 Denied", denied)
    c4.metric("✅ Finished Reviews", f"{finished_pct:.0f}%")

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

    # Calculate remaining days
    today = datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0)
    base["remaining_days"] = base["end_dt"].apply(
        lambda dt: (dt.replace(hour=0, minute=0, second=0, microsecond=0, tzinfo=None) - today.replace(tzinfo=None)).days if dt else None
    )
    
    display = base.copy()
    display["End Date"] = display.apply(
        lambda r: f"{_fmt_date(r['end_dt'])} ({r['remaining_days']}d)" if r['remaining_days'] is not None else _fmt_date(r['end_dt']),
        axis=1
    )
    display = display.rename(columns={"def_name": "Review Name"})
    display = display[["Review Name", "End Date", "Not Reviewed", "Approved", "Denied"]]

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
    st.caption("Pattern: **Guest Recertification - No Sponsor**")

    self_inst = inst_df[(inst_df["type"] == "Self-Review") & (inst_df["status"] == "InProgress")]
    self_dec = dec_df[dec_df["type"] == "Self-Review"] if not dec_df.empty else pd.DataFrame(columns=_DEC_COLS)

    # Aggregate metrics
    pending = int(self_dec["decision"].isin(["NotReviewed", "NotNotified"]).sum())
    approved = int((self_dec["decision"] == "Approve").sum())
    denied = int((self_dec["decision"] == "Deny").sum())

    c1, c2, c3 = st.columns(3)
    c1.metric("🟡 Pending Decisions", pending)
    c2.metric("🟢 Approved", approved)
    c3.metric("🔴 Denied", denied)

    st.divider()

    st.subheader("Active Self-Reviews")

    if self_inst.empty:
        st.info("No active self-reviews found.")
        return

    # Prepare data for display
    display_rows = []
    for _, dec in self_dec.iterrows():
        display_rows.append({
            "Guest": dec["principal"],
            "Decision": dec["decision"],
            "Justification": dec["justification"] if dec["justification"] else "—",
        })
    
    if display_rows:
        # Calculate remaining days for the review
        if not self_inst.empty:
            end_date = self_inst.iloc[0]["end_dt"]
            today = datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0)
            remaining = (end_date.replace(hour=0, minute=0, second=0, microsecond=0, tzinfo=None) - today.replace(tzinfo=None)).days if end_date else None
            st.write(f"**End Date:** {_fmt_date(end_date)} ({remaining}d remaining)" if remaining is not None else f"**End Date:** {_fmt_date(end_date)}")
        
        st.dataframe(pd.DataFrame(display_rows), use_container_width=True, hide_index=True)
    else:
        st.info("No decisions recorded yet.")


# ──────────────────────────────────────────────────────────────
# Sidebar
# ──────────────────────────────────────────────────────────────
def render_sidebar() -> str | None:
    logger.info("→ render_sidebar() START")
    with st.sidebar:
        st.title("🔐 Configuration")
        st.divider()

        tenant_id = st.text_input("Tenant ID", key="tenant_id")
        client_id = st.text_input("Client ID", key="client_id")
        client_secret = st.text_input("Client Secret", key="client_secret", type="password")

        st.divider()
        connect = st.button("Connect", use_container_width=True, type="primary")
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Use Cache", use_container_width=True, key="btn_use_cache"):
                st.session_state["force_refresh"] = False
                logger.info("User clicked 'Use Cache' button")
                st.rerun()
        with col2:
            if st.button("Refresh API", use_container_width=True, key="btn_refresh_api"):
                logger.info("User clicked 'Refresh API' button - clearing cache")
                _clear_cache()
                st.session_state["force_refresh"] = True
                st.rerun()
        
        # Show cache status
        cache_info = _get_cache_info()
        st.caption(f"💾 {cache_info}")

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
                        logger.error(f"Authentication failed: {exc}")
                        st.error(f"Authentication failed: {exc}")
                        st.session_state.pop("token", None)
                    except Exception as exc:
                        logger.error(f"Unexpected error during authentication: {exc}", exc_info=True)
                        st.error(f"Unexpected error: {exc}")
                        st.session_state.pop("token", None)

        if "token" in st.session_state:
            st.success("Status: Connected")
        else:
            st.warning("Not connected")

    logger.info(f"→ render_sidebar() END, returning token: {st.session_state.get('token') is not None}")
    return st.session_state.get("token")


# ──────────────────────────────────────────────────────────────
# Entry point
# ──────────────────────────────────────────────────────────────
def main() -> None:
    try:
        st.title("Access Review Dashboard")
        st.caption("Azure Entra Identity Governance — GuestReview KPIs")
        
        # Initialize cache system
        try:
            cache_file = _get_cache_file()
            logger.info(f"Cache system initialized: {cache_file}")
        except Exception as e:
            st.error(f"⚠️ Critical: Cache system not working - {e}")
            logger.error(f"Cache initialization failed: {e}", exc_info=True)
            return

        logger.info("Calling render_sidebar()...")
        token = render_sidebar()
        logger.info(f"render_sidebar() returned: token={'(set)' if token else '(none)'}")

        logger.info("="*60)
        logger.info("DASHBOARD LOADING")
        logger.info("="*60)
        
        # Determine if we should bypass cache
        use_cache = not st.session_state.get("force_refresh", False)
        
        # Try to load from cache first (no token needed!)
        if use_cache:
            logger.info("Attempting to load from cache (no token required)...")
            cache_result = _load_cache()
            if cache_result is not None:
                cached_inst, cached_dec = cache_result
                if not cached_inst.empty:
                    logger.info(f"✓ Cache loaded! {len(cached_inst)} instances, {len(cached_dec)} decisions")
                    st.success(f"✓ Loaded {len(cached_inst)} reviews with {len(cached_dec)} decisions (from cache)")
                    st.divider()
                    tab1, tab2 = st.tabs(["👤 Sponsor Reviews", "🧑 Self-Reviews"])
                    with tab1:
                        tab_sponsor_reviews(cached_inst, cached_dec)
                    with tab2:
                        tab_self_reviews(cached_inst, cached_dec)
                    return
            logger.info("Cache empty or not found, will fetch from API...")

        # Cache didn't work, need token to fetch from API
        if not token:
            st.info("💾 **Cached data not available.** Enter credentials to load from API:")
            st.markdown(
                "**Required Graph API permission:** `AccessReview.Read.All` or `AccessReview.ReadWrite.All` (Application)"
            )
            return

        logger.info("="*60)
        logger.info("DASHBOARD LOADED FROM API")
        logger.info("="*60)
        
        logger.info(f"Loading data from API: use_cache={use_cache}")
        logger.info("→ Calling load_data()...")
        inst_df, dec_df = load_data(token, use_cache=use_cache)
        logger.info("← load_data() completed")
        
        # Reset force_refresh flag
        if "force_refresh" in st.session_state:
            st.session_state["force_refresh"] = False

        logger.info(f"Data loaded: {len(inst_df)} instances, {len(dec_df)} decisions")
        
        if inst_df.empty:
            logger.error("inst_df is empty!")
            st.warning(
                "No Guest Recertification access reviews found. "
                "**Verify the service principal has `AccessReview.Read.All` or `AccessReview.ReadWrite.All` permission.** "
                "Check the console logs for detailed error information."
            )
            st.info("📋 **Troubleshooting:**\n"
                    "1. Verify the service principal has the required permissions in Azure Entra\n"
                    "2. Check the console output for detailed error logs\n"
                    "3. Ensure access review definitions exist in your organization\n"
                    "4. Confirm Guest Recertification naming pattern is used in your definitions")
            return
        
        # Show status
        st.success(f"✓ Loaded {len(inst_df)} reviews with {len(dec_df)} decisions")
        st.divider()

        tab1, tab2 = st.tabs(["👤 Sponsor Reviews", "🧑 Self-Reviews"])
        with tab1:
            tab_sponsor_reviews(inst_df, dec_df)
        with tab2:
            tab_self_reviews(inst_df, dec_df)
        
    except Exception as e:
        logger.error(f"✗ CRITICAL ERROR in main(): {e}", exc_info=True)
        st.error(f"⚠️ Fatal error: {e}")
        st.info("Check console logs for details")


if __name__ == "__main__":
    main()
