"""
Stonebridge Fills Maker - Streamlit app.

Run locally:   streamlit run app.py
"""

import base64
import hmac
import io
import logging
import traceback
import zipfile
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import streamlit as st

import fills_core as core

log = logging.getLogger(__name__)

APP_TITLE = "Stonebridge Fills Maker"
ICON_PATH = Path(__file__).parent / "OpenShift-icon.png"   # lives next to app.py in the repo


def load_icon():
    """The app icon as a PIL image, or None if the file isn't there."""
    try:
        from PIL import Image
        return Image.open(ICON_PATH)
    except Exception:
        return None


# must be the first Streamlit call; page_icon is the browser-tab icon
st.set_page_config(page_title=APP_TITLE, page_icon=load_icon() or "📅")

# ==========================================
# SETTINGS - read from Streamlit secrets
# ==========================================
# On Streamlit Community Cloud: App settings -> Secrets. Locally: create
# .streamlit/secrets.toml next to this file. Contents (top level, no [section]):
#
#   APP_PASSWORD = "the password for schedulers"
#   WELLSKY_USERNAME = "dedicated view-only WellSky username"
#   WELLSKY_PASSWORD = "that account's password"

try:
    APP_PASSWORD = st.secrets["APP_PASSWORD"]
    WELLSKY_USERNAME = st.secrets["WELLSKY_USERNAME"]
    WELLSKY_PASSWORD = st.secrets["WELLSKY_PASSWORD"]
except (KeyError, FileNotFoundError):
    st.error(
        "This app isn't configured yet: APP_PASSWORD, WELLSKY_USERNAME and "
        "WELLSKY_PASSWORD need to be set in the app's secrets."
    )
    st.stop()

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ==========================================
# STYLE (navy + warm neutrals; base colors come from .streamlit/config.toml)
# ==========================================
st.markdown(
    """
    <style>
    #MainMenu, footer {visibility: hidden;}
    .block-container {max-width: 720px; padding-top: 2.5rem; padding-bottom: 4rem;}

    h1 {color:#1F3A5F; font-weight:700; letter-spacing:-0.02em;}
    h3 {color:#1F3A5F; font-weight:600; letter-spacing:-0.01em;}

    /* cards */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background:#FFFFFF; border:1px solid #E6E0D4 !important;
        border-radius:14px; box-shadow:0 1px 3px rgba(31,58,95,0.06);
    }

    /* inputs */
    div[data-baseweb="select"] > div, div[data-baseweb="input"] > div {
        border-radius:10px; border-color:#DDD6C8; background:#FFFFFF;
    }

    /* buttons */
    .stButton > button, .stDownloadButton > button {
        border-radius:10px; font-weight:600; padding:0.55rem 1.3rem;
        border:1px solid #D5CDBD; background:#FFFFFF; color:#1F3A5F;
        transition:all .15s ease;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        border-color:#1F3A5F; color:#1F3A5F; box-shadow:0 2px 8px rgba(31,58,95,0.12);
    }
    .stButton > button[kind="primary"] {
        background:#1F3A5F; color:#FFFFFF; border:1px solid #1F3A5F;
    }
    .stButton > button[kind="primary"]:hover {
        background:#2B4C7A; color:#FFFFFF; border-color:#2B4C7A;
    }

    /* expanders and status */
    div[data-testid="stExpander"], div[data-testid="stStatus"] {
        border:1px solid #E6E0D4; border-radius:12px; background:#FFFFFF;
    }
    div[data-testid="stAlert"] {border-radius:12px;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ==========================================
# PAGE
# ==========================================

def render_title():
    """Title with the icon just to its left, about the same height as the text."""
    try:
        icon_b64 = base64.b64encode(ICON_PATH.read_bytes()).decode()
    except OSError:
        st.title(APP_TITLE)
        return
    st.markdown(
        f"""
        <h1 style="padding:1rem 0; font-size:3.1rem;">
            <img src="data:image/png;base64,{icon_b64}" alt=""
                 style="height:1.1em; width:auto; position:relative; top:-0.09em; margin-right:0.2em;">
            <span>{APP_TITLE}</span>
        </h1>
        """,
        unsafe_allow_html=True,
    )


render_title()


# ==========================================
# PASSWORD GATE
# ==========================================

def check_password():
    entered = st.session_state.get("password_input", "")
    if hmac.compare_digest(entered.encode(), APP_PASSWORD.encode()):
        st.session_state["authed"] = True
        st.session_state["password_wrong"] = False
    else:
        st.session_state["password_wrong"] = True
    st.session_state["password_input"] = ""


if not st.session_state.get("authed"):
    st.text_input(
        "Password",
        type="password",
        key="password_input",
        on_change=check_password,
    )
    if st.session_state.get("password_wrong"):
        st.error("Incorrect password.")
    st.stop()


# ==========================================
# MAIN APP
# ==========================================

with st.container(border=True):
    st.subheader("Fills to make:")

    week_label = st.selectbox(
        "Week",
        options=list(core.WEEK_OPTIONS),
        index=list(core.WEEK_OPTIONS).index(core.DEFAULT_WEEK),
        label_visibility="collapsed",
    )

    selected = [
        label
        for label in core.REGIONS
        if st.checkbox(label, key=f"region_{label}")
    ]

    generate = st.button("Generate Fills", type="primary")
    st.caption("If WellSky is slow, the app keeps retrying on its own for up to "
               f"{core.TOTAL_PATIENCE_SECONDS // 60} minutes. Keep this tab open while it works.")

if generate:
    st.session_state.pop("results", None)

    if not selected:
        st.warning("Select at least one fills schedule.")
    else:
        status = st.status("Working with WellSky...", expanded=True)
        try:
            result = core.make_fills(
                selected, WELLSKY_USERNAME, WELLSKY_PASSWORD,
                client_file=st.session_state.get("client_override"),
                weeks=core.WEEK_OPTIONS[week_label],
                progress=status.write,
            )
            result["pulled_at"] = datetime.now(ZoneInfo("America/Denver"))
            st.session_state["results"] = result
            status.update(label="Done", state="complete", expanded=False)
        except core.FillsError as e:
            log.warning("Generate Fills: %s", e)
            status.update(label="Couldn't finish", state="error", expanded=False)
            st.error(str(e))
        except Exception as e:
            log.exception("Generate Fills failed")
            status.update(label="Couldn't finish", state="error", expanded=False)
            st.error("Something went wrong generating the fills. Please try again, "
                     "and let your manager know if it keeps happening.")
            with st.expander("Technical details (send these to your manager)"):
                st.code(f"{type(e).__name__}: {e}\n\n{traceback.format_exc()}")


# ==========================================
# RESULTS (kept in session so downloads don't reset the page)
# ==========================================

results = st.session_state.get("results")
if results:
    t = results["pulled_at"]
    st.success(
        f"Fills are ready. Pulled from WellSky at {t.hour % 12 or 12}:{t.minute:02d} "
        f"{t.strftime('%p')} on {t.strftime('%b')} {t.day}."
    )

    for label, (filename, data) in results["files"].items():
        count = results["counts"][label]
        st.download_button(
            f"Download {label} ({count} open shift{'s' if count != 1 else ''})",
            data=data,
            file_name=filename,
            mime=XLSX_MIME,
            key=f"dl_{label}",
        )

    if len(results["files"]) > 1:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for filename, data in results["files"].values():
                zf.writestr(filename, data)
        st.download_button(
            "Download all (.zip)",
            data=zip_buf.getvalue(),
            file_name="Fills.zip",
            mime="application/zip",
            key="dl_zip",
        )

    if results["unmatched"]:
        with st.expander(
            f"{len(results['unmatched'])} client(s) with open shifts weren't found in the client list"
        ):
            st.write(
                "These clients' shifts are missing from the fills above. "
                "Re-download the Clients report from WellSky if the list is out of date."
            )
            st.write(", ".join(results["unmatched"]))


# ==========================================
# BACKUP OPTION
# ==========================================

with st.expander("Use a Clients.xlsx file instead (optional)"):
    st.file_uploader(
        "Client list (Clients.xlsx)",
        type=["xlsx"],
        key="client_override",
        help="Normally not needed: the client list is pulled from WellSky automatically. "
             "If a file is uploaded here, it is used instead.",
    )
