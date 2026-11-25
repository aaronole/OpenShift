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
HERE = Path(__file__).parent
ICON_PATH = HERE / "favicon.png"             # browser-tab icon (the logo's arch)
LOGO_PATH = HERE / "stonebridge-logo.png"    # full logo shown on the page


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
    @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@500;600&display=swap');

    #MainMenu, footer {visibility: hidden;}
    .block-container {max-width: 700px; padding-top: 2.5rem; padding-bottom: 4rem;}

    /* header */
    .sb-header {text-align:center; padding:0 0 1.6rem; margin-bottom:1.8rem; border-bottom:1px solid #D9DEE2;}
    .sb-logo {height:92px; width:auto; display:block; margin:0 auto 1.4rem;}
    .sb-header h1 {
        font-family:'Cormorant Garamond', Georgia, serif; font-weight:600;
        font-size:2.4rem; letter-spacing:0.04em; color:#1D5066; padding:0; margin:0;
    }

    /* section heading */
    h3 {
        font-size:0.8rem !important; font-weight:600 !important; letter-spacing:0.16em;
        text-transform:uppercase; color:#6B6B6B !important; padding-bottom:0.4rem !important;
    }

    /* cards */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background:#FFFFFF; border:1px solid #D9DEE2 !important;
        border-radius:4px; box-shadow:0 1px 2px rgba(29,80,102,0.05);
    }

    /* inputs */
    div[data-baseweb="select"] > div, div[data-baseweb="input"] > div {
        border-radius:3px; border-color:#C9D0D5; background:#FFFFFF;
    }

    /* buttons */
    .stButton > button, .stDownloadButton > button {
        border-radius:3px; font-weight:500; letter-spacing:0.04em; padding:0.55rem 1.5rem;
        border:1px solid #1D5066; background:#FFFFFF; color:#1D5066; transition:all .15s ease;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background:#F1F5F7; border-color:#1D5066; color:#1D5066;
    }
    .stButton > button[kind="primary"] {background:#1D5066; color:#FFFFFF;}
    .stButton > button[kind="primary"]:hover {background:#163E50; color:#FFFFFF;}

    /* expanders and status */
    div[data-testid="stExpander"], div[data-testid="stStatus"] {
        border:1px solid #D9DEE2; border-radius:4px; background:#FFFFFF;
    }
    div[data-testid="stAlert"] {border-radius:4px;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ==========================================
# PAGE
# ==========================================

def render_title():
    """Logo above a refined, understated title."""
    try:
        logo_b64 = base64.b64encode(LOGO_PATH.read_bytes()).decode()
        logo = f'<img class="sb-logo" src="data:image/png;base64,{logo_b64}" alt="Stonebridge Home Care Solutions">'
    except OSError:
        logo = ""
    st.markdown(
        f"""
        <div class="sb-header">
            {logo}
            <h1>{APP_TITLE}</h1>
        </div>
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
