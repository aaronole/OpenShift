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
        <h1 style="display:flex; align-items:baseline; gap:1.2rem; padding:1rem 0; font-size:3.1rem;">
            <img src="data:image/png;base64,{icon_b64}" alt=""
                 style="height:0.75em; width:auto;">
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

if st.button("Generate Fills"):
    st.session_state.pop("results", None)

    if not selected:
        st.warning("Select at least one fills schedule.")
    else:
        try:
            with st.spinner("Pulling the schedule and client list from WellSky "
                            "(this can take up to a minute)..."):
                result = core.make_fills(
                    selected, WELLSKY_USERNAME, WELLSKY_PASSWORD,
                    client_file=st.session_state.get("client_override"),
                    weeks=core.WEEK_OPTIONS[week_label],
                )
            result["pulled_at"] = datetime.now(ZoneInfo("America/Denver"))
            st.session_state["results"] = result
        except core.FillsError as e:
            log.warning("Generate Fills: %s", e)
            st.error(str(e))
        except Exception as e:
            log.exception("Generate Fills failed")
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
