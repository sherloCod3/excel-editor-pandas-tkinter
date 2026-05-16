import base64
import html
import json
import os
import stat
import time

import streamlit as st
import streamlit.components.v1 as components

from core import fetch_soc_data, load_data
import pandas as pd
from streamlit.delta_generator import DeltaGenerator

st.set_page_config(
    page_title="SecOps Portal",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

CONFIG_FILE = "config.json"

# CSS and JS injection.
INJECT_JS = """
<script>
(function patchUI() {
    var iconFont = "'Material Symbols Rounded','Material Icons','Material Symbols Outlined'";

    function patch() {
        var doc = window.parent.document;
        if (!doc || !doc.body) return;

        // 1. Restore Material Icons font
        doc.querySelectorAll(
            '.material-icons,.material-symbols-rounded,.material-symbols-outlined,' +
            '[class*="material-icons"],[class*="material-symbols"]'
        ).forEach(function(el) {
            el.style.setProperty('font-family', iconFont, 'important');
        });

        // 2. Hide Deploy button / 3-dot menu; keep sidebar toggle
        ['[data-testid="stDeployButton"]',
         '[data-testid="stStatusWidget"]',
         '[data-testid="stToolbarActionButtonContainer"]'
        ].forEach(function(sel) {
            doc.querySelectorAll(sel).forEach(function(el) {
                el.style.setProperty('display','none','important');
            });
        });

        // 3. Hide radio circle wrapper; keep our text labels
        doc.querySelectorAll(
            '[data-testid="stSidebar"] [data-testid="stRadio"] label > div:first-child'
        ).forEach(function(el) {
            el.style.setProperty('display','none','important');
        });

        // 4. Fix "uploadUpload" — hide icon span inside upload button
        doc.querySelectorAll(
            '[data-testid="stFileUploaderDropzone"] button span,' +
            '[data-testid="stFileUploader"] button > span:first-child'
        ).forEach(function(span) {
            if (span.textContent.trim().toLowerCase() === 'upload') {
                span.style.setProperty('display','none','important');
            }
        });

        // 5. Typing effect on terminal headings
        doc.querySelectorAll('.terminal-heading').forEach(function(h) {
            var span = h.querySelectorAll('span')[1]; // second span = command text
            if (!span || span.dataset.typed) return;
            span.dataset.typed = '1';
            var full = span.textContent;
            span.textContent = '';
            var i = 0;
            (function type() {
                if (i <= full.length) { span.textContent = full.slice(0,i++); setTimeout(type,38); }
            })();
        });
    }

    patch();
    setTimeout(patch, 250);
    setTimeout(patch, 700);
    setTimeout(patch, 1500);

    new MutationObserver(patch).observe(
        window.parent.document.body, {childList: true, subtree: true}
    );
})();
</script>
"""

def inject_css():
    """Inject the Ghost in the Shell Terminal stylesheet + DOM patches."""
    css_path = os.path.join(os.path.dirname(__file__), "assets", "style.css")
    with open(css_path, "r") as f:
        css = f.read()
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)

    # st.markdown strips <script> tags even with unsafe_allow_html=True.
    # components.html() renders in an iframe that DOES execute JS;
    # window.parent.document reaches the outer Streamlit app DOM.
    components.html(INJECT_JS, height=0)


# Config helpers.
def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        except Exception:
            return {"webhook_url": ""}
    return {"webhook_url": ""}


def save_config(url: str) -> None:
    with open(CONFIG_FILE, "w") as f:
        json.dump({"webhook_url": url}, f)
    # Restrict to owner read/write only so the embedded webhook token is not world-readable.
    os.chmod(CONFIG_FILE, stat.S_IRUSR | stat.S_IWUSR)


# Reusable UI components.
def terminal_heading(command: str, title: str):
    """Render a terminal-prompt style page heading with blinking cursor."""
    st.markdown(
        f"""
        <div class="terminal-heading">
            <span class="prompt">&gt;</span>
            <span>{command}</span>
            <span style="color:var(--text-secondary);font-size:1rem;margin-left:0.5em">// {title}</span>
            <span class="cursor"></span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def threat_score_bar(score: int):
    """Render a color-coded threat score meter with a fill bar."""
    if score <= 30:
        color = "var(--accent-safe)"
        gradient = "linear-gradient(90deg, #30D158, #30D158)"
        level = "LOW RISK"
        level_color = "var(--accent-safe)"
    elif score <= 70:
        color = "var(--accent-warn)"
        gradient = "linear-gradient(90deg, #30D158, #F5A623)"
        level = "MEDIUM RISK"
        level_color = "var(--accent-warn)"
    else:
        color = "var(--accent-alert)"
        gradient = "linear-gradient(90deg, #30D158, #F5A623, #FF3B5C)"
        level = "HIGH RISK"
        level_color = "var(--accent-alert)"

    st.markdown(
        f"""
        <div class="threat-bar-container">
            <div class="threat-bar-label">Threat Score</div>
            <div class="threat-bar-score" style="color:{color}">{score}<span style="font-size:1rem;color:var(--text-secondary)">/100</span></div>
            <div class="threat-bar-track">
                <div class="threat-bar-fill" style="width:{score}%;background:{gradient};box-shadow: 0 0 8px {color};"></div>
            </div>
            <div class="threat-bar-zones">
                <span>0 LOW</span>
                <span>30 ──── 70</span>
                <span>HIGH 100</span>
            </div>
            <div class="threat-status" style="color:{level_color}">⬡ {level}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def result_card(label: str, value: str, color: str = "var(--accent-cyber)", delay_ms: int = 0):
    """Render a single result info card."""
    st.markdown(
        f"""
        <div class="result-card" style="animation-delay:{delay_ms}ms;border-left:3px solid {color};">
            <div class="result-card-label">{label}</div>
            <div class="result-card-value" style="color:{color}">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def ai_summary_card(summary: str):
    """Render the AI-generated plain-English threat summary from n8n.

    Only shown when n8n returns a non-empty summary (i.e. OPENAI_API_KEY is
    configured in the n8n environment). In mock mode or when AI is disabled,
    this function is a no-op.
    """
    if not summary or "AI summary disabled" in summary:
        return
    st.markdown(
        f"""
        <div class="result-card" style="border-left:3px solid var(--accent-cyber);margin-top:0.5rem;">
            <div class="result-card-label"
                 style="display:flex;align-items:center;gap:0.4em;">
                <span style="color:var(--accent-cyber)">⬡</span> AI THREAT SUMMARY
                <span style="font-size:0.65rem;opacity:0.5;margin-left:auto;">n8n · LLM enrichment</span>
            </div>
            <div class="result-card-value"
                 style="color:var(--text-primary);font-size:0.85rem;
                        font-family:'Space Mono',monospace;line-height:1.6;
                        font-weight:400;white-space:pre-wrap;">{html.escape(summary)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def mode_badge(is_live: bool):
    """Render MOCK MODE or LIVE MODE badge in the sidebar."""
    if is_live:
        st.sidebar.markdown(
            """
            <div class="live-mode-badge">
                <span class="dot"></span> LIVE MODE
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.sidebar.markdown(
            """
            <div class="mock-mode-badge">
                <span class="dot"></span> MOCK MODE
            </div>
            """,
            unsafe_allow_html=True,
        )


# Navigation.
def get_image_base64(path: str) -> str:
    """Read an image file and return its base64 string."""
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")
    except Exception:
        return ""

def main():
    inject_css()

    logo_path = os.path.join(os.path.dirname(__file__), "assets", "logo.png")
    logo_base64 = get_image_base64(logo_path)
    
    if logo_base64:
        # Sidebar avatar — rendered as its own markdown block (no f-string nesting)
        st.sidebar.markdown(
            f'<div style="text-align:center;padding-top:0.5rem;padding-bottom:0.75rem;">'
            f'<img src="data:image/png;base64,{logo_base64}" '
            f'style="width:120px;height:120px;object-fit:cover;border-radius:50%;'
            f'box-shadow:0 0 20px rgba(0,255,204,0.4);border:2px solid var(--accent-cyber);" '
            f'alt="Operator Face"></div>',
            unsafe_allow_html=True,
        )
        # Full-page watermark injected as a <style> block
        st.markdown(
            f'<style>[data-testid="stAppViewContainer"]::before{{'
            f'content:"";position:fixed;top:0;left:0;right:0;bottom:0;'
            f'background-image:url("data:image/png;base64,{logo_base64}");'
            f'background-size:cover;background-position:center;background-repeat:no-repeat;'
            f'opacity:0.05;z-index:-1;pointer-events:none;filter:grayscale(100%);}}</style>',
            unsafe_allow_html=True,
        )

    # [SECOPS] logo text — separate call, no dynamic interpolation
    st.sidebar.markdown(
        '<div style="font-family:\'Space Mono\',monospace;font-size:1.1rem;'
        'color:#00FFCC;letter-spacing:0.1em;'
        'border-bottom:1px solid rgba(0,255,204,0.18);'
        'padding-bottom:0.75rem;margin-bottom:1rem;text-align:center;">'
        '<span class="secops-logo" data-text="[SECOPS]">[SECOPS]</span>'
        '</div>',
        unsafe_allow_html=True,
    )

    app_mode = st.sidebar.radio(
        "Navigate",
        ["⬡  LOG_ANALYSIS", "⬡  THREAT_INTEL"],
        label_visibility="collapsed",
    )

    if app_mode == "⬡  LOG_ANALYSIS":
        show_data_editor()
    elif app_mode == "⬡  THREAT_INTEL":
        show_soc_investigator()


# Log analysis page.
def show_data_editor():
    terminal_heading("LOG_ANALYSIS", "Upload & clean network logs")

    st.markdown(
        "<p>Upload a CSV or Excel file containing network logs or IP lists. "
        "Edit inline and export the cleaned dataset.</p>",
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader(
        "Drop file here — CSV / XLSX / XLS",
        type=["csv", "xlsx", "xls"],
        label_visibility="visible",
    )

    if "df" not in st.session_state:
        st.session_state.df = None

    col1, col2 = st.columns([1, 3])
    with col1:
        if st.button("⬡  LOAD SAMPLE LOGS"):
            st.session_state.df = pd.DataFrame({
                "Timestamp": [
                    "2026-05-15 10:00:01",
                    "2026-05-15 10:02:14",
                    "2026-05-15 10:05:33",
                    "2026-05-15 10:11:05",
                ],
                "Source IP":      ["192.168.1.50",   "10.0.0.15",      "172.16.0.100",  "192.168.1.50"],
                "Destination IP": ["8.8.8.8",        "198.51.100.23",  "203.0.113.5",   "104.21.25.10"],
                "Port":           [53,               443,              22,              80],
                "Protocol":       ["UDP",            "TCP",            "TCP",           "TCP"],
                "Action":         ["Allowed",        "Allowed",        "Blocked",       "Allowed"],
            })

    if uploaded_file is not None:
        st.session_state.df = load_data(uploaded_file)

    if st.session_state.df is not None:
        st.markdown("<br>", unsafe_allow_html=True)
        edited_df = st.data_editor(
            st.session_state.df,
            num_rows="dynamic",
            width="stretch",
        )
        st.markdown("<br>", unsafe_allow_html=True)
        csv = edited_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="↓  EXPORT CSV",
            data=csv,
            file_name="cleaned_logs.csv",
            mime="text/csv",
        )
    else:
        if uploaded_file is not None:
            st.error("⬡  Failed to parse file — check format and try again.")


# SOC investigator page.
def _scan_log(target: str) -> DeltaGenerator:
    """Show an animated terminal log stream while the investigation runs."""
    box = st.empty()
    steps = [
        f"INIT  ─ resolving target: {html.escape(target)}",
        "QUERY ─ querying threat intelligence feed #1...",
        "QUERY ─ querying threat intelligence feed #2...",
        "QUERY ─ cross-referencing known blocklists...",
        "GEO   ─ fetching geolocation metadata...",
        "SCORE ─ computing threat score algorithm...",
        "OUT   ─ compiling analysis report...",
    ]
    shown: list[str] = []
    for step in steps:
        shown.append(step)
        lines = "".join(
            f'<div class="log-line" style="animation-delay:{i*80}ms">'
            f'<span style="color:var(--accent-cyber)">{'✓' if i < len(shown)-1 else '→'}</span>'
            f' {s}</div>'
            for i, s in enumerate(shown)
        )
        box.markdown(f'<div class="scan-log">{lines}</div>', unsafe_allow_html=True)
        time.sleep(0.17)
    return box


def _history_html(history: list) -> str:
    """Render recent scan history as HTML."""
    if not history:
        return ""
    items = ""
    for h in history:
        score = h["score"]
        if score <= 30:
            sc, bg = "var(--accent-safe)", "rgba(48,209,88,0.1)"
        elif score <= 70:
            sc, bg = "var(--accent-warn)", "rgba(245,166,35,0.1)"
        else:
            sc, bg = "var(--accent-alert)", "rgba(255,59,92,0.1)"
        mal = "MALICIOUS" if h["malicious"] else "CLEAR"
        items += (
            f'<div class="history-item">'
            f'<span>◎ {html.escape(h["target"])}</span>'
            f'<span class="history-score" style="color:{sc};background:{bg}">{score} · {mal}</span>'
            f'</div>'
        )
    return f'<div class="history-wrap"><div class="history-title">Recent Scans</div>{items}</div>'


def show_soc_investigator():
    terminal_heading("THREAT_INTEL", "IP & URL investigation via n8n SOC agent")

    # Initialize session state on first run.
    if "scan_history" not in st.session_state:
        st.session_state.scan_history = []
    if "scan_count" not in st.session_state:
        st.session_state.scan_count = 0

    config = load_config()
    saved_url = config.get("webhook_url", "")

    # Sidebar webhook configuration.
    st.sidebar.markdown(
        "<div style='font-family:\"Space Mono\",monospace;font-size:0.7rem;"
        "color:var(--text-secondary,#8B949E);text-transform:uppercase;"
        "letter-spacing:0.12em;margin-bottom:0.4rem;'>Webhook Config</div>",
        unsafe_allow_html=True,
    )
    webhook_url = st.sidebar.text_input(
        "Webhook URL",
        value=saved_url,
        type="password",
        placeholder="https://n8n.example.com/webhook/...",
        label_visibility="collapsed",
    )

    if st.sidebar.button("💾  PERSIST URL"):
        save_config(webhook_url or "")
        st.sidebar.success("URL saved.")

    mode_badge(is_live=bool(webhook_url))

    # Sidebar status bar.
    mode_str = "LIVE" if webhook_url else "MOCK"
    mode_color = "var(--accent-safe)" if webhook_url else "var(--accent-warn)"
    st.sidebar.markdown(
        f"""
        <div class="sidebar-status">
            <div class="sidebar-status-row">
                <span>SESSION SCANS</span><span>{st.session_state.scan_count}</span>
            </div>
            <div class="sidebar-status-row">
                <span>MODE</span>
                <span style="color:{mode_color}">{mode_str}</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Main target input.
    st.markdown("<br>", unsafe_allow_html=True)
    target = st.text_input(
        "Target",
        placeholder="Enter IP address or URL to investigate...",
        label_visibility="collapsed",
    )

    st.markdown("<br>", unsafe_allow_html=True)
    run = st.button("⬡  RUN INVESTIGATION")

    if run:
        if target:
            # Run the animated scan log before the real request so the UI feels responsive.
            log_box = _scan_log(target)
            results = fetch_soc_data(target, webhook_url)
            log_box.empty()

            st.session_state.scan_count += 1

            if "error" in results:
                st.error(f"⬡  Connection error: {results['error']}")
            else:
                # Keep only the five most recent scans to avoid unbounded session state growth.
                st.session_state.scan_history.insert(0, {
                    "target": target,
                    "score": int(results.get("threat_score", 0)),
                    "malicious": results.get("known_malicious", False),
                })
                st.session_state.scan_history = st.session_state.scan_history[:5]

                st.markdown("<br>", unsafe_allow_html=True)

                score = int(results.get("threat_score", 0))
                threat_score_bar(score)

                # Render the AI summary only when n8n returns one (requires OPENAI_API_KEY in n8n environment).
                ai_summary_card(str(results.get("summary", "")))

                st.markdown("<br>", unsafe_allow_html=True)

                col1, col2 = st.columns(2)
                with col1:
                    result_card(
                        "Geolocation",
                        f"◎  {html.escape(results.get('location', 'N/A'))}",
                        color="var(--accent-cyber)",
                        delay_ms=150,
                    )
                with col2:
                    malicious = results.get("known_malicious", False)
                    result_card(
                        "Threat Status",
                        "⚠  MALICIOUS" if malicious else "✓  CLEAR",
                        color="var(--accent-alert)" if malicious else "var(--accent-safe)",
                        delay_ms=300,
                    )

                st.markdown("<br>", unsafe_allow_html=True)

                with st.expander("⬡  RAW RESPONSE DATA"):
                    st.json(results)

                st.markdown(
                    _history_html(st.session_state.scan_history),
                    unsafe_allow_html=True,
                )
        else:
            st.warning("⬡  No target specified — enter an IP or URL above.")
    else:
        # Show a placeholder prompt before any scan has been run.
        if not st.session_state.scan_history:
            st.markdown(
                """
                <div class="empty-state">
                    <span class="empty-state-icon">⬡</span>
                    AWAITING TARGET<br>
                    <span style="font-size:0.72rem;opacity:0.55;">
                        Enter an IP or URL above and run investigation
                    </span>
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            # Keep history visible between scans so the user can compare recent results.
            st.markdown(
                _history_html(st.session_state.scan_history),
                unsafe_allow_html=True,
            )

    st.divider()

    with st.expander("ℹ  HOW TO INTERPRET RESULTS"):
        st.markdown("""
        **Threat Score (0–100)**
        A normalized score where `0` is benign and `100` is critical.
        - `0–30` — Low risk. Likely safe or benign noise.
        - `31–70` — Medium risk. Warrants manual review.
        - `71–100` — High risk. Known malicious activity; immediate action required.

        **Geolocation**
        Country of origin for the IP. Unexpected geolocations may indicate compromised credentials or an attacker pivot.

        **Threat Status**
        Definitive `MALICIOUS` or `CLEAR` based on known blocklists and threat feeds.

        *No Webhook URL configured? The portal runs in **MOCK MODE** — results are randomized for development and testing.*
        """)


if __name__ == "__main__":
    main()
