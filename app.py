import hashlib
import base64
import re
import time
from datetime import datetime
from html import escape
from pathlib import Path
from zoneinfo import ZoneInfo

import streamlit as st
from sentence_transformers import SentenceTransformer

from advanced_engine import MODEL_NAME, rank_resumes_advanced, tailoring_eligibility
from database import (
    authenticate,
    create_user,
    get_or_create_oauth_user,
    initialize_database,
    list_analyses,
    load_analysis_bundle,
    save_analysis,
)
from document_parser import extract_document_text
from recruiter_tools import (
    NOT_SUITABLE,
    SUITABLE,
    add_suitability,
    build_filtered_resume_zip,
    filter_recruiter_results,
    original_resume_payloads,
    recruiter_display_table,
)
from resume_tailor import build_targeted_questions, create_enhanced_resumes_batch


st.set_page_config(page_title="MatchMind AI", page_icon="assets/matchmind_star.png", layout="wide")
initialize_database()


COBALT_CSS = """
<style>
    :root {
        --cobalt-900: #10223f;
        --cobalt-700: #1769e0;
        --cobalt-500: #04a4cc;
        --cobalt-100: #e8f1ff;
    }
    .stApp {
        background:
            radial-gradient(circle at 90% 4%, rgba(23,105,224,.12), transparent 28rem),
            linear-gradient(180deg, #f8fbff 0%, #f3f7fd 100%);
    }
    [data-testid="stHeader"] { background: transparent; }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #091a32 0%, #102a4d 100%);
        border-right: 1px solid rgba(255,255,255,.08);
    }
    [data-testid="stSidebar"] * { color: #edf6ff; }
    [data-testid="stSidebar"] [data-baseweb="select"] * { color: #10223f; }
    .block-container { max-width: 1220px; padding-top: 2rem; padding-bottom: 3rem; }
    .cobalt-brand { display:flex; align-items:center; gap:.8rem; margin-bottom:1.9rem; }
    .cobalt-mark {
        width:52px; height:52px; display:grid; place-items:center;
        color:white; background:transparent;
    }
    .cobalt-mark img {width:100%;height:100%;object-fit:contain}
    .cobalt-brand-title { color:#10223f; font-size:1.15rem; font-weight:700; }
    .cobalt-brand-subtitle { color:#677892; font-size:.78rem; }
    .cobalt-hero {
        padding:1.35rem 1.5rem; margin-bottom:1.2rem; border-radius:18px;
        background:linear-gradient(125deg,#0d2850 0%,#174f9a 65%,#0789ac 100%);
        box-shadow:0 18px 45px rgba(16,55,105,.18); color:white;
    }
    .cobalt-hero h1 { color:white; font-size:2rem; margin:0; }
    .cobalt-hero p { color:#d7ebff; margin:.4rem 0 0; }
    .cobalt-step {
        display:inline-flex; align-items:center; gap:.45rem; color:#1769e0; font-weight:700;
        font-size:.8rem; letter-spacing:.04em; text-transform:uppercase; margin-bottom:.3rem;
    }
    .cobalt-panel-title { color:#10223f; font-size:1.05rem; font-weight:700; margin-bottom:.15rem; }
    .cobalt-panel-copy { color:#6a7990; font-size:.84rem; margin-bottom:.7rem; }
    div[data-testid="stForm"], div[data-testid="stExpander"],
    [data-testid="stDataFrame"], [data-testid="stFileUploader"] section {
        border-color:#d7e3f2 !important; border-radius:14px !important;
    }
    div[data-testid="stForm"] { background:rgba(255,255,255,.78); padding:1.1rem; }
    .stButton > button, .stFormSubmitButton > button, .stDownloadButton > button {
        border-radius:10px; min-height:2.7rem; font-weight:650;
    }
    .stButton > button[kind="primary"], .stFormSubmitButton > button[kind="primary"] {
        background:linear-gradient(90deg,#1769e0,#078db4); border:0;
        box-shadow:0 8px 22px rgba(23,105,224,.22);
    }
    .cobalt-divider { display:flex; align-items:center; gap:.8rem; color:#8390a4; font-size:.8rem; margin:.9rem 0; }
    .cobalt-divider:before,.cobalt-divider:after { content:""; flex:1; height:1px; background:#dce5f2; }
    .cobalt-login {
        max-width:560px; margin:2rem auto; padding:1.6rem; border:1px solid #d9e5f3;
        border-radius:20px; background:rgba(255,255,255,.88); box-shadow:0 24px 70px rgba(21,65,116,.13);
    }
    .cobalt-login h1 { color:#10223f; text-align:center; margin:.5rem 0 .25rem; }
    .cobalt-login p { color:#6b7b92; text-align:center; margin:0 0 1rem; }
    .cobalt-result-head {
        margin-top:1.6rem; padding:1rem 1.2rem; border-radius:14px 14px 0 0;
        color:white; background:linear-gradient(100deg,#102a4d,#1769e0);
    }
    @media (max-width: 700px) {
        .block-container { padding:1rem; }
        .cobalt-hero h1 { font-size:1.5rem; }
        .cobalt-login { margin:.5rem auto; padding:1rem; }
    }
    .stApp, [data-testid="stAppViewContainer"] {
        background:radial-gradient(circle at 88% 2%,#102c50 0%,#051225 34%,#030b17 100%) !important;
        color:#eef6ff !important;
    }
    [data-testid="stMainBlockContainer"] { color:#eef6ff; }
    [data-testid="stMainBlockContainer"] h1,
    [data-testid="stMainBlockContainer"] h2,
    [data-testid="stMainBlockContainer"] h3,
    [data-testid="stMainBlockContainer"] label,
    [data-testid="stMainBlockContainer"] p { color:#eef6ff; }
    [data-testid="stMainBlockContainer"] .stCaptionContainer p { color:#9fb2ce; }
    div[data-testid="stForm"], div[data-testid="stExpander"],
    [data-testid="stFileUploader"] section {
        background:#0b1d36 !important; border-color:#214366 !important;
    }
    [data-baseweb="input"] > div, [data-baseweb="textarea"] > div,
    [data-baseweb="select"] > div {
        background:#0b1d36 !important; border-color:#214366 !important; color:#eef6ff !important;
    }
    input, textarea { color:#eef6ff !important; caret-color:#4d9cff !important; }
    div.stButton > button, div.stDownloadButton > button,
    div.stFormSubmitButton > button {
        color:#ffffff !important; background:linear-gradient(100deg,#1769e0,#04a4cc) !important;
        border:1px solid #328ce8 !important; box-shadow:0 8px 24px rgba(0,85,190,.24) !important;
    }
    div.stButton > button:disabled { color:#8fa6c4 !important; background:#183252 !important; border-color:#284766 !important; }
    [data-testid="stSidebar"] div.stButton > button {
        color:#eef6ff !important; background:#102c50 !important; border-color:#28527d !important;
    }
    button[data-baseweb="tab"] { color:#9fb2ce !important; }
    button[data-baseweb="tab"][aria-selected="true"] { color:#4d9cff !important; }
    .cobalt-login { background:#0b1d36; border-color:#214366; box-shadow:0 25px 80px rgba(0,0,0,.3); }
    .cobalt-login h1 { color:#eef6ff; }
    .cobalt-login p { color:#9fb2ce; }
    .cobalt-brand-title { color:#eef6ff; }
    .cobalt-brand-subtitle, .cobalt-panel-copy { color:#9fb2ce; }
    .cobalt-panel-title { color:#eef6ff; }
    .cobalt-divider { color:#9fb2ce; }
    .cobalt-divider:before,.cobalt-divider:after { background:#214366; }
    .cobalt-flow {display:flex;align-items:center;gap:.55rem;margin:1rem 0 1.4rem;color:#9fb2ce;font-size:.82rem}
    .cobalt-flow b {display:inline-grid;place-items:center;min-width:1.7rem;height:1.7rem;border-radius:50%;background:#102c50;color:#78b6ff}
    .cobalt-flow b.active {background:linear-gradient(100deg,#1769e0,#04a4cc);color:white}
    .cobalt-flow b.done {background:#123f46;color:#51e0b2}
    .cobalt-flow span {white-space:nowrap;color:#9fb2ce}
    .cobalt-flow i {height:1px;background:#214366;flex:1}
    .shortlist-shell {border:1px solid #214366;border-radius:22px;overflow:hidden;background:#07172a;margin-top:.5rem}
    .shortlist-top {display:flex;align-items:center;justify-content:space-between;padding:1.15rem 1.35rem;border-bottom:1px solid #214366;background:#0c213c}
    .shortlist-top strong {font-size:1.1rem}.shortlist-meta {color:#9fb2ce;font-size:.86rem;margin-top:.2rem}
    .stage-pills {display:flex;align-items:center;gap:.75rem;padding:1rem 1.35rem .35rem;color:#9fb2ce;font-size:.84rem}
    .stage-pills b {display:inline-grid;place-items:center;width:1.85rem;height:1.85rem;border-radius:50%;background:#12345c;color:#8ec4ff}
    .stage-pills .active {background:#368cff;color:white}.stage-pills .done {background:#12345c;color:#51e0b2}
    .candidate-card {display:flex;align-items:center;gap:1rem;padding:.95rem 1.1rem;border:1px solid #214366;border-radius:13px;background:#0b1d36;margin:.6rem 0}
    .candidate-rank {display:grid;place-items:center;min-width:2.35rem;height:2.35rem;border-radius:10px;background:#10345f;color:#78b6ff;font-weight:700}
    .candidate-main {flex:1}.candidate-name {font-weight:750;color:#eef6ff}.candidate-meta {font-size:.8rem;color:#9fb2ce;margin-top:.18rem}
    .eligible {color:#51e0b2;font-weight:750}.not-eligible {color:#ffbe68;font-weight:700}
    .metric-chip {padding:.3rem .55rem;border-radius:8px;background:#102c50;color:#78b6ff;font-size:.78rem;font-weight:700}
    .role-choice {min-height:12rem;padding:1.2rem;border:1px solid #214366;border-radius:16px;background:#0b1d36;margin:.4rem 0 .8rem}
    .role-choice h3 {margin:.1rem 0 .45rem;color:#eef6ff!important}.role-choice p {color:#9fb2ce!important;margin:.2rem 0!important}
    .role-badge {display:inline-flex;padding:.28rem .55rem;border-radius:8px;background:#102c50;color:#78b6ff;font-size:.76rem;font-weight:700;text-transform:uppercase;letter-spacing:.04em}
    .tailor-note {padding:1rem 1.1rem;border:1px solid #214366;border-radius:13px;background:#0b1d36;margin:.7rem 0 1rem}
    .tailor-note strong {color:#eef6ff}.tailor-note p {color:#9fb2ce!important;margin:.3rem 0 0!important;font-size:.85rem}
    [data-testid="stDataFrame"] {background:#0b1d36!important;border:1px solid #214366!important;border-radius:14px!important;overflow:hidden}
    .welcome-title {white-space:nowrap;font-size:clamp(1.45rem,4vw,2rem)!important}
    .role-choice-wide {min-height:0;padding:1rem 1.15rem;margin:.35rem 0;border-radius:18px}
    .role-choice-wide h3 {margin:.35rem 0 .25rem!important}
    .aurora-welcome {
        position:relative;overflow:hidden;margin:.2rem 0 1.35rem;padding:1.25rem 1.5rem 1.65rem;
        border:1px solid #214d74;border-radius:26px;background:
        radial-gradient(circle at 16% 12%,rgba(21,180,218,.2),transparent 18rem),
        radial-gradient(circle at 86% 86%,rgba(255,137,39,.14),transparent 19rem),
        linear-gradient(140deg,#07162a 0%,#0b2847 58%,#101d35 100%);
        box-shadow:0 24px 70px rgba(0,0,0,.28)
    }
    .aurora-welcome:before {content:"";position:absolute;width:240px;height:240px;right:-95px;top:-85px;border:42px solid rgba(87,223,235,.055);border-radius:50%}
    .aurora-brand {position:relative;z-index:1;display:flex;align-items:center;justify-content:space-between;gap:1rem}
    .aurora-brand-main {display:flex;align-items:center;gap:.7rem;color:#f5fbff;font-weight:750}
    .aurora-brand-main img {width:38px;height:38px;object-fit:contain}
    .aurora-brand-note {color:#8daac5;font-size:.75rem}
    .aurora-hero {position:relative;z-index:1;max-width:790px;margin:1.1rem auto .2rem;text-align:center}
    .aurora-kicker {color:#69e4e7;font-size:.73rem;font-weight:750;text-transform:uppercase;letter-spacing:.13em}
    .aurora-hero h1 {margin:.55rem 0 0!important;color:#f7fbff!important;font-size:clamp(2rem,5vw,3.25rem)!important;line-height:1.08;letter-spacing:-.045em}
    .aurora-hero h1 span {color:#69e4e7}
    .aurora-hero p {max-width:650px;margin:.8rem auto 0!important;color:#a7bfd5!important;line-height:1.6}
    .aurora-process {display:flex;align-items:center;justify-content:center;gap:.55rem;flex-wrap:wrap;margin-top:1.15rem;color:#cbdbe9;font-size:.75rem}
    .aurora-process span {display:inline-flex;align-items:center;gap:.4rem;padding:.42rem .65rem;border:1px solid rgba(113,177,217,.16);border-radius:999px;background:rgba(255,255,255,.055)}
    .aurora-process i {width:20px;height:20px;display:inline-grid;place-items:center;border-radius:50%;background:#123f5f;color:#6de2e5;font-style:normal;font-size:.66rem;font-weight:800}
    .aurora-process b {color:#52708c;font-weight:500}
    .aurora-role-card {position:relative;overflow:hidden;min-height:225px;margin:.15rem 0 .55rem;padding:1.35rem;border:1px solid #244b70;border-radius:22px;background:linear-gradient(155deg,rgba(16,49,80,.97),rgba(9,29,54,.97));box-shadow:0 18px 42px rgba(0,0,0,.2)}
    .aurora-role-card:after {content:"";position:absolute;width:120px;height:120px;right:-42px;top:-50px;border-radius:50%;background:rgba(72,211,225,.09)}
    .aurora-role-card.recruiter:after {background:rgba(255,144,42,.11)}
    .aurora-role-top {position:relative;z-index:1;display:flex;align-items:center;justify-content:space-between;gap:1rem}
    .aurora-role-icon {width:46px;height:46px;display:grid;place-items:center;border-radius:14px;background:#123f5f;color:#71e1e6;font-weight:850;font-size:1.05rem}
    .aurora-role-card.recruiter .aurora-role-icon {background:#3b3027;color:#ffad4c}
    .aurora-role-label {color:#7897b4;font-size:.7rem;font-weight:750;text-transform:uppercase;letter-spacing:.12em}
    .aurora-role-card h2 {position:relative;z-index:1;margin:1rem 0 .4rem!important;color:#f4f9ff!important;font-size:1.35rem!important}
    .aurora-role-card p {position:relative;z-index:1;color:#9fb6ca!important;font-size:.85rem;line-height:1.55;margin:0!important}
    .aurora-role-points {position:relative;z-index:1;display:flex;gap:.45rem;flex-wrap:wrap;margin-top:.9rem}
    .aurora-role-points span {padding:.28rem .48rem;border-radius:7px;background:rgba(84,192,211,.09);color:#9dd7df;font-size:.68rem}
    .aurora-role-card.recruiter .aurora-role-points span {background:rgba(255,151,55,.09);color:#ffc17d}
    div.st-key-welcome_candidate.stButton>button {background:linear-gradient(100deg,#1769e0,#04a4cc)!important}
    div.st-key-welcome_recruiter.stButton>button {color:#172235!important;background:linear-gradient(100deg,#ffb638,#ff7e31)!important;border-color:#ff9e38!important;box-shadow:0 9px 24px rgba(255,126,49,.2)!important}
    .auth-visual {
        position:relative;min-height:610px;padding:2rem;border:1px solid #28527d;border-radius:24px;
        overflow:hidden;background:
        radial-gradient(circle at 75% 20%,rgba(4,164,204,.28),transparent 13rem),
        linear-gradient(145deg,#0d3150 0%,#07182e 55%,#061223 100%);
        box-shadow:0 24px 70px rgba(0,0,0,.28)
    }
    .auth-visual:before,.auth-visual:after {content:"";position:absolute;border:1px solid rgba(120,182,255,.18);border-radius:50%}
    .auth-visual:before {width:330px;height:330px;right:-120px;top:-110px}
    .auth-visual:after {width:210px;height:210px;left:-100px;bottom:-90px}
    .auth-brand {display:flex;align-items:center;gap:.7rem;position:relative;z-index:2}
    .auth-brand img {width:42px;height:42px}.auth-brand strong {display:block;color:#fff;font-size:1.05rem}
    .auth-brand span {display:block;color:#8eb8df;font-size:.72rem}
    .auth-copy {position:relative;z-index:2;margin-top:3.2rem;max-width:28rem}
    .auth-copy small {display:inline-flex;padding:.32rem .62rem;border:1px solid #2c6594;border-radius:999px;color:#7eddf0;font-weight:700;text-transform:uppercase;letter-spacing:.05em}
    .auth-copy h2 {font-size:2.05rem;line-height:1.15;margin:.9rem 0 .65rem;color:#fff!important}
    .auth-copy p {color:#9fc3df!important;line-height:1.6}
    .match-graphic {position:relative;z-index:2;height:245px;margin-top:1.5rem}
    .resume-sheet {position:absolute;left:4%;top:18px;width:42%;height:185px;padding:1rem;border:1px solid #3a6d99;border-radius:15px;background:rgba(7,28,50,.88);transform:rotate(-3deg);box-shadow:0 18px 35px rgba(0,0,0,.28)}
    .resume-sheet b {display:block;color:#dff3ff;font-size:.82rem;margin-bottom:.7rem}.resume-line {height:5px;margin:.55rem 0;border-radius:5px;background:#234b70}.resume-line.accent {width:70%;background:linear-gradient(90deg,#1769e0,#04a4cc)}
    .match-score {position:absolute;left:37%;top:54px;width:102px;height:102px;display:grid;place-items:center;border:8px solid #174d75;border-top-color:#29c8dc;border-right-color:#3b8cff;border-radius:50%;background:#071a31;color:#fff;font-size:1.15rem;font-weight:800;box-shadow:0 0 35px rgba(4,164,204,.2)}
    .match-score span {display:block;color:#87accb;font-size:.58rem;font-weight:600;text-transform:uppercase;text-align:center}
    .candidate-stack {position:absolute;right:3%;top:20px;width:35%;display:grid;gap:.65rem}
    .candidate-node {display:flex;align-items:center;gap:.55rem;padding:.62rem .7rem;border:1px solid #28527d;border-radius:12px;background:rgba(12,39,68,.92);color:#cce7fb;font-size:.7rem}
    .candidate-node i {width:25px;height:25px;display:grid;place-items:center;border-radius:50%;font-style:normal;font-weight:800;background:#123b65;color:#77b8ff}.candidate-node.active {border-color:#1ba9c9}.candidate-node.active i {background:#0e5d70;color:#62edc3}
    .auth-form-head {padding:.5rem 0 1rem}.auth-form-head small {color:#69baff;font-weight:750;text-transform:uppercase;letter-spacing:.05em}.auth-form-head h1 {margin:.35rem 0 .35rem;color:#eef6ff!important;font-size:2rem}.auth-form-head p {color:#9fb2ce!important;margin:0!important}
    @media (max-width: 700px) {
        .block-container {padding:.8rem .7rem 2rem!important;max-width:100%!important}
        .cobalt-brand {gap:.6rem;margin-bottom:1rem}
        .cobalt-mark {width:42px;height:42px;flex:0 0 42px}
        .cobalt-brand-title {font-size:1rem}
        .cobalt-brand-subtitle {font-size:.72rem}
        .cobalt-hero {padding:1rem;margin-bottom:.9rem;border-radius:14px}
        .cobalt-hero h1 {font-size:1.45rem;line-height:1.2;overflow-wrap:anywhere}
        .cobalt-hero p {font-size:.88rem;line-height:1.45}
        .cobalt-login {margin:.35rem auto;padding:.9rem;border-radius:15px}
        .cobalt-login h1 {font-size:1.55rem}
        .role-choice {min-height:0;padding:1rem;margin:.3rem 0 .55rem}
        .welcome-title {font-size:clamp(1.25rem,7vw,1.55rem)!important}
        .role-choice-wide {padding:.85rem 1rem;margin:.25rem 0}
        .aurora-welcome {padding:1rem .85rem 1.25rem;border-radius:20px;margin-top:0}
        .aurora-brand-note {display:none}
        .aurora-brand-main img {width:34px;height:34px}
        .aurora-hero {margin-top:1rem}
        .aurora-hero h1 {font-size:clamp(1.75rem,9vw,2.25rem)!important}
        .aurora-hero p {font-size:.84rem}
        .aurora-process {gap:.35rem}.aurora-process b {display:none}
        .aurora-role-card {min-height:0;padding:1.05rem;margin-top:.2rem}
        .auth-visual {min-height:315px;padding:1.15rem;border-radius:18px}
        .auth-copy {margin-top:1.25rem}.auth-copy h2 {font-size:1.45rem}.auth-copy p {font-size:.83rem}
        .match-graphic {height:125px;margin-top:.5rem}.resume-sheet {height:105px;padding:.65rem}.resume-sheet .resume-line:nth-of-type(n+4){display:none}
        .match-score {left:38%;top:30px;width:68px;height:68px;border-width:5px;font-size:.8rem}.candidate-stack {top:5px}.candidate-node {padding:.35rem;font-size:.58rem}.candidate-node i {width:20px;height:20px}
        .cobalt-flow {gap:.28rem;margin:.7rem 0 1rem;width:100%}
        .cobalt-flow b {min-width:1.9rem;width:1.9rem;height:1.9rem;font-size:.78rem}
        .cobalt-flow span {display:none}
        .cobalt-flow i {min-width:.5rem}
        .candidate-card {align-items:flex-start;flex-wrap:wrap;gap:.55rem;padding:.8rem}
        .candidate-main {min-width:calc(100% - 3rem)}
        .candidate-name,.candidate-meta,.cobalt-panel-title {overflow-wrap:anywhere}
        .metric-chip,.eligible,.not-eligible {margin-left:2.9rem;white-space:normal}
        .shortlist-top,.stage-pills {padding:.85rem;flex-wrap:wrap}
        [data-testid="stDataFrame"] {display:none!important}
        [data-testid="stFileUploaderDropzone"] {padding:.75rem!important;min-height:auto!important}
        [data-testid="stFileUploaderDropzone"] button {min-height:2.75rem}
        div[data-baseweb="tab-list"] {overflow-x:auto;scrollbar-width:thin;justify-content:flex-start}
        button[data-baseweb="tab"] {flex:0 0 auto;white-space:nowrap;padding-left:.7rem!important;padding-right:.7rem!important}
        div.stButton > button,div.stDownloadButton > button,div.stFormSubmitButton > button {
            width:100%;min-height:2.75rem
        }
        input,textarea {font-size:16px!important}
    }
    @media (max-width: 420px) {
        .cobalt-flow {gap:.2rem}
        .cobalt-flow i {min-width:.3rem}
        .metric-chip,.eligible,.not-eligible {margin-left:0}
    }
</style>
"""
st.markdown(COBALT_CSS, unsafe_allow_html=True)

LOGO_DATA_URI = "data:image/png;base64," + base64.b64encode(
    Path(__file__).with_name("assets").joinpath("matchmind_star.png").read_bytes()
).decode("ascii")

APP_ROLES = {"candidate", "recruiter"}


def compact_candidate_label(candidate: str, maximum: int = 26) -> str:
    """Keep candidate tabs readable on narrow screens without hiding the full name in content."""
    label = str(candidate)
    if len(label) <= maximum:
        return label
    return label[: maximum - 1].rstrip() + "…"


def clear_workflow_state() -> None:
    """Clear one analysis without changing the signed-in account or selected role."""
    for key in (
        "loaded_results", "analysis_context", "tailored_files", "result_view",
        "draft_job_title", "draft_job_description", "role_ocr_notice",
    ):
        st.session_state.pop(key, None)


def current_app_role() -> str | None:
    role = str(st.session_state.get("app_role", "")).strip().lower()
    return role if role in APP_ROLES else None


@st.cache_resource(show_spinner="Loading semantic intelligence…")
def load_semantic_model():
    return SentenceTransformer(MODEL_NAME)


@st.cache_resource(show_spinner="Preparing OCR for this scanned document…")
def load_ocr_reader():
    import easyocr

    return easyocr.Reader(["en"], gpu=False)


def google_auth_configured() -> bool:
    try:
        auth = st.secrets.get("auth", {})
        return all(auth.get(key) for key in ("redirect_uri", "cookie_secret", "client_id", "client_secret"))
    except (FileNotFoundError, KeyError):
        return False


def external_user() -> dict | None:
    streamlit_user = getattr(st, "user", None)
    if not getattr(streamlit_user, "is_logged_in", False):
        return None
    subject = str(streamlit_user.get("sub", ""))
    email = str(streamlit_user.get("email", ""))
    name = str(streamlit_user.get("name", email or "Google user"))
    return get_or_create_oauth_user("google", subject, email, name)


def current_user() -> dict | None:
    oidc_user = external_user()
    return oidc_user or st.session_state.get("local_user")


def sign_out() -> None:
    st.session_state.pop("local_user", None)
    st.session_state.pop("app_role", None)
    clear_workflow_state()
    streamlit_user = getattr(st, "user", None)
    if getattr(streamlit_user, "is_logged_in", False) and hasattr(st, "logout"):
        st.logout()
    st.rerun()


def switch_app_role() -> None:
    """Return to the shared welcome page without signing the account out."""
    st.session_state.pop("app_role", None)
    clear_workflow_state()
    st.rerun()


def set_verified_choice(state_key: str, value: str) -> None:
    """Store one mutually exclusive Yes/No answer across Streamlit reruns."""
    st.session_state[state_key] = value


def is_yes_no_question(question: str) -> bool:
    """Return True only when Yes or No completely answers the question."""
    return bool(re.match(
        r"^(?:do|does|did|is|are|was|were|can|could|have|has)\b",
        question.strip(),
        re.IGNORECASE,
    ))


def yes_answer_evidence(question: str) -> str:
    """Tag an affirmative answer so it is merged into the correct résumé section."""
    skill_match = re.search(r"hands-on experience with\s+(.+?)\?*$", question, re.IGNORECASE)
    if skill_match:
        return f"[SKILL] {skill_match.group(1).strip(' .?')}"
    evidence_text = re.sub(r"\?$", ".", question.strip())
    return f"[EVIDENCE] {evidence_text}"


def written_answer_evidence(question: str, answer: str) -> str:
    """Tag a written answer by purpose without changing the user's wording."""
    question_text = question.casefold()
    if "qualification" in question_text or "certification" in question_text:
        category = "QUALIFICATION"
    elif "measurable result" in question_text:
        category = "RESULT"
    else:
        category = "ACCOMPLISHMENT"
    return f"[{category}] {answer.strip()}"


def show_role_selection() -> None:
    st.markdown(
        f"""
        <div class="aurora-welcome">
          <div class="aurora-brand">
            <div class="aurora-brand-main"><img src="{LOGO_DATA_URI}" alt="MatchMind AI">MatchMind AI</div>
            <div class="aurora-brand-note">Candidate success · Recruiter clarity</div>
          </div>
          <div class="aurora-hero">
            <div class="aurora-kicker">Smarter résumé decisions</div>
            <h1 class="welcome-title">One platform. <span>Two clear paths.</span></h1>
            <p>Choose how you want to use MatchMind AI. Your Candidate and Recruiter journeys remain completely separate.</p>
            <div class="aurora-process">
              <span><i>1</i>Résumé</span><b>→</b>
              <span><i>2</i>Job match</span><b>→</b>
              <span><i>3</i>Clear decision</span>
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    candidate_column, recruiter_column = st.columns(2, gap="large")
    with candidate_column:
        st.markdown(
            '<div class="aurora-role-card candidate"><div class="aurora-role-top">'
            '<span class="aurora-role-icon">C</span><span class="aurora-role-label">Candidate</span></div>'
            '<h2>Check and enhance my résumé</h2>'
            '<p>Understand your match, confirm genuine evidence, and create a focused application résumé.</p>'
            '<div class="aurora-role-points"><span>Check suitability</span><span>Enhance résumé</span></div></div>',
            unsafe_allow_html=True,
        )
        candidate_clicked = st.button(
            "Continue as Candidate →",
            type="primary",
            use_container_width=True,
            key="welcome_candidate",
        )
        if candidate_clicked:
            st.session_state.app_role = "candidate"
            clear_workflow_state()
            st.rerun()
    with recruiter_column:
        st.markdown(
            '<div class="aurora-role-card recruiter"><div class="aurora-role-top">'
            '<span class="aurora-role-icon">R</span><span class="aurora-role-label">Recruiter</span></div>'
            '<h2>Screen and compare applicants</h2>'
            '<p>Rank multiple résumés and review clear suitability evidence for every candidate.</p>'
            '<div class="aurora-role-points"><span>Compare applicants</span><span>Download shortlist</span></div></div>',
            unsafe_allow_html=True,
        )
        recruiter_clicked = st.button(
            "Continue as Recruiter →",
            use_container_width=True,
            key="welcome_recruiter",
        )
        if recruiter_clicked:
            st.session_state.app_role = "recruiter"
            clear_workflow_state()
            st.rerun()


def show_login(app_role: str) -> None:
    role_label = "Candidate" if app_role == "candidate" else "Recruiter"
    role_description = (
        "Check your suitability and create a truthful, job-focused résumé."
        if app_role == "candidate"
        else "Screen multiple applicants with explainable matching results."
    )
    if st.button("← Change role"):
        switch_app_role()
    visual_column, login_column = st.columns([1.08, .92], gap="large", vertical_alignment="top")
    with visual_column:
        visual_title = (
            "Turn your résumé into a focused application."
            if app_role == "candidate"
            else "Move from applications to a clear shortlist."
        )
        st.markdown(
            f"""
            <div class="auth-visual">
              <div class="auth-brand">
                <img src="{LOGO_DATA_URI}" alt="MatchMind AI">
                <div><strong>MatchMind AI</strong><span>Explainable candidate matching</span></div>
              </div>
              <div class="auth-copy">
                <small>{role_label} workspace</small>
                <h2>{visual_title}</h2>
                <p>Evidence-based matching, clear gaps and decisions you can understand.</p>
              </div>
              <div class="match-graphic" aria-hidden="true">
                <div class="resume-sheet"><b>Résumé evidence</b><div class="resume-line accent"></div><div class="resume-line"></div><div class="resume-line"></div><div class="resume-line"></div><div class="resume-line"></div></div>
                <div class="match-score"><div>78%<span>job match</span></div></div>
                <div class="candidate-stack"><div class="candidate-node active"><i>✓</i>Strong evidence</div><div class="candidate-node"><i>2</i>Skills reviewed</div><div class="candidate-node"><i>3</i>Gaps explained</div></div>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with login_column:
        st.markdown(
            f'<div class="auth-form-head"><small>{role_label}</small><h1>Welcome back</h1><p>{role_description}</p></div>',
            unsafe_allow_html=True,
        )
        if google_auth_configured():
            if st.button("Continue with Google", type="primary", use_container_width=True):
                st.login()
            st.markdown('<div class="cobalt-divider">or use your local account</div>', unsafe_allow_html=True)
        else:
            st.caption("Use your local username and password. Google sign-in can be enabled later.")
            with st.expander("Google sign-in setup"):
                st.markdown(
                    """
                    1. Open Google Cloud Console and create an **OAuth 2.0 Client ID** with application type **Web application**.
                    2. Add `http://localhost:8501/oauth2callback` as an **Authorized redirect URI**.
                    3. Copy `.streamlit/secrets.local.example.toml` to `.streamlit/secrets.toml`.
                    4. Replace the three placeholder values with your Google Client ID, Client Secret and a long random cookie secret.
                    5. Stop Streamlit with **Ctrl+C**, then run `streamlit run app.py` again.

                    Never upload `secrets.toml` to GitHub or share its contents.
                    """
                )

        login_tab, register_tab = st.tabs(["Sign in", "Create account"])
        with login_tab:
            with st.form("login_form"):
                username = st.text_input("Username")
                password = st.text_input("Password", type="password")
                login_clicked = st.form_submit_button("Sign in", type="primary", use_container_width=True)
            if login_clicked:
                user = authenticate(username, password)
                if user:
                    st.session_state.local_user = user
                    st.rerun()
                st.error("Incorrect username or password.")
        with register_tab:
            with st.form("register_form"):
                new_username = st.text_input("Choose username")
                new_password = st.text_input("Choose password", type="password")
                confirm_password = st.text_input("Confirm password", type="password")
                register_clicked = st.form_submit_button("Create account", use_container_width=True)
            if register_clicked:
                if new_password != confirm_password:
                    st.error("Passwords do not match.")
                else:
                    ok, message = create_user(new_username, new_password)
                    (st.success if ok else st.error)(message)


def show_stage_header(active: int, app_role: str, show_tailoring: bool = False) -> None:
    """Display progress only; movement happens through explicit Previous/Continue buttons."""
    if app_role == "candidate":
        labels = ["Job", "Résumé", "Review", "Enhance"]
    else:
        labels = ["Job", "Résumés", "Screening"]
    items = []
    for number, label in enumerate(labels, start=1):
        state = "active" if number == active else "done" if number < active else ""
        marker = "✓" if number < active else str(number)
        items.append(f'<b class="{state}">{marker}</b><span>{number}. {label}</span>')
        if number < len(labels):
            items.append("<i></i>")
    st.markdown('<div class="cobalt-flow">' + "".join(items) + "</div>", unsafe_allow_html=True)


def show_recruiter_results(results, context) -> None:
    role = (context or {}).get("job_title", "Candidate shortlist")
    job_description = (context or {}).get("job_description", "")
    screening_results = add_suitability(results, job_description)
    eligible_count = int((screening_results["Suitability"] == SUITABLE).sum())
    st.markdown(f"## {role} shortlist")
    st.caption(f"{len(results)} resumes evaluated against job description, skills and experience.")
    show_stage_header(3, "recruiter")

    st.markdown("### Ranked candidates")
    st.caption(
        "Candidate details are available here. On phones, use the candidate cards below; "
        "the wide ranking table is hidden to prevent horizontal scrolling. Use the filters to narrow the shortlist."
    )
    filter_1, filter_2 = st.columns(2)
    with filter_1:
        suitability_filter = st.selectbox(
            "Suitability",
            ["All candidates", SUITABLE, NOT_SUITABLE],
            key="recruiter_suitability_filter",
        )
    with filter_2:
        experience_rule = st.selectbox(
            "Experience",
            ["Any experience", "At least", "More than", "At most"],
            key="recruiter_experience_rule",
        )
    experience_years = 0.0
    if experience_rule != "Any experience":
        experience_years = float(st.slider(
            "Experience years",
            min_value=0,
            max_value=30,
            value=4,
            step=1,
            key="recruiter_experience_years",
        ))
    match_filter, skill_filter = st.columns(2)
    with match_filter:
        minimum_match = st.slider(
            "Minimum match %",
            min_value=0,
            max_value=100,
            value=0,
            step=5,
            key="recruiter_minimum_match",
        )
    with skill_filter:
        minimum_skill_coverage = st.slider(
            "Minimum skill coverage %",
            min_value=0,
            max_value=100,
            value=0,
            step=5,
            key="recruiter_minimum_skill_coverage",
        )
    filtered_results = filter_recruiter_results(
        screening_results,
        suitability=suitability_filter,
        experience_rule=experience_rule,
        experience_years=experience_years,
        minimum_match=minimum_match,
        minimum_skill_coverage=minimum_skill_coverage,
    )
    display_table = recruiter_display_table(filtered_results)
    st.caption(f"Showing {len(filtered_results)} of {len(screening_results)} candidates.")
    st.dataframe(display_table, hide_index=True, use_container_width=True)
    if filtered_results.empty:
        st.info("No candidates match the selected filters. Adjust a filter to continue.")
    for _, row in filtered_results.iterrows():
        eligible, reason = tailoring_eligibility(row, job_description)
        eligibility_class = "eligible" if eligible else "not-eligible"
        eligibility_label = "Suitable for this Job" if eligible else "Not Suitable for this Job"
        candidate = escape(str(row["Candidate"]))
        experience = escape(str(row["Experience"]))
        experience_label = f"{experience} years" if experience.replace(".", "", 1).isdigit() else experience
        st.markdown(
            f'<div class="candidate-card"><div class="candidate-rank">{int(row["Rank"])}</div>'
            f'<div class="candidate-main"><div class="candidate-name">{candidate}</div>'
            f'<div class="candidate-meta">Detected experience: {experience_label}</div></div>'
            f'<span class="metric-chip">{float(row["Match %"]):.1f}% match</span>'
            f'<span class="{eligibility_class}">{eligibility_label}</span></div>',
            unsafe_allow_html=True,
        )
        with st.expander(f"View candidate details · {row['Candidate']}"):
            strength_column, gap_column = st.columns(2)
            with strength_column:
                st.markdown("**Why this candidate may be suitable**")
                st.success(row["Strengths"])
            with gap_column:
                st.markdown("**Why this candidate may not be suitable**")
                st.warning(row["Gaps"])
            metric_1, metric_2, metric_3 = st.columns(3)
            metric_1.metric("Experience", f"{row['Experience']} years")
            metric_2.metric("Job Description Match", f"{row['Semantic %']:.1f}%")
            metric_3.metric("Skill coverage", f"{row['Skill coverage %']:.1f}%")
            displayed_reason = reason.replace("Eligible:", "Suitable for this Job:")
            (st.success if eligible else st.warning)(displayed_reason)
    if not filtered_results.empty:
        table_column, resume_column = st.columns(2)
        with table_column:
            st.download_button(
                "Download filtered table (CSV)",
                display_table.to_csv(index=False).encode("utf-8-sig"),
                "filtered_candidate_table.csv",
                "text/csv",
                use_container_width=True,
            )
        payloads = original_resume_payloads(filtered_results, (context or {}).get("original_files", {}))
        with resume_column:
            if len(payloads) == 1:
                payload = payloads[0]
                st.download_button(
                    "Download filtered résumé",
                    payload["content"],
                    payload["name"],
                    payload["mime"],
                    use_container_width=True,
                )
            elif len(payloads) > 1:
                st.download_button(
                    f"Download {len(payloads)} résumés + table (ZIP)",
                    build_filtered_resume_zip(display_table, payloads),
                    "MatchMind_filtered_candidates.zip",
                    "application/zip",
                    use_container_width=True,
                )
            else:
                st.caption("Original résumé files are unavailable for this saved analysis. Rerun it once to enable résumé downloads.")
    st.divider()
    if eligible_count:
        st.success(f"{eligible_count} candidate{'s meet' if eligible_count != 1 else ' meets'} the current suitability requirements.")
    else:
        st.info("No candidate currently meets all suitability requirements. Review each candidate's specific gaps above.")
    st.caption("Recruiter mode provides screening and decision support only. Candidate résumés are not modified.")


def show_candidate_results(results, context) -> None:
    """Show one candidate a focused suitability review before enhancement."""
    row = results.iloc[0]
    job_description = (context or {}).get("job_description", "")
    eligible, reason = tailoring_eligibility(row, job_description)
    status_label = "Suitable for this Job" if eligible else "Not Suitable for this Job"

    st.markdown(f"## Your match for {(context or {}).get('job_title', 'this job')}")
    st.caption("Your résumé was compared with the job description, required skills and stated experience range.")
    show_stage_header(3, "candidate")
    (st.success if eligible else st.warning)(f"{status_label} — {reason.replace('Eligible:', '').strip()}")

    metric_1, metric_2, metric_3 = st.columns(3)
    metric_1.metric("Overall Match", f"{float(row['Match %']):.1f}%")
    metric_2.metric("Job Description Match", f"{float(row['Semantic %']):.1f}%")
    metric_3.metric("Skill coverage", f"{float(row['Skill coverage %']):.1f}%")

    strength_column, gap_column = st.columns(2)
    with strength_column:
        st.markdown("### What matches")
        st.success(row["Strengths"])
    with gap_column:
        st.markdown("### What is missing or unclear")
        st.warning(row["Gaps"])

    experience = row["Experience"]
    experience_label = f"{experience} years" if isinstance(experience, (int, float)) else str(experience)
    st.info(f"Detected experience: {experience_label}")
    st.caption(
        "You can now confirm missing evidence, add truthful details and create a job-focused résumé. "
        "Unverified qualifications will not be invented."
    )
    if st.button("Review details and enhance my résumé →", type="primary", use_container_width=True):
        st.session_state.result_view = "tailoring"
        st.rerun()


def show_resume_tailoring(results) -> None:
    context = st.session_state.get("analysis_context")
    app_role = current_app_role() or "candidate"
    candidate_mode = app_role == "candidate"
    if not context:
        st.error(
            "This saved analysis was created by an older app version and does not contain source résumé text. "
            "Run it once more; all new analyses remain available for tailoring after reload."
        )
        return

    heading_col, back_col = st.columns([4, 1], vertical_alignment="center")
    with heading_col:
        st.markdown("## Enhance your résumé" if candidate_mode else f"## {context.get('job_title', 'Candidate')} tailoring")
        st.caption("Create a complete, job-focused application résumé without deleting source content.")
    with back_col:
        if st.button("Back to review" if candidate_mode else "Back to ranking", use_container_width=True):
            st.session_state.result_view = "ranking"
            st.rerun()
    show_stage_header(4, app_role, show_tailoring=True)
    st.markdown("### Current suitability result")
    if not candidate_mode:
        st.caption(
            "Available when overall job match and skill coverage are both at least 50%, "
            "and detected experience falls inside the job's stated range. Job Description Match is explanatory, not a separate rejection gate."
        )
    tailoring_candidates = []
    original_files = context.get("original_files", {})
    for _, row in results.iterrows():
        eligible, reason = tailoring_eligibility(row, context["job_description"])
        candidate = row["Candidate"]
        displayed_reason = reason.replace("Eligible:", "Suitable for this Job:")
        if eligible:
            st.success(f"✓ {row['Candidate']} — {displayed_reason}")
            if candidate in context.get("resume_texts", {}) and candidate in original_files:
                tailoring_candidates.append(candidate)
            else:
                st.info(
                    f"{candidate}: upload the original PDF or DOCX and run the analysis again. "
                    "The original file is required to preserve fonts, colours and layout."
                )
        else:
            st.warning(f"✕ {row['Candidate']} — Not Suitable for this Job. {displayed_reason}")
            if (
                candidate_mode
                and candidate in context.get("resume_texts", {})
                and candidate in original_files
            ):
                tailoring_candidates.append(candidate)

    if not tailoring_candidates:
        st.warning("No complete saved résumé text is currently available for enhancement.")
        return

    st.markdown(
        '<div class="tailor-note"><strong>Complete résumé enhancement</strong>'
        '<p>The job requirements are compared with each complete résumé. Supported summary, skills, experience and project content is strengthened without inventing skills, metrics, qualifications, titles or dates. Original content is not filtered out and no coloured highlighting is added.</p></div>',
        unsafe_allow_html=True,
    )
    st.info(
        "The enhanced résumé includes one editable DOCX and one selectable, searchable PDF. "
        "Fonts, colours, headings, tables, columns, spacing and page design are retained from the uploaded file. "
        "If that formatting cannot be retained safely, generation stops instead of producing a plain-text résumé."
    )

    st.success(
        "Built-in evidence-based résumé enhancement is ready. "
        "No API key or external account is required."
    )

    result_rows = {row["Candidate"]: row for _, row in results.iterrows()}
    additional_information = {}
    st.markdown("### Step 1 of 3 — Add verified information (optional)")
    st.info(
        "These questions help strengthen the résumé. Answer only when the information is true. "
        "You may leave any question unanswered, then continue to Step 2 below."
    )
    candidate_tabs = (
        st.tabs(
            [
                f"{index + 1}. {compact_candidate_label(candidate)}"
                for index, candidate in enumerate(tailoring_candidates)
            ]
        )
        if len(tailoring_candidates) > 1
        else [st.container()]
    )
    for candidate_index, (candidate, candidate_tab) in enumerate(zip(tailoring_candidates, candidate_tabs)):
        row = result_rows[candidate]
        questions = build_targeted_questions(
            context["job_description"],
            context["resume_texts"][candidate],
            str(row.get("Gaps", "")),
        )
        yes_no_questions = [question for question in questions if is_yes_no_question(question)]
        written_questions = [question for question in questions if not is_yes_no_question(question)]
        answered_count = sum(
            st.session_state.get(
                f"verified_choice_{hashlib.sha256(f'{candidate}|{question}'.encode('utf-8')).hexdigest()[:12]}"
            ) in {"Yes", "No"}
            for question in yes_no_questions
        )

        with candidate_tab:
            with st.container(border=True):
                st.markdown(f"#### {candidate}")
                st.caption(
                    f"Candidate {candidate_index + 1} of {len(tailoring_candidates)} · "
                    "Optional answers are merged into the appropriate résumé section."
                )
                if yes_no_questions:
                    st.caption(f"Confirmation questions answered: {answered_count}/{len(yes_no_questions)}")

                verified_evidence = []
                question_columns = st.columns(2) if len(yes_no_questions) > 1 else [st.container()]
                for index, question in enumerate(yes_no_questions):
                    question_id = hashlib.sha256(f"{candidate}|{question}".encode("utf-8")).hexdigest()[:12]
                    answer_key = f"verified_choice_{question_id}"
                    with question_columns[index % len(question_columns)]:
                        with st.container(border=True):
                            st.markdown(f"**{question}**  ·  `Optional`")
                            selected = st.segmented_control(
                                "Choose Yes or No",
                                options=["Yes", "No"],
                                selection_mode="single",
                                key=answer_key,
                                label_visibility="collapsed",
                            )
                            if selected == "Yes":
                                st.success("✓ Selected: Yes")
                                verified_evidence.append(yes_answer_evidence(question))
                            elif selected == "No":
                                st.info("✓ Selected: No — nothing will be added")
                            else:
                                st.caption("No answer selected")

                if written_questions:
                    st.markdown("##### Written information (optional)")
                    written_columns = st.columns(2) if len(written_questions) > 1 else [st.container()]
                    for index, question in enumerate(written_questions):
                        question_id = hashlib.sha256(f"{candidate}|{question}".encode("utf-8")).hexdigest()[:12]
                        with written_columns[index % len(written_columns)]:
                            written_answer = st.text_area(
                                f"{question} (Optional)",
                                key=f"verified_text_{question_id}",
                                placeholder="Type only true, verifiable information. Leave blank when not applicable.",
                                height=90,
                            ).strip()
                        if written_answer:
                            verified_evidence.append(written_answer_evidence(question, written_answer))

                other_evidence = st.text_area(
                    "Other verified information (optional)",
                    key=f"tailoring_other_evidence_{candidate}",
                    placeholder="Type any other job-relevant fact that is not already in the résumé.",
                    height=80,
                ).strip()
                if other_evidence:
                    verified_evidence.append(f"[EVIDENCE] {other_evidence}")
                additional_information[candidate] = "\n".join(verified_evidence)

    st.markdown("### Step 2 of 3 — Confirm authorization")
    tailoring_input_fingerprint = hashlib.sha256(
        repr(sorted(additional_information.items())).encode("utf-8")
    ).hexdigest()
    st.markdown('<span style="color:#ff6b6b;font-weight:700">* Required</span>', unsafe_allow_html=True)
    consent = st.checkbox(
        "I confirm this résumé and the information supplied are mine or I am authorized to create a job-matched copy. *",
        key="tailoring_consent",
    )
    st.markdown("### Step 3 of 3 — Generate enhanced files")
    if not consent:
        st.caption("Complete the required authorization checkbox above to enable generation.")
    if st.button(
        "Create enhanced DOCX and PDF résumé" if candidate_mode else f"Create enhanced DOCX and PDF resumes ({len(tailoring_candidates)})",
        disabled=not consent,
        use_container_width=True,
    ):
        batch_started = time.perf_counter()
        with st.status("Preparing enhanced résumé files…", expanded=True) as generation_status:
            progress = st.progress(0, text="Starting résumé enhancement…")

            def report_generation_progress(completed, total, candidate, phase):
                if phase == "tailoring" and completed == 0:
                    generation_status.write("Comparing all suitable résumés with the job description in parallel…")
                elif phase == "formatting":
                    generation_status.write(f"Recreating {candidate} as DOCX and searchable PDF…")
                ratio = 0.5 * completed / max(total, 1) if phase == "tailoring" else 0.5 + 0.5 * completed / max(total, 1)
                progress.progress(
                    min(ratio, 1.0),
                    text=(
                        f"Content enhancement completed for {completed} of {total} candidates"
                        if phase == "tailoring"
                        else f"Formatted {completed} of {total} candidates"
                    ),
                )

            generation_requests = []
            for candidate in tailoring_candidates:
                original = original_files.get(candidate)
                generation_requests.append({
                    "candidate_name": candidate,
                    "target_role": context.get("job_title", ""),
                    "job_description": context["job_description"],
                    "source_resume": context["resume_texts"][candidate],
                    "source_name": original["name"] if original else None,
                    "source_bytes": base64.b64decode(original["content_b64"]) if original else None,
                    "additional_information": additional_information.get(candidate, ""),
                    "reported_gaps": str(result_rows[candidate].get("Gaps", "")),
                })

            generated, failures = create_enhanced_resumes_batch(
                generation_requests,
                progress_callback=report_generation_progress,
                maximum_workers=3,
            )
            generation_status.update(
                label="Résumé files are ready" if generated else "Résumé generation failed",
                state="complete" if generated else "error",
                expanded=False,
            )
        st.session_state.tailored_files = generated
        st.session_state.tailoring_input_fingerprint = tailoring_input_fingerprint
        st.session_state.tailoring_elapsed_seconds = round(time.perf_counter() - batch_started, 1)
        if failures:
            st.warning("Some resumes could not be generated: " + "; ".join(failures))

    tailored_files_are_current = (
        st.session_state.get("tailoring_input_fingerprint") == tailoring_input_fingerprint
    )
    if st.session_state.get("tailored_files") and not tailored_files_are_current:
        st.warning("Your verified answers changed. Create the enhanced files again before downloading them.")
    if st.session_state.get("tailored_files") and tailored_files_are_current:
        st.success(f"Résumé enhancement completed in {st.session_state.get('tailoring_elapsed_seconds', 0):.1f} seconds.")
    current_tailored_files = st.session_state.get("tailored_files", {}) if tailored_files_are_current else {}
    for candidate, files in current_tailored_files.items():
        st.markdown(f"#### {candidate}")
        st.success("Complete tailored résumé ready in editable DOCX and searchable PDF formats.")
        st.caption(files.get("layout_note", "The complete résumé was prepared without coloured highlighting."))
        if files.get("matching_layout"):
            st.caption("The PDF was exported from the DOCX, so both download layouts match.")
        (st.success if files.get("qa_passed") else st.warning)(files.get("qa_summary", "PDF quality review was not available."))
        if files.get("warning"):
            st.warning(files["warning"])
        if files.get("focused_terms"):
            st.caption("Job focus: " + ", ".join(files["focused_terms"][:12]))
        with st.expander("What changed, strong matches and remaining gaps"):
            st.markdown("**Changes made**")
            for item in files.get("change_summary", []) or ["Source content and layout were retained as closely as possible."]:
                st.markdown(f"- {item}")
            st.markdown("**Strong matches**")
            for item in files.get("strong_matches", []) or ["No additional structured match explanation was returned."]:
                st.markdown(f"- {item}")
            st.markdown("**Remaining genuine gaps**")
            for item in files.get("remaining_gaps", []) or ["No additional qualification gap was identified; human verification is still required."]:
                st.markdown(f"- {item}")
            if files.get("questions"):
                st.markdown("**Questions that could strengthen a future version**")
                for item in files["questions"]:
                    st.markdown(f"- {item}")
        original = original_files.get(candidate)
        docx_column, pdf_column, original_column = st.columns(3)
        with docx_column:
            st.download_button(
                "Download tailored DOCX",
                files["docx_bytes"],
                files["docx_name"],
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"enhanced_docx_{candidate}",
                use_container_width=True,
            )
        with pdf_column:
            st.download_button(
                "Download tailored PDF",
                files["pdf_bytes"],
                files["pdf_name"],
                "application/pdf",
                key=f"enhanced_pdf_{candidate}",
                use_container_width=True,
            )
        if original:
            with original_column:
                st.download_button(
                    "Download original (unchanged)",
                    base64.b64decode(original["content_b64"]),
                    original["name"],
                    original["mime"],
                    key=f"original_{candidate}",
                    use_container_width=True,
                )


def analysis_option_label(analysis_id: int | None, history: list[dict]) -> str:
    if analysis_id is None:
        return "Select an analysis…"
    item = next(entry for entry in history if entry["id"] == analysis_id)
    try:
        created = datetime.fromisoformat(item["created_at"]).astimezone(ZoneInfo("Asia/Kolkata"))
        created_label = created.strftime("%d %b %Y, %I:%M %p")
    except (TypeError, ValueError):
        created_label = "saved analysis"
    return f"#{item['id']} · {item['job_title']} · {created_label}"


app_role = current_app_role()
if app_role is None:
    show_role_selection()
    st.stop()

user = current_user()
if user is None:
    show_login(app_role)
    st.stop()

candidate_mode = app_role == "candidate"
display_name = user.get("display_name") or user.get("username") or "User"
with st.sidebar:
    brand_icon, brand_name = st.columns([1, 3], vertical_alignment="center")
    brand_icon.image("assets/matchmind_star.png", width=54)
    brand_name.markdown("### MatchMind AI")
    st.caption("Candidate résumé workspace" if candidate_mode else "Recruiter screening workspace")
    st.markdown(f'<span class="role-badge">{app_role}</span>', unsafe_allow_html=True)
    st.divider()
    st.write(f"Signed in as **{display_name}**")
    if user.get("auth_provider") == "google":
        st.caption("Google account")
    if st.button("Switch role", use_container_width=True):
        switch_app_role()
    if st.button("Sign out", use_container_width=True):
        sign_out()
    st.divider()
    st.subheader("Résumé-check history" if candidate_mode else "Screening history")
    history = list_analyses(user["id"])
    selected_history = st.selectbox(
        "Open a saved check" if candidate_mode else "Open a saved shortlist",
        options=[None] + [item["id"] for item in history],
        format_func=lambda value: analysis_option_label(value, history),
    )
    if selected_history and st.button("Load analysis", use_container_width=True):
        loaded_results, loaded_context = load_analysis_bundle(user["id"], selected_history)
        saved_role = (loaded_context or {}).get("app_role")
        if saved_role in APP_ROLES:
            st.session_state.app_role = saved_role
        st.session_state.loaded_results = loaded_results
        st.session_state.analysis_context = loaded_context
        st.session_state.pop("tailored_files", None)
        st.session_state.pop("draft_job_title", None)
        st.session_state.pop("draft_job_description", None)
        st.session_state.result_view = "ranking"
        st.rerun()
    if st.session_state.get("loaded_results") is not None:
        new_label = "Start new résumé check" if candidate_mode else "Start new screening"
        if st.button(new_label, type="primary", use_container_width=True):
            clear_workflow_state()
            st.rerun()

st.markdown(
    f"""
    <div class="cobalt-brand">
      <span class="cobalt-mark"><img src="{LOGO_DATA_URI}" alt="MatchMind AI"></span>
      <div><div class="cobalt-brand-title">MatchMind AI</div>
      <div class="cobalt-brand-subtitle">{'Candidate résumé enhancement' if candidate_mode else 'Explainable candidate screening'}</div></div>
    </div>
    """,
    unsafe_allow_html=True,
)

current_view = st.session_state.get("result_view", "ranking")
if st.session_state.get("loaded_results") is None and current_view not in {"edit_role", "edit_resumes"}:
    current_view = "edit_role"
    st.session_state.result_view = "edit_role"
editing_inputs = current_view in {"edit_role", "edit_resumes"}
if st.session_state.get("loaded_results") is None or editing_inputs:
    existing_context = st.session_state.get("analysis_context") or {}
    hero_title = "Check and enhance your résumé" if candidate_mode else "Screen and compare candidates"
    hero_copy = (
        "Compare your résumé with a job, verify missing evidence and create a stronger job-focused version."
        if candidate_mode
        else "Rank multiple applicants and understand clearly why each candidate is or is not suitable."
    )
    st.markdown(
        f'<div class="cobalt-hero"><h1>{hero_title}</h1><p>{hero_copy}</p></div>',
        unsafe_allow_html=True,
    )
    active_input_stage = 2 if current_view == "edit_resumes" else 1
    show_stage_header(active_input_stage, app_role)

    if active_input_stage == 1:
        st.markdown('<div class="cobalt-step">01 · Add the job</div>', unsafe_allow_html=True)
        st.markdown('<div class="cobalt-panel-title">Job information</div>', unsafe_allow_html=True)
        st.markdown('<div class="cobalt-panel-copy">Job title and job description are mandatory. Saved information appears here when you return.</div>', unsafe_allow_html=True)
        job_title = st.text_input(
            "Job title / reference *",
            value=st.session_state.get("draft_job_title", existing_context.get("job_title", "")),
            placeholder="Senior Data Analyst · DA-204",
        )
        jd_file = st.file_uploader("Upload job description", type=["pdf", "docx", "txt"], key="role_jd_upload")
        jd_text = st.text_area(
            "Or paste job description *",
            value=st.session_state.get("draft_job_description", existing_context.get("job_description", "")),
            height=300,
        )
        continue_label = "Save job and continue to my résumé" if candidate_mode else "Save job and continue to résumés"
        if st.button(continue_label, type="primary", use_container_width=True):
            try:
                if not job_title.strip():
                    raise ValueError("Job title / reference is mandatory.")
                saved_jd, used_ocr = extract_document_text(jd_file, load_ocr_reader) if jd_file else (jd_text, False)
                if not saved_jd.strip():
                    raise ValueError("Job description is mandatory.")
                st.session_state.draft_job_title = job_title.strip()
                st.session_state.draft_job_description = saved_jd
                if existing_context:
                    updated_context = dict(existing_context)
                    updated_context.update({"job_title": job_title.strip(), "job_description": saved_jd, "app_role": app_role})
                    st.session_state.analysis_context = updated_context
                if used_ocr:
                    st.session_state.role_ocr_notice = jd_file.name
                st.session_state.result_view = "edit_resumes"
                st.rerun()
            except Exception as error:
                st.error(str(error))
    else:
        edited_resume_texts = {}
        job_title = st.session_state.get("draft_job_title", existing_context.get("job_title", ""))
        final_jd = st.session_state.get("draft_job_description", existing_context.get("job_description", ""))
        if not job_title or not final_jd:
            st.warning("Complete the mandatory Job step first.")
            if st.button("Go to Job", type="primary"):
                st.session_state.result_view = "edit_role"
                st.rerun()
            st.stop()
        st.info(f"Job ready: {job_title}")
        st.markdown('<div class="cobalt-step">02 · Add résumé</div>' if candidate_mode else '<div class="cobalt-step">02 · Add candidates</div>', unsafe_allow_html=True)
        st.markdown('<div class="cobalt-panel-title">Your résumé</div>' if candidate_mode else '<div class="cobalt-panel-title">Candidate résumés</div>', unsafe_allow_html=True)
        upload_copy = "Upload one PDF, DOCX, TXT or scanned-PDF résumé." if candidate_mode else "Upload multiple PDF, DOCX, TXT or scanned-PDF résumés."
        st.markdown(f'<div class="cobalt-panel-copy">{upload_copy}</div>', unsafe_allow_html=True)
        if candidate_mode:
            resume_upload = st.file_uploader("Upload your résumé", type=["pdf", "docx", "txt"], accept_multiple_files=False)
            resume_files = [resume_upload] if resume_upload else []
        else:
            resume_files = st.file_uploader(
                "Upload multiple résumés",
                type=["pdf", "docx", "txt"],
                accept_multiple_files=True,
            ) or []
        existing_resumes = existing_context.get("resume_texts", {})
        if candidate_mode and existing_resumes:
            existing_resumes = dict(list(existing_resumes.items())[:1])
        if existing_resumes:
            st.success("Currently saved: " + ", ".join(existing_resumes))
            st.caption("Open the saved résumé text below to review or correct it. Upload a new file only to replace it." if candidate_mode else "Open a saved résumé below to review or correct its extracted text. Upload new files only to replace the set.")
            for saved_name, saved_text in existing_resumes.items():
                with st.expander(f"Review / edit · {saved_name}"):
                    edited_resume_texts[saved_name] = st.text_area(
                        f"Extracted résumé text for {saved_name}",
                        value=saved_text,
                        height=220,
                        key=f"resume_editor_{saved_name}",
                    )
        st.caption(
            "Scanned PDFs automatically use OCR. Extracted text is saved privately with the analysis "
            "so the workflow still works after you reopen it."
        )
        back_column, continue_column = st.columns([1, 3])
        with back_column:
            back_to_role = st.button("← Back to Job", use_container_width=True)
        with continue_column:
            if existing_context:
                action_label = "Update and check again" if candidate_mode else "Update and rerun screening"
            else:
                action_label = "Check my résumé →" if candidate_mode else "Run candidate screening →"
            run_matching = st.button(action_label, type="primary", use_container_width=True)
        if back_to_role:
            st.session_state.result_view = "edit_role"
            st.rerun()
        if run_matching:
            status_label = "Checking your résumé…" if candidate_mode else "Loading and analysing résumés…"
            analysis_status = st.status(status_label, expanded=True)
            try:
                analysis_status.write("Checking the job and uploaded file…" if candidate_mode else "Checking the job and uploaded files…")
                if not resume_files and not existing_resumes:
                    raise ValueError("Upload your résumé before checking it." if candidate_mode else "Upload at least one résumé before screening.")
                seen_hashes, resumes, skipped, ocr_files, original_files = set(), [], [], [], {}
                if resume_files:
                    for file in resume_files:
                        analysis_status.write(f"Reading {file.name}…")
                        content = file.getvalue()
                        digest = hashlib.sha256(content).hexdigest()
                        if digest in seen_hashes:
                            skipped.append(f"{file.name} (duplicate)")
                            continue
                        seen_hashes.add(digest)
                        text, used_ocr = extract_document_text(file, load_ocr_reader)
                        candidate = file.name.rsplit(".", 1)[0]
                        if used_ocr:
                            ocr_files.append(file.name)
                        if not text.strip():
                            skipped.append(f"{file.name} (no readable text after extraction)")
                        else:
                            resumes.append((candidate, text))
                            original_files[candidate] = {
                                "name": file.name,
                                "mime": getattr(file, "type", None) or "application/octet-stream",
                                "content_b64": base64.b64encode(content).decode("ascii"),
                            }
                else:
                    resumes = list((edited_resume_texts or existing_resumes).items())
                    original_files = existing_context.get("original_files", {})
                if candidate_mode:
                    resumes = resumes[:1]
                    original_files = {name: original_files[name] for name, _ in resumes if name in original_files}
                if not resumes:
                    raise ValueError("The uploaded résumé did not contain readable text." if candidate_mode else "None of the uploaded résumés contained readable text.")

                analysis_status.write("Loading the semantic AI model…")
                model = load_semantic_model()
                analysis_status.write("Comparing the résumé with the job requirements…" if candidate_mode else "Calculating job-description, skill and experience matches…")
                results = rank_resumes_advanced(
                    final_jd,
                    resumes,
                    lambda texts: model.encode(texts, normalize_embeddings=True, show_progress_bar=False),
                )
                context = {
                    "job_description": final_jd,
                    "job_title": job_title,
                    "resume_texts": dict(resumes),
                    "original_files": original_files,
                    "app_role": app_role,
                }
                analysis_id = save_analysis(user["id"], job_title, results, context)
                st.session_state.loaded_results = results
                st.session_state.analysis_context = context
                st.session_state.tailored_files = {}
                st.session_state.result_view = "ranking"
                analysis_status.update(label="Résumé check completed" if candidate_mode else "Screening completed", state="complete", expanded=False)
                if ocr_files:
                    st.info("OCR used for: " + ", ".join(ocr_files))
                if skipped:
                    st.warning("Skipped: " + "; ".join(skipped))
                st.success(f"Analysis saved with ID {analysis_id}.")
                st.rerun()
            except Exception as error:
                analysis_status.update(label="Analysis could not be completed", state="error", expanded=True)
                st.error(str(error))
else:
    if st.session_state.get("result_view", "ranking") == "tailoring":
        if candidate_mode:
            show_resume_tailoring(st.session_state.loaded_results)
        else:
            st.session_state.result_view = "ranking"
            st.rerun()
    elif candidate_mode:
        show_candidate_results(st.session_state.loaded_results, st.session_state.get("analysis_context"))
    else:
        show_recruiter_results(st.session_state.loaded_results, st.session_state.get("analysis_context"))

if candidate_mode:
    st.caption("Résumé guidance only. Verify every statement before submitting an application.")
else:
    st.caption("Decision-support only. Human review is required before any hiring decision.")
