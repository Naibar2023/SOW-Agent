"""SOW Agent as a web page (works in any browser, including on a phone).

    streamlit run streamlit_app.py

Settings come from the .env file on your own computer, or from the app's Secrets
when it is deployed on Streamlit Community Cloud:

    ANTHROPIC_API_KEY = "sk-ant-..."
    APP_PASSWORD      = "choose-a-password"     (asked before anyone can use the page)
    SOW_MODEL         = "claude-opus-5-5"       (optional)
"""

import hmac
import os
import re
from pathlib import Path

import streamlit as st

from sa_agent import AgentError, SOWAgent
from sa_config import load_env, workspace_dir

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
UPLOAD_TYPES = ["pdf", "docx", "xlsx", "png", "jpg", "jpeg", "txt", "csv"]
SETTING_KEYS = ("ANTHROPIC_API_KEY", "APP_PASSWORD", "SOW_MODEL", "SOW_WEB_SEARCH", "SOW_MAX_TOKENS")

st.set_page_config(page_title="SOW Agent", page_icon="📄", layout="centered")


# --------------------------------------------------------------------------- settings

def load_settings():
    load_env()
    try:                                   # Streamlit secrets (cloud, or .streamlit/secrets.toml)
        for key in SETTING_KEYS:
            if key in st.secrets and not os.environ.get(key):
                os.environ[key] = str(st.secrets[key])
    except Exception:
        pass                               # no secrets file: the .env file is enough


def check_password():
    """Ask for the password once per browser session. Returns True when access is allowed."""
    password = os.environ.get("APP_PASSWORD", "")
    if not password:
        return True
    if st.session_state.get("unlocked"):
        return True
    st.title("SOW Agent")
    with st.form("login"):
        entered = st.text_input("Password", type="password")
        ok = st.form_submit_button("Enter")
    if ok:
        if hmac.compare_digest(entered.encode("utf-8"), password.encode("utf-8")):
            st.session_state["unlocked"] = True
            st.rerun()
        st.error("Wrong password.")
    return False


# --------------------------------------------------------------------------- helpers

def show_text(text):
    """Model replies are plain text; keep line breaks and stop '$' from turning into formulas."""
    return (text or "").replace("$", "\\$").replace("\n", "  \n")


_ATTACHED = re.compile(r'<attached_file name="([^"]*)">.*?</attached_file>\s*', re.S)
_GONE = re.compile(r"\[The user attached (.+?) here\..*?\]\s*", re.S)


def history_for_display(agent):
    """Rebuild the visible conversation from a saved project."""
    out = []
    for h in agent.state["history"]:
        text = _ATTACHED.sub(lambda m: f"(file: {m.group(1)})\n", h["text"])
        text = _GONE.sub(lambda m: f"(file: {m.group(1)})\n", text)
        out.append({"role": h["role"], "text": text.strip(), "files": []})
    return out


def file_buttons(paths, prefix):
    for i, p in enumerate(paths):
        p = Path(p)
        if p.is_file():
            st.download_button(f"Download {p.name}", data=p.read_bytes(), file_name=p.name,
                               mime=DOCX_MIME, key=f"{prefix}_{i}_{p.name}")


def safe_name(name):
    keep = "".join(c if c.isalnum() or c in "._- " else "_" for c in (name or "file"))
    return keep.strip(" .")[:120] or "file"


# --------------------------------------------------------------------------- page

load_settings()
if not check_password():
    st.stop()

if "agent" not in st.session_state:
    st.session_state["agent"] = SOWAgent(workspace_dir())
    st.session_state["msgs"] = history_for_display(st.session_state["agent"])
agent = st.session_state["agent"]
msgs = st.session_state["msgs"]

with st.sidebar:
    st.header("SOW Agent")
    st.caption(f"Project: {agent.state['name']}  ·  V{agent.state['version']}")

    latest = agent.latest_file()
    if latest:
        st.download_button("Download latest Word file", data=latest.read_bytes(), file_name=latest.name,
                           mime=DOCX_MIME, key="latest_file")

    st.divider()
    new_name = st.text_input("New project name (optional)", key="new_name")
    if st.button("Start new SOW"):
        agent.new_project(new_name.strip() or None)
        st.session_state["msgs"] = []
        st.rerun()

    projects = agent.list_projects()
    if projects:
        st.divider()
        labels = {pid: f"{name} (V{ver})" for pid, name, ver, _ in projects}
        ids = list(labels)
        current = ids.index(agent.project_id) if agent.project_id in ids else 0
        chosen = st.selectbox("Saved projects", ids, index=current, format_func=lambda pid: labels[pid])
        if st.button("Open project") and chosen != agent.project_id:
            agent.open_project(chosen)
            st.session_state["msgs"] = history_for_display(agent)
            st.rerun()

    if agent.state["versions"]:
        st.divider()
        st.caption("Versions")
        for v in reversed(agent.state["versions"][-8:]):
            st.caption(f"V{v['version']} · {v['time']} · {v['note'] or 'saved'}")

    if not os.environ.get("ANTHROPIC_API_KEY"):
        st.error("ANTHROPIC_API_KEY is not set (add it to .env or to the app Secrets).")
    if not os.environ.get("APP_PASSWORD"):
        st.warning("No APP_PASSWORD is set. Set one before putting this page online.")

if not msgs:
    st.title("SOW Agent")
    st.write("Describe the lighting or electrical job you need a Scope of Work for. "
             "You can attach drawings, a BOQ or an old SOW with the paperclip.")

for n, m in enumerate(msgs):
    with st.chat_message(m["role"]):
        st.markdown(show_text(m["text"]))
        file_buttons(m.get("files", []), f"hist{n}")

prompt = st.chat_input("Describe the job, or ask for a change...", accept_file="multiple",
                       file_type=UPLOAD_TYPES)

if prompt:
    text = (prompt.text or "").strip()
    saved = []
    for up in prompt.files or []:
        folder = agent.project_dir / "received"
        folder.mkdir(parents=True, exist_ok=True)
        path = folder / safe_name(up.name)
        path.write_bytes(up.getvalue())
        saved.append(path)

    shown = "\n".join([f"(file: {p.name})" for p in saved] + ([text] if text else []))
    msgs.append({"role": "user", "text": shown, "files": []})
    with st.chat_message("user"):
        st.markdown(show_text(shown))

    with st.chat_message("assistant"):
        status = st.empty()
        box = st.empty()
        buf = []

        def on_text(chunk):
            buf.append(chunk)
            box.markdown(show_text("".join(buf)))

        def on_status(note):
            status.caption(note)
            if buf and not buf[-1].endswith("\n\n"):
                buf.append("\n\n")

        agent.on_text, agent.on_status = on_text, on_status
        status.caption("Working...")
        try:
            reply = agent.send(text, saved)
            reply_text, files = reply.text, [str(f) for f in reply.files]
        except AgentError as e:
            reply_text, files = f"Problem: {e}", []
        finally:
            agent.on_text = agent.on_status = lambda *_: None
        status.empty()
        box.markdown(show_text(reply_text))
        file_buttons(files, f"new{len(msgs)}")

    msgs.append({"role": "assistant", "text": reply_text, "files": files})
    if files:
        st.rerun()        # refresh the sidebar (version list and latest file)
