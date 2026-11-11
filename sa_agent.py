"""The SOW agent: conversation loop, tools, projects and versions.

One SOWAgent object serves one user. It keeps a workspace folder:

    workspace/projects/<project-id>/state.json        conversation and current SOW
    workspace/projects/<project-id>/SOW_<title>_V1.docx, _V2.docx, ...
    workspace/projects/<project-id>/sow_v1.json, ...  the data behind each version
"""

from __future__ import annotations

import base64
import copy
import datetime as _dt
import json
import os
import re
from dataclasses import dataclass, field
from pathlib import Path

import sa_prompts as prompts
from sa_renderer import SOWValidationError, render_sow, validate_sow

DEFAULT_MODEL = "claude-opus-5-5"
MAX_LOOPS = 12                 # safety limit on tool rounds within one user message
MAX_TEXT_ATTACHMENT = 150_000  # characters
MAX_BINARY_ATTACHMENT = 20 * 1024 * 1024

IMAGE_TYPES = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
               ".gif": "image/gif", ".webp": "image/webp"}
TEXT_TYPES = {".txt", ".md", ".csv", ".json", ".log", ".tsv"}


@dataclass
class Reply:
    text: str
    files: list[Path] = field(default_factory=list)


class AgentError(Exception):
    """A problem the user can act on (missing key, no credit, connection...)."""


# --------------------------------------------------------------------------- tools given to the model

TOOLS = [
    {
        "name": "save_sow",
        "description": (
            "Create the Word (.docx) SOW in the house format from a complete SOW JSON object and send the "
            "file to the user. Use for the first draft of a SOW or when most sections are rewritten. "
            "Each call creates a new version. For small changes use edit_sow instead."),
        "input_schema": {
            "type": "object",
            "properties": {
                "sow": {
                    "type": "object",
                    "description": "The complete SOW in the SOW JSON format described in the system prompt.",
                    "properties": {
                        "cover_title": {"type": "string"},
                        "document_title": {"type": "string"},
                        "sections": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "heading": {"type": "string"},
                                    "blocks": {"type": "array", "items": {"type": "object"}},
                                },
                                "required": ["heading", "blocks"],
                            },
                        },
                    },
                    "required": ["cover_title", "sections"],
                },
                "change_note": {"type": "string", "description": "One line describing this version."},
            },
            "required": ["sow"],
        },
    },
    {
        "name": "edit_sow",
        "description": (
            "Change the current SOW draft without resending all of it, then rebuild the Word file as a new "
            "version and send it to the user. Operations are applied in order. Section headings are given "
            "without their letter."),
        "input_schema": {
            "type": "object",
            "properties": {
                "operations": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "op": {"type": "string",
                                   "enum": ["replace_section", "insert_section", "delete_section",
                                            "rename_section", "set_cover_title"]},
                            "heading": {"type": "string",
                                        "description": "Existing section heading (replace, delete, rename)."},
                            "section": {"type": "object",
                                        "description": "Full new section {heading, blocks} (replace, insert)."},
                            "after": {"type": "string",
                                      "description": "insert_section: heading of the section to insert after. "
                                                     "Omit to add at the end; use \"START\" for the beginning."},
                            "new_heading": {"type": "string", "description": "rename_section: the new heading."},
                            "cover_title": {"type": "string", "description": "set_cover_title: the new title."},
                        },
                        "required": ["op"],
                    },
                },
                "change_note": {"type": "string", "description": "One line describing what changed."},
            },
            "required": ["operations"],
        },
    },
    {
        "name": "restore_version",
        "description": "Make an earlier version of the SOW the current draft again. It is saved as a new version.",
        "input_schema": {
            "type": "object",
            "properties": {"version": {"type": "integer", "description": "Version number to restore, e.g. 2."}},
            "required": ["version"],
        },
    },
]

WEB_SEARCH_TOOL = {"type": "web_search_20250305", "name": "web_search", "max_uses": 5}


# --------------------------------------------------------------------------- helpers

def _slug(text, limit=70):
    s = re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_")
    return s[:limit].rstrip("_") or "Untitled"


def _norm_heading(h):
    h = re.sub(r"^[A-Z]{1,2}[.)]\s+", "", str(h).strip())
    return h.rstrip(":").strip().lower()


def _now():
    return _dt.datetime.now().strftime("%Y-%m-%d %H:%M")


def _block_attr(block, name, default=None):
    if isinstance(block, dict):
        return block.get(name, default)
    return getattr(block, name, default)


def read_attachment(path: Path):
    """Turn a file into (content_blocks, note). Text-like files become text; PDFs and images are sent as-is.

    Returns (blocks, text_for_history, error). Exactly one of blocks/error is set.
    """
    path = Path(path)
    if not path.is_file():
        return None, None, f"File not found: {path}"
    ext = path.suffix.lower()
    name = path.name
    try:
        if ext in IMAGE_TYPES or ext == ".pdf":
            data = path.read_bytes()
            if len(data) > MAX_BINARY_ATTACHMENT:
                return None, None, f"{name} is larger than 20 MB."
            b64 = base64.standard_b64encode(data).decode("ascii")
            if ext == ".pdf":
                blk = {"type": "document", "title": name,
                       "source": {"type": "base64", "media_type": "application/pdf", "data": b64}}
            else:
                blk = {"type": "image",
                       "source": {"type": "base64", "media_type": IMAGE_TYPES[ext], "data": b64}}
            return [{"type": "text", "text": f"Attached file: {name}"}, blk], None, None
        if ext == ".docx":
            from docx import Document
            doc = Document(str(path))
            lines = []
            for el in doc.element.body:
                tag = el.tag.split("}")[1]
                if tag == "p":
                    t = "".join(x.text or "" for x in el.iter() if x.tag.endswith("}t")).strip()
                    if t:
                        lines.append(t)
                elif tag == "tbl":
                    for tr in (x for x in el if x.tag.endswith("}tr")):
                        cells = []
                        for tc in (x for x in tr if x.tag.endswith("}tc")):
                            cells.append(" ".join("".join(y.text or "" for y in tc.iter()
                                                          if y.tag.endswith("}t")).split()))
                        lines.append(" | ".join(cells))
            text = "\n".join(lines)
        elif ext in (".xlsx", ".xlsm"):
            try:
                import openpyxl
            except ImportError:
                return None, None, (f"Cannot read {name}: Excel support is not installed. "
                                    "Run:  pip install openpyxl   (or save the sheet as CSV).")
            wb = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
            parts = []
            for ws in wb.worksheets:
                parts.append(f"## Sheet: {ws.title}")
                for row in ws.iter_rows(values_only=True):
                    if any(c is not None and str(c).strip() for c in row):
                        parts.append(" | ".join("" if c is None else str(c) for c in row))
            text = "\n".join(parts)
        elif ext in TEXT_TYPES:
            text = path.read_text(encoding="utf-8", errors="replace")
        else:
            return None, None, (f"Cannot read {name}: supported files are PDF, Word (.docx), Excel (.xlsx), "
                                "images (png, jpg) and text (txt, csv).")
    except Exception as e:  # unreadable or corrupt file
        return None, None, f"Cannot read {name}: {e}"
    if not text.strip():
        return None, None, f"{name} contains no readable text."
    if len(text) > MAX_TEXT_ATTACHMENT:
        text = text[:MAX_TEXT_ATTACHMENT] + "\n[... file truncated ...]"
    wrapped = f'<attached_file name="{name}">\n{text}\n</attached_file>'
    return [{"type": "text", "text": wrapped}], wrapped, None


# --------------------------------------------------------------------------- the agent

class SOWAgent:
    def __init__(self, workspace, client=None, model=None, web_search=None, on_text=None, on_status=None):
        self.workspace = Path(workspace)
        self.projects_dir = self.workspace / "projects"
        self.projects_dir.mkdir(parents=True, exist_ok=True)
        self.model = model or os.environ.get("SOW_MODEL") or DEFAULT_MODEL
        self.max_tokens = int(os.environ.get("SOW_MAX_TOKENS") or 32000)
        if web_search is None:
            web_search = (os.environ.get("SOW_WEB_SEARCH", "1").strip().lower() not in ("0", "no", "false", "off"))
        self.web_search = web_search
        self.on_text = on_text or (lambda s: None)       # live text as the model writes
        self.on_status = on_status or (lambda s: None)   # short progress notes
        self._client = client
        self._static_prompt = prompts.static_system_prompt()
        self.project_id = None
        self.state = None
        self._live_content = {}   # history index -> full content blocks (PDFs/images), this session only
        self._turn_files: list[Path] = []
        last = self.workspace / "last_project.txt"
        if last.is_file() and (self.projects_dir / last.read_text().strip() / "state.json").is_file():
            self.open_project(last.read_text().strip())
        else:
            self.new_project()

    # ---- projects

    @property
    def project_dir(self) -> Path:
        return self.projects_dir / self.project_id

    def new_project(self, name=None):
        stamp = _dt.datetime.now().strftime("%Y%m%d-%H%M%S")
        self.project_id = f"{stamp}_{_slug(name, 40)}" if name else stamp
        self.state = {"name": name or "New SOW", "created": _now(), "history": [], "sow": None,
                      "version": 0, "versions": []}
        self._live_content = {}
        self._remember()
        return self.project_id

    def list_projects(self):
        """Saved projects, newest first: [(project_id, name, version, updated)]."""
        out = []
        for d in sorted(self.projects_dir.iterdir(), reverse=True):
            f = d / "state.json"
            if f.is_file():
                try:
                    st = json.loads(f.read_text(encoding="utf-8"))
                    out.append((d.name, st.get("name", d.name), st.get("version", 0), st.get("updated", "")))
                except Exception:
                    continue
        return out

    def open_project(self, project_id):
        f = self.projects_dir / project_id / "state.json"
        if not f.is_file():
            raise AgentError(f"No saved project called {project_id}.")
        self.state = json.loads(f.read_text(encoding="utf-8"))
        self.project_id = project_id
        self._live_content = {}
        self._remember()

    def _remember(self):
        (self.workspace / "last_project.txt").write_text(self.project_id, encoding="utf-8")

    def _save_state(self):
        self.project_dir.mkdir(parents=True, exist_ok=True)
        self.state["updated"] = _now()
        tmp = self.project_dir / "state.json.tmp"
        tmp.write_text(json.dumps(self.state, ensure_ascii=False, indent=1), encoding="utf-8")
        tmp.replace(self.project_dir / "state.json")
        self._remember()

    def latest_file(self):
        if self.state["versions"]:
            p = self.project_dir / self.state["versions"][-1]["file"]
            return p if p.is_file() else None
        return None

    # ---- tool implementations

    def _commit(self, sow, note):
        """Validate, render and record a new version. Returns the tool result text."""
        errs = validate_sow(sow)
        if errs:
            raise SOWValidationError("\n".join(errs))
        version = self.state["version"] + 1
        fname = f"SOW_{_slug(sow['cover_title'])}_V{version}.docx"
        self.project_dir.mkdir(parents=True, exist_ok=True)
        path = render_sow(sow, self.project_dir / fname)
        (self.project_dir / f"sow_v{version}.json").write_text(
            json.dumps(sow, ensure_ascii=False, indent=1), encoding="utf-8")
        self.state["sow"] = sow
        self.state["version"] = version
        self.state["name"] = sow["cover_title"]
        self.state["versions"].append({"version": version, "file": fname, "note": (note or "").strip(),
                                       "time": _now()})
        self._save_state()
        self._turn_files.append(path)
        self.on_status(f"Word file ready: {fname}")
        return (f"Saved version V{version} as {fname} ({len(sow['sections'])} sections). "
                "The Word file has been sent to the user.")

    def _tool_save_sow(self, args):
        sow = args.get("sow")
        if isinstance(sow, str):          # tolerate a JSON string
            sow = json.loads(sow)
        return self._commit(sow, args.get("change_note"))

    def _tool_edit_sow(self, args):
        if not self.state["sow"]:
            raise SOWValidationError("There is no SOW draft yet. Create one with save_sow first.")
        sow = copy.deepcopy(self.state["sow"])
        secs = sow["sections"]

        def find(heading):
            key = _norm_heading(heading or "")
            for i, s in enumerate(secs):
                if _norm_heading(s["heading"]) == key:
                    return i
            for i, s in enumerate(secs):
                if key and (key in _norm_heading(s["heading"]) or _norm_heading(s["heading"]) in key):
                    return i
            names = "; ".join(s["heading"] for s in secs)
            raise SOWValidationError(f'No section called "{heading}". Sections are: {names}')

        ops = args.get("operations")
        if not isinstance(ops, list) or not ops:
            raise SOWValidationError('"operations" must be a non-empty list.')
        for n, op in enumerate(ops, 1):
            kind = op.get("op")
            if kind == "replace_section":
                if not isinstance(op.get("section"), dict):
                    raise SOWValidationError(f'Operation {n}: replace_section needs "section".')
                secs[find(op.get("heading") or op["section"].get("heading"))] = op["section"]
            elif kind == "insert_section":
                if not isinstance(op.get("section"), dict):
                    raise SOWValidationError(f'Operation {n}: insert_section needs "section".')
                after = op.get("after")
                if after is None or after == "":
                    secs.append(op["section"])
                elif str(after).strip().upper() == "START":
                    secs.insert(0, op["section"])
                else:
                    secs.insert(find(after) + 1, op["section"])
            elif kind == "delete_section":
                del secs[find(op.get("heading"))]
            elif kind == "rename_section":
                if not op.get("new_heading"):
                    raise SOWValidationError(f'Operation {n}: rename_section needs "new_heading".')
                secs[find(op.get("heading"))]["heading"] = op["new_heading"]
            elif kind == "set_cover_title":
                if not op.get("cover_title"):
                    raise SOWValidationError(f'Operation {n}: set_cover_title needs "cover_title".')
                sow["cover_title"] = op["cover_title"]
            else:
                raise SOWValidationError(f'Operation {n}: unknown op "{kind}".')
        return self._commit(sow, args.get("change_note"))

    def _tool_restore_version(self, args):
        v = args.get("version")
        f = self.project_dir / f"sow_v{v}.json"
        if not f.is_file():
            have = ", ".join(f"V{x['version']}" for x in self.state["versions"]) or "none"
            raise SOWValidationError(f"Version {v} does not exist. Available: {have}.")
        return self._commit(json.loads(f.read_text(encoding="utf-8")), f"Restored from V{v}")

    def _run_tool(self, name, args):
        try:
            if name == "save_sow":
                return self._tool_save_sow(args), False
            if name == "edit_sow":
                return self._tool_edit_sow(args), False
            if name == "restore_version":
                return self._tool_restore_version(args), False
            return f"Unknown tool: {name}", True
        except SOWValidationError as e:
            return "The SOW was NOT saved. Fix these problems and call the tool again:\n" + str(e), True
        except Exception as e:
            return f"The SOW was NOT saved because of an internal error: {type(e).__name__}: {e}", True

    # ---- talking to the model

    def _get_client(self):
        if self._client is None:
            if not os.environ.get("ANTHROPIC_API_KEY"):
                raise AgentError("No API key found. Put your key in the .env file as "
                                 "ANTHROPIC_API_KEY=sk-ant-...  (see README, step 2).")
            try:
                import anthropic
            except ImportError:
                raise AgentError("The 'anthropic' package is not installed. Run setup again "
                                 "(pip install -r requirements.txt).")
            self._client = anthropic.Anthropic()
        return self._client

    def _system(self):
        st = self.state
        dyn = [f"Today's date: {_dt.date.today().isoformat()}.", f"Current project: {st['name']}."]
        if st["versions"]:
            dyn.append("Versions saved so far (the user has received each file):")
            for v in st["versions"][-15:]:
                dyn.append(f"- V{v['version']}  {v['file']}  {v['time']}  {v['note']}".rstrip())
            dyn.append(f"\nCurrent SOW draft (V{st['version']}) in SOW JSON format. edit_sow operations apply to this:")
            dyn.append(json.dumps(st["sow"], ensure_ascii=False, separators=(",", ":")))
        else:
            dyn.append("No SOW has been saved in this project yet.")
        return [
            {"type": "text", "text": self._static_prompt, "cache_control": {"type": "ephemeral"}},
            {"type": "text", "text": "\n".join(dyn), "cache_control": {"type": "ephemeral"}},
        ]

    def _history_messages(self):
        msgs = []
        for i, h in enumerate(self.state["history"]):
            content = self._live_content.get(i) or h["text"]
            msgs.append({"role": h["role"], "content": content})
        return msgs

    def _call(self, messages):
        client = self._get_client()
        tools = list(TOOLS) + ([WEB_SEARCH_TOOL] if self.web_search else [])
        with client.messages.stream(model=self.model, max_tokens=self.max_tokens, system=self._system(),
                                    tools=tools, messages=messages) as stream:
            for chunk in stream.text_stream:
                self.on_text(chunk)
            return stream.get_final_message()

    def send(self, text, attachments=()) -> Reply:
        """Send one user message (optionally with files) and return the agent's reply."""
        text = (text or "").strip()
        blocks, hist_parts, problems, binary = [], [], [], False
        for a in attachments:
            b, hist, err = read_attachment(Path(a))
            if err:
                problems.append(err)
                continue
            blocks.extend(b)
            if hist:
                hist_parts.append(hist)
            else:
                binary = True
                hist_parts.append(f"[The user attached {Path(a).name} here. It is no longer loaded; "
                                  "ask the user to attach it again if you need to read it.]")
        if problems and not blocks and not text:
            return Reply("\n".join(problems))
        if not text and not blocks:
            return Reply("")
        if not text:
            text = "Please read the attached file(s) and tell me briefly what they are and how you can use them."
        user_content = blocks + [{"type": "text", "text": text}] if blocks else text
        hist_text = "\n\n".join(hist_parts + [text])

        messages = self._history_messages() + [{"role": "user", "content": user_content}]
        self._turn_files = []
        said = []
        cut_retries = 0
        try:
            for _ in range(MAX_LOOPS):
                msg = self._call(messages)
                content = list(msg.content)
                turn_text = "".join(_block_attr(b, "text", "") or "" for b in content
                                    if _block_attr(b, "type") == "text")
                if turn_text.strip():
                    said.append(turn_text.strip())
                stop = msg.stop_reason

                if stop == "max_tokens":
                    # The reply was cut off. Drop any unfinished tool call and ask for a split delivery.
                    keep = [b for b in content if _block_attr(b, "type") == "text"
                            and (_block_attr(b, "text") or "").strip()]
                    cut_retries += 1
                    if cut_retries > 2:
                        said.append("(The reply was too long and was cut off. Ask me to continue, or to "
                                    "build the SOW in two parts.)")
                        break
                    messages.append({"role": "assistant",
                                     "content": keep or [{"type": "text", "text": "(reply cut off)"}]})
                    messages.append({"role": "user", "content": (
                        "[System note: your last reply hit the length limit and was cut off, so nothing was "
                        "saved. If you were sending a SOW, send it in two steps: save_sow with the first half "
                        "of the sections, then edit_sow with insert_section operations for the rest.]")})
                    self.on_status("Reply was too long; asking the agent to send it in parts...")
                    continue

                messages.append({"role": "assistant", "content": content})
                if stop == "tool_use":
                    results = []
                    for b in content:
                        if _block_attr(b, "type") == "tool_use":
                            name = _block_attr(b, "name")
                            self.on_status({"save_sow": "Building the Word file...",
                                            "edit_sow": "Updating the Word file...",
                                            "restore_version": "Restoring the version..."}.get(name, name))
                            out, is_err = self._run_tool(name, _block_attr(b, "input") or {})
                            res = {"type": "tool_result", "tool_use_id": _block_attr(b, "id"), "content": out}
                            if is_err:
                                res["is_error"] = True
                            results.append(res)
                    messages.append({"role": "user", "content": results})
                    continue
                if stop == "pause_turn":      # long web search: send the turn back unchanged to continue
                    continue
                break
            else:
                said.append("(Stopped after too many steps. Tell me how to continue.)")
        except AgentError:
            raise
        except Exception as e:
            raise AgentError(_friendly_error(e)) from e

        reply_text = "\n\n".join(said).strip()
        if not reply_text:
            reply_text = "The SOW file is ready." if self._turn_files else "(no reply)"
        if problems:
            reply_text = "\n".join(problems) + "\n\n" + reply_text

        hist = self.state["history"]
        hist.append({"role": "user", "text": hist_text})
        if binary:
            self._live_content[len(hist) - 1] = user_content
        hist.append({"role": "assistant", "text": reply_text})
        self._trim_history()
        self._save_state()
        return Reply(reply_text, list(self._turn_files))

    def _trim_history(self, keep=60):
        """Keep the conversation from growing without limit; the SOW itself is stored separately."""
        hist = self.state["history"]
        if len(hist) <= keep:
            return
        drop = len(hist) - keep
        drop += drop % 2                       # always drop whole user/assistant pairs
        self.state["history"] = hist[drop:]
        self._live_content = {i - drop: c for i, c in self._live_content.items() if i >= drop}


def _friendly_error(e):
    name = type(e).__name__
    status = getattr(e, "status_code", None)
    detail = ""
    body = getattr(e, "body", None)
    if isinstance(body, dict):
        detail = (body.get("error") or {}).get("message", "") if isinstance(body.get("error"), dict) else ""
    detail = detail or str(e)
    if name == "AuthenticationError" or status == 401:
        return "The API key was rejected. Check ANTHROPIC_API_KEY in the .env file."
    if name == "PermissionDeniedError" or status == 403:
        return f"The API refused the request (permission denied): {detail}"
    if name == "RateLimitError" or status == 429:
        return "The API rate limit was reached. Wait a minute and send the message again."
    if name == "NotFoundError" or status == 404:
        return f"The model name was not found. Check SOW_MODEL in the .env file. ({detail})"
    if name in ("APIConnectionError", "APITimeoutError"):
        return "Could not reach the Claude API. Check the internet connection and try again."
    if status == 400 and "credit" in detail.lower():
        return "The API account has no credit left. Add credit in the Claude Console, then try again."
    if status in (500, 529) or name in ("InternalServerError", "OverloadedError"):
        return "The Claude API is busy right now. Try again in a minute."
    return f"The request failed ({name}): {detail}"
