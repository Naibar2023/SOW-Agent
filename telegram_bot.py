"""Run the SOW agent as a private Telegram bot, so you can use it from your phone.

    python telegram_bot.py

Needs TELEGRAM_BOT_TOKEN and TELEGRAM_ALLOWED_USERS in the .env file (see README).
The bot works for as long as this program is running.

Commands in Telegram:
    /new [name]     start a new SOW project
    /projects       list saved projects
    /open <number>  reopen a project from the list
    /file           send the latest Word file again
    /help           show this list
You can also send a PDF, Word, Excel or image file; the caption is your message.
"""

import json
import mimetypes
import os
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
from pathlib import Path

from sow_agent.agent import AgentError, SOWAgent
from sow_agent.config import load_env, workspace_dir

HELP = ("SOW Agent. Describe the project and I will prepare the Scope of Work as a Word file.\n\n"
        "/new [name] - start a new SOW project\n"
        "/projects - list saved projects\n"
        "/open <number> - reopen a project\n"
        "/file - send the latest Word file again\n\n"
        "You can send me PDF, Word, Excel or image files; write your request in the caption.")


class Telegram:
    def __init__(self, token):
        self.base = f"https://api.telegram.org/bot{token}/"
        self.file_base = f"https://api.telegram.org/file/bot{token}/"

    def call(self, method, params=None, timeout=30):
        data = urllib.parse.urlencode(params or {}).encode()
        req = urllib.request.Request(self.base + method, data=data)
        with urllib.request.urlopen(req, timeout=timeout) as r:
            out = json.loads(r.read().decode("utf-8"))
        if not out.get("ok"):
            raise RuntimeError(out.get("description", "Telegram error"))
        return out["result"]

    def send(self, chat_id, text):
        text = text or ""
        while text:
            part, text = text[:4000], text[4000:]
            if text:                                   # split at a line break when possible
                cut = part.rfind("\n")
                if cut > 2000:
                    part, text = part[:cut], part[cut:] + text
            try:
                self.call("sendMessage", {"chat_id": chat_id, "text": part})
            except Exception as e:
                print("sendMessage failed:", e)

    def action(self, chat_id, action="typing"):
        try:
            self.call("sendChatAction", {"chat_id": chat_id, "action": action}, timeout=10)
        except Exception:
            pass

    def send_document(self, chat_id, path: Path):
        boundary = uuid.uuid4().hex
        ctype = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        body = b"".join([
            f'--{boundary}\r\nContent-Disposition: form-data; name="chat_id"\r\n\r\n{chat_id}\r\n'.encode(),
            (f'--{boundary}\r\nContent-Disposition: form-data; name="document"; filename="{path.name}"\r\n'
             f"Content-Type: {ctype}\r\n\r\n").encode(),
            path.read_bytes(),
            f"\r\n--{boundary}--\r\n".encode(),
        ])
        req = urllib.request.Request(self.base + "sendDocument", data=body,
                                     headers={"Content-Type": f"multipart/form-data; boundary={boundary}"})
        with urllib.request.urlopen(req, timeout=120) as r:
            out = json.loads(r.read().decode("utf-8"))
        if not out.get("ok"):
            raise RuntimeError(out.get("description", "Telegram error"))

    def download(self, file_id, dest: Path):
        info = self.call("getFile", {"file_id": file_id})
        dest.parent.mkdir(parents=True, exist_ok=True)
        with urllib.request.urlopen(self.file_base + info["file_path"], timeout=120) as r:
            dest.write_bytes(r.read())
        return dest


class Typing:
    """Shows 'typing...' in the chat while the agent is working."""

    def __init__(self, tg, chat_id):
        self.tg, self.chat_id = tg, chat_id
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True)

    def _run(self):
        while not self._stop.is_set():
            self.tg.action(self.chat_id)
            self._stop.wait(4)

    def __enter__(self):
        self._thread.start()
        return self

    def __exit__(self, *exc):
        self._stop.set()


def safe_name(name):
    keep = "".join(c if c.isalnum() or c in "._- " else "_" for c in (name or "file"))
    return keep.strip(" .")[:120] or "file"


def main():
    load_env()
    token = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
    if not token:
        sys.exit("TELEGRAM_BOT_TOKEN is missing in the .env file. See README, 'Use it from your phone'.")
    allowed = {x.strip() for x in os.environ.get("TELEGRAM_ALLOWED_USERS", "").replace(";", ",").split(",")
               if x.strip()}
    tg = Telegram(token)
    try:
        me = tg.call("getMe")
    except Exception as e:
        sys.exit(f"Could not connect to Telegram with this token: {e}")
    print(f"Bot @{me.get('username')} is running. Leave this window open. Press Ctrl+C to stop.")
    if not allowed:
        print("TELEGRAM_ALLOWED_USERS is empty: the bot will only tell people their user ID until you set it.")

    agents, listings = {}, {}

    def agent_for(user_id, chat_id):
        if user_id not in agents:
            ws = workspace_dir() if len(allowed) == 1 else workspace_dir() / f"telegram_{user_id}"
            agents[user_id] = SOWAgent(ws)
        agents[user_id].on_status = lambda note: tg.send(chat_id, note)
        return agents[user_id]

    def handle(msg):
        chat_id = msg["chat"]["id"]
        user_id = str((msg.get("from") or {}).get("id", ""))
        if user_id not in allowed:
            tg.send(chat_id, f"This is a private bot. Your Telegram user ID is {user_id}.\n"
                             "If this is your bot, put that number in the .env file as "
                             "TELEGRAM_ALLOWED_USERS and restart the program.")
            return
        agent = agent_for(user_id, chat_id)
        text = (msg.get("text") or msg.get("caption") or "").strip()

        if text.startswith("/"):
            cmd, _, rest = text.partition(" ")
            cmd, rest = cmd.split("@")[0].lower(), rest.strip()
            if cmd in ("/start", "/help"):
                tg.send(chat_id, HELP)
                return
            if cmd == "/new":
                agent.new_project(rest or None)
                tg.send(chat_id, "New project started. Describe the work.")
                return
            if cmd == "/projects":
                listings[user_id] = agent.list_projects()
                rows = [f"{'*' if pid == agent.project_id else ' '} {i}. {name} (V{ver}, {upd})"
                        for i, (pid, name, ver, upd) in enumerate(listings[user_id], 1)]
                tg.send(chat_id, "\n".join(rows) or "No saved projects yet.")
                return
            if cmd == "/open":
                lst = listings.get(user_id) or agent.list_projects()
                if rest.isdigit() and 1 <= int(rest) <= len(lst):
                    agent.open_project(lst[int(rest) - 1][0])
                    tg.send(chat_id, f'Opened: {agent.state["name"]} (V{agent.state["version"]}).')
                else:
                    tg.send(chat_id, "Send /projects to see the list, then /open <number>.")
                return
            if cmd == "/file":
                f = agent.latest_file()
                if f:
                    tg.send_document(chat_id, f)
                else:
                    tg.send(chat_id, "No Word file has been made in this project yet.")
                return
            tg.send(chat_id, "Unknown command. Send /help.")
            return

        attachments = []
        inbox = agent.project_dir / "received"
        try:
            if "document" in msg:
                d = msg["document"]
                if d.get("file_size", 0) > 20 * 1024 * 1024:
                    tg.send(chat_id, "That file is larger than 20 MB, which Telegram bots cannot download.")
                    return
                attachments.append(tg.download(d["file_id"], inbox / safe_name(d.get("file_name"))))
            elif "photo" in msg:
                ph = msg["photo"][-1]                   # largest size
                attachments.append(tg.download(ph["file_id"], inbox / f"photo_{msg['message_id']}.jpg"))
        except Exception as e:
            tg.send(chat_id, f"Could not download the file: {e}")
            return
        if not text and not attachments:
            tg.send(chat_id, "I can read text messages and files (PDF, Word, Excel, images).")
            return

        with Typing(tg, chat_id):
            try:
                reply = agent.send(text, attachments)
            except AgentError as e:
                tg.send(chat_id, f"Problem: {e}")
                return
        tg.send(chat_id, reply.text)
        for f in reply.files:
            try:
                tg.send_document(chat_id, f)
            except Exception as e:
                tg.send(chat_id, f"The Word file was made but could not be sent ({e}). It is saved at: {f}")

    offset = None
    while True:
        try:
            params = {"timeout": 50, "allowed_updates": json.dumps(["message"])}
            if offset is not None:
                params["offset"] = offset
            updates = tg.call("getUpdates", params, timeout=70)
        except KeyboardInterrupt:
            print("\nStopped.")
            return
        except Exception as e:
            print("Connection problem, retrying in 5 s:", e)
            time.sleep(5)
            continue
        for up in updates:
            offset = up["update_id"] + 1
            msg = up.get("message")
            if not msg:
                continue
            try:
                handle(msg)
            except KeyboardInterrupt:
                print("\nStopped.")
                return
            except Exception as e:
                print("Error while handling a message:", repr(e))
                try:
                    tg.send(msg["chat"]["id"], f"Something went wrong: {e}")
                except Exception:
                    pass


if __name__ == "__main__":
    main()
