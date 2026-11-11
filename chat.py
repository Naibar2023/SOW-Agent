"""Chat with the SOW agent in a terminal window.

    python chat.py

Commands inside the chat:
    /new [name]          start a new SOW project
    /projects            list saved projects
    /open <number>       reopen a project from the list
    /attach <file> [message]   give the agent a file (PDF, Word, Excel, image, text)
    /file                show where the latest Word file is
    /help                show this list
    /quit                leave
"""

import os
import shlex
import sys
from pathlib import Path

from sow_agent.agent import AgentError, SOWAgent
from sow_agent.config import load_env, workspace_dir

HELP = __doc__.split("Commands inside the chat:")[1].rstrip()


def main():
    load_env()
    for stream in (sys.stdout, sys.stdin):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass
    state = {"open": False, "any": False}   # is a reply line open / has any reply text been shown

    def on_text(chunk):
        if not state["open"]:
            sys.stdout.write("\nAgent: ")
            state["open"] = True
        state["any"] = True
        sys.stdout.write(chunk)
        sys.stdout.flush()

    def on_status(note):
        sys.stdout.write(f"\n  [{note}]\n")
        sys.stdout.flush()
        state["open"] = False

    agent = SOWAgent(workspace_dir(), on_text=on_text, on_status=on_status)
    print("SOW Agent. Describe the project you need a Scope of Work for.")
    print("Type /help for commands, /quit to leave.")
    st = agent.state
    if st["history"]:
        print(f'\nContinuing project: {st["name"]} (V{st["version"]}). Type /new to start another one.')
    listing = []

    while True:
        try:
            line = input("\nYou: ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            break
        if not line:
            continue
        attachments = []
        if line.startswith("/"):
            cmd, _, rest = line.partition(" ")
            cmd, rest = cmd.lower(), rest.strip()
            if cmd in ("/quit", "/exit"):
                break
            if cmd == "/help":
                print(HELP)
                continue
            if cmd == "/new":
                agent.new_project(rest or None)
                print("New project started. Describe the work.")
                continue
            if cmd == "/projects":
                listing = agent.list_projects()
                if not listing:
                    print("No saved projects yet.")
                for i, (pid, name, ver, upd) in enumerate(listing, 1):
                    mark = "*" if pid == agent.project_id else " "
                    print(f" {mark}{i}. {name}  (V{ver}, {upd})")
                continue
            if cmd == "/open":
                listing = listing or agent.list_projects()
                if rest.isdigit() and 1 <= int(rest) <= len(listing):
                    agent.open_project(listing[int(rest) - 1][0])
                    print(f'Opened: {agent.state["name"]} (V{agent.state["version"]}).')
                else:
                    print("Use /projects to see the list, then /open <number>.")
                continue
            if cmd == "/file":
                f = agent.latest_file()
                print(f if f else "No Word file has been made in this project yet.")
                continue
            if cmd == "/attach":
                try:
                    parts = shlex.split(rest, posix=(os.name != "nt"))
                except ValueError:
                    parts = rest.split()
                if not parts:
                    print('Usage: /attach "C:\\path\\to\\file.pdf" optional message')
                    continue
                path = Path(parts[0].strip('"')).expanduser()
                if not path.is_file():
                    print(f"File not found: {path}")
                    continue
                attachments = [path]
                line = " ".join(parts[1:])
            else:
                print("Unknown command. Type /help.")
                continue
        state["open"] = state["any"] = False
        try:
            reply = agent.send(line, attachments)
        except AgentError as e:
            print(f"\nProblem: {e}")
            continue
        except KeyboardInterrupt:
            print("\n(stopped)")
            continue
        if not state["any"] and reply.text:
            print(f"\nAgent: {reply.text}")
        else:
            print()
        for f in reply.files:
            print(f"\n  Word file saved: {f}")


if __name__ == "__main__":
    main()
