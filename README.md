# SOW Agent

An AI agent that writes Scope of Work (SOW) documents for lighting and electrical projects and
delivers them as Word files in the Osool SOW format (cover page, header, footer, lettered sections,
tables and schedule exactly as in your Granada Business Park SOW).

You chat with it, on your computer or from your phone through Telegram. It asks for what it needs,
writes the SOW, sends you the Word file, and revises it when you ask ("change the warranty to
3 years", "add a section for emergency lighting", "go back to version 2").

It runs on the Claude API with your own API key.

## What you need

- A computer with Windows, macOS or Linux and an internet connection
- Python 3.10 or newer: https://www.python.org/downloads/ (on Windows, tick "Add python.exe to PATH")
- A Claude API key (step 2 below). API use is billed to your Console account, separately from a
  Claude subscription; check the Billing page in the Console before first use.

## Setup (once)

1. Unzip the folder somewhere permanent, for example `Documents\sow-agent`.
2. Create your API key at https://platform.claude.com/settings/keys and copy it.
3. Windows: double-click `setup.bat`. It installs what is needed, creates the `.env` file and
   builds a sample Word file to prove the installation works.
   macOS / Linux: open a terminal in the folder and run

       python3 -m venv .venv && source .venv/bin/activate
       pip install -r requirements.txt
       cp .env.example .env
       python tools/selftest.py

4. Open the `.env` file with Notepad, paste the key after `ANTHROPIC_API_KEY=` and save.
5. Optional check of the key: `.venv\Scripts\python tools\selftest.py --api`

## Use it on the computer

Double-click `start_chat.bat` (macOS / Linux: `python chat.py`) and describe the job, for example:

    I need a SOW to replace the basement car park lighting in Tower B with LED fittings and
    motion sensors, 3 basements, about 600 fittings, 3-year warranty, 10 weeks.

The Word files are saved in `workspace\projects\<project>\` as `SOW_<title>_V1.docx`, `_V2.docx`, ...

Commands in the chat:

| Command | What it does |
|---|---|
| `/new [name]` | start a new SOW project |
| `/projects` | list saved projects |
| `/open <number>` | reopen a project from the list and continue |
| `/attach <file> [message]` | give the agent a file: PDF, Word, Excel, image or text (a BOQ, drawing, old SOW, contractor offer) |
| `/file` | show where the latest Word file is |
| `/quit` | leave |

## Use it as a web page (Streamlit)

On your computer: double-click `start_web.bat` (macOS / Linux: `streamlit run streamlit_app.py`).
The page opens in your browser at http://localhost:8501. Attach files with the paperclip in the
message box; each SOW appears with a Download button.

Online, so it works from your phone at any time (Streamlit Community Cloud, free):

1. Create a free GitHub account and a new **private** repository, for example `sow-agent`.
   Private matters: the folder contains your Osool template and reference SOW.
2. Upload everything in this folder to the repository (on the repository page: Add file, Upload
   files, drag the contents in). Do not upload `.env`, `.venv` or `workspace`.
3. Go to https://share.streamlit.io, sign in with GitHub, click **Create app**, then
   **Yup, I have an app**.
4. Choose your repository, branch `main`, and main file path `streamlit_app.py`.
5. Click **Advanced settings** and paste this into **Secrets**, with your own values:

       ANTHROPIC_API_KEY = "sk-ant-..."
       APP_PASSWORD = "choose-a-long-password"
       SOW_MODEL = "claude-opus-5-5"

6. Click **Save**, then deploy. After a few minutes you get an address ending in `streamlit.app`.
   Open it on your phone and add it to the home screen.

Notes for the online version:
- An app from a private repository is private: only you can open it after signing in, and you can
  invite viewers by email from the Share button. Community Cloud allows one private app at a time.
- `APP_PASSWORD` is a second lock in front of your API key. Keep it set.
- Download each Word file when it is made. Treat projects saved on the cloud server as temporary,
  because the server can be restarted.
- Community Cloud servers are in the United States. Check that this is acceptable for your
  company's documents before using it for real projects; otherwise run the page on a company
  machine with `start_web.bat`.

## Use it from your phone (Telegram)

1. In Telegram, open a chat with **@BotFather**, send `/newbot`, choose a name, and copy the token
   it gives you.
2. Paste the token in `.env` after `TELEGRAM_BOT_TOKEN=`.
3. Double-click `start_telegram_bot.bat` and send any message to your new bot. It replies with your
   Telegram user ID. Put that number in `.env` after `TELEGRAM_ALLOWED_USERS=` and start the bot
   again. From now on only you can use it.
4. Chat with the bot. It sends each SOW back as a Word file. You can also send it files; write your
   request in the caption.

The bot answers only while `start_telegram_bot.bat` is running. On your own PC that means while the
PC is on. To have it available day and night, run the same folder on a small always-on machine
(a cloud server or an office PC that stays on) with `python telegram_bot.py`.

## Make it yours

- `config\house_rules.md`: your standing instructions (client name, brand policy, default warranty,
  which clauses never change). Add one line per rule; the agent follows them in every SOW.
- `reference\reference_sow.json`: the worked example the agent learns structure and wording from
  (your Granada RGB lighting SOW). To teach it from a different SOW in the same format, run
  `python tools/import_docx.py "your SOW.docx" reference/reference_sow.json`.
- `templates\osool_sow_base.docx`: the Word file that supplies the cover, header, footer and styles.
- `.env`: model and web-search settings.

## Good to know

- Read every SOW before issuing it. The agent marks unknowns as `[To be confirmed]` and lists its
  assumptions in the chat, but you are the engineer responsible for the document.
- Payment terms, penalties, document control, risk and tenant-access clauses are carried over from
  your reference SOW unless you tell the agent to change them.
- Documents are in English. Arabic (right-to-left) documents are not supported.
- Only SOW documents are produced as Word files. Other requests (questions, reviews, emails) are
  answered in the chat.
- Each message costs API usage. A full new SOW is the most expensive step; small edits cost less
  because only the changed sections are rewritten. Set `SOW_MODEL=claude-sonnet-5-5` in `.env` for
  lower cost.
- Your conversations and SOWs stay in the `workspace` folder on your machine. The text you send is
  processed by the Claude API.

## If something goes wrong

| Message | What to do |
|---|---|
| No API key found / key rejected | Check `ANTHROPIC_API_KEY` in `.env` (no spaces or quotes) |
| No credit left | Add credit on the Billing page of the Claude Console |
| Model name was not found | Check `SOW_MODEL` in `.env` |
| Could not reach the Claude API | Check the internet connection or office firewall |
| Python was not found | Install Python and tick "Add python.exe to PATH", then run `setup.bat` again |
