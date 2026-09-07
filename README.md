# Kimi Teams Plugin

A Kimi plugin to manage your **Microsoft Teams** account from chat, with read/write access through the **Microsoft Graph API** (delegated permissions). Read and send messages in channels and chats, manage teams and channels, create meetings and calendar events with Teams link, manage OneDrive files, check presence, and manage Microsoft To Do tasks.

---

## ✨ Features

- 💬 **Messaging** — Read and send messages in team channels (with search) and in 1:1 or group chats
- 👥 **Teams and channels** — List joined teams, list and create channels (public or private)
- 📅 **Meetings and calendar** — View upcoming events, create events with Teams link, create online meetings (joinUrl + phone access data)
- 📁 **Files** — List, upload, and download files in OneDrive
- 🟢 **Presence** — Check your status or any user's status in the tenant
- ✅ **Tasks** — Manage Microsoft To Do lists and tasks
- 🔐 **OAuth2 Authentication** — Microsoft device flow (MSAL), token auto-refreshed and stored locally in `~/.kimi-teams/`

---

## 📋 Prerequisites

Before installing you need:

1. **Kimi Desktop** installed on your machine
2. A **Microsoft 365** account with Teams license
3. Python 3 with `msal` and `requests` packages (already included in Kimi Work runtime)
4. Delegated Microsoft Graph permissions granted. The plugin reuses an existing public Azure AD app; **some permissions (e.g. `ChannelMessage.Read.All`, `Chat.Read.All`) may require tenant admin consent**. If a command fails due to permissions, they must be added in **Azure Portal → Microsoft Entra ID → App registrations → API permissions** (type *delegated*) and granted.

---

## 🚀 Installation

### Step 1: clone the repo

```
git clone https://github.com/FlacH7/kimi-teams-plugin.git
```

### Step 2: register the plugin in your personal marketplace

#### Option A (recommended):

In a Kimi conversation, ask:

> "Register this local plugin in my personal marketplace: `<absolute path to repo>`"

The agent validates and registers it. Then open Kimi's Plugins page → **「个人」 (Personal)** tab → click **＋** on "Microsoft Teams Manager". No restart.

#### Option B (manual, if A fails):

Open a terminal:

**Windows (PowerShell):**

```
$share = "$env:APPDATA\kimi-desktop\daimon-share"
& "$env:APPDATA\kimi-desktop\daimon-bundle\bin\kimi-daimon.cmd" kimi-plugin register-personal "C:\full\path\to\kimi-teams-plugin" --share-dir "$share"
```

**macOS:**

```
kimi-daimon kimi-plugin register-personal ~/path/to/kimi-teams-plugin \
  --share-dir ~/Library/Application\ Support/kimi-desktop/daimon-share
```

**Linux:**

```
kimi-daimon kimi-plugin register-personal ~/path/to/kimi-teams-plugin \
  --share-dir ~/.kimi/daimon-share
```

### Step 3: install from Kimi Desktop

1. Open **Kimi Desktop**
2. Go to the **Plugins** panel (插件)
3. Switch to the **「个人」** (Personal) tab
4. Search for **"Microsoft Teams Manager"**
5. Click **＋** or **「安装」** (Install)
6. Done! The plugin activates instantly, **no restart**

---

## 🔐 Configuration and authentication

The first time, Kimi will run the device flow:

1. Kimi generates a code and shows you the Microsoft message
2. Open **https://login.microsoft.com/device** in your browser
3. Enter the code and **sign in with your Microsoft 365 account**
4. Approve the requested permissions
5. Kimi completes the flow and saves the token to `~/.kimi-teams/msal_cache.json` (auto-refreshes)

> **Security:** the token never leaves your machine. There are no hardcoded secrets in the plugin; the `client_id` belongs to a public Azure AD app and can be overwritten with environment variables `TEAMS_CLIENT_ID` and `TEAMS_TENANT_ID`.

---

## 💬 Usage

Once installed and authenticated, tell Kimi in natural language what you want to do:

| Request | Action |
|---|---|
| _"Read messages from the General channel of the Marketing team"_ | Read channel |
| _"Send a message to Ana on Teams saying I'll be 10 minutes late"_ | Send chat |
| _"What meetings do I have this week?"_ | View calendar |
| _"Create a meeting with Juan tomorrow at 10"_ | Create meeting |
| _"Upload this report to OneDrive"_ | Upload file |
| _"Is María available?"_ | Presence |
| _"Add to my task list: prepare the presentation, due Friday"_ | To Do |

---

## 🏗️ Architecture

```
kimi-teams-plugin/
├── kimi.plugin.json          # Plugin manifest
├── README.md
├── LICENSE
├── scripts/
│   └── teams_cli.py          # CLI over Microsoft Graph API v1.0 (MSAL device flow)
└── skills/
    └── teams/
        └── SKILL.md          # Agent instructions and command reference
```

### `teams_cli.py` — Command Reference

```
# Authentication
python3 scripts/teams_cli.py auth                    # device flow (blocks until approved)
python3 scripts/teams_cli.py auth --no-wait          # shows the code and exits
python3 scripts/teams_cli.py auth-complete           # completes the flow after approval

# Profile
python3 scripts/teams_cli.py me

# Teams and channels
python3 scripts/teams_cli.py teams
python3 scripts/teams_cli.py channels --team <team>
python3 scripts/teams_cli.py channel-create --team <t> --name <n> [--description d] [--private]

# Channel messages
python3 scripts/teams_cli.py messages --team <t> --channel <c> [--limit N] [--query text]
python3 scripts/teams_cli.py send --team <t> --channel <c> --text "..."

# Chats
python3 scripts/teams_cli.py chats [--limit N]
python3 scripts/teams_cli.py chat-messages --chat <id or member> [--limit N]
python3 scripts/teams_cli.py chat-send --chat <id or member> --text "..."
python3 scripts/teams_cli.py chat-create --users a@x.com,b@x.com [--topic t]

# Calendar and meetings
python3 scripts/teams_cli.py events [--days N]
python3 scripts/teams_cli.py event-create --title T --start "YYYY-MM-DD HH:MM" --end "..." [--attendees a,b] [--body html]
python3 scripts/teams_cli.py meeting-create --title T --start "..." --end "..." [--attendees a,b]

# Files (OneDrive)
python3 scripts/teams_cli.py files [--path path]
python3 scripts/teams_cli.py file-upload --local <file> [--remote path]
python3 scripts/teams_cli.py file-download --remote <path> --local <file>

# Presence
python3 scripts/teams_cli.py presence [--user <upn>]

# Tasks (To Do)
python3 scripts/teams_cli.py todo-lists
python3 scripts/teams_cli.py todo-tasks --list <id>
python3 scripts/teams_cli.py todo-add --list <id> --title T [--due "YYYY-MM-DD HH:MM"]
python3 scripts/teams_cli.py todo-complete --list <id> --task <id>
```

Notes:
- Team/channel/chat names accept the display name or id; the CLI resolves them.
- Dates use format `YYYY-MM-DD HH:MM` (UTC).
- Graph permissions are requested by groups (teams / chats / calendar / files / presence / todo); consent is incremental by command.

---

## 🛠️ Troubleshooting

### "kimi-daimon is not recognized as a command"

Find the executable in your Kimi Desktop installation:

- **Windows:** `C:\Users\<USER>\AppData\Roaming\kimi-desktop\daimon-bundle\bin\kimi-daimon.cmd`
- **macOS:** `~/Library/Application Support/kimi-desktop/daimon-bundle/bin/kimi-daimon`
- **Linux:** `~/.kimi/daimon-bundle/bin/kimi-daimon`

### "The plugin does not appear in the 个人 tab"

Check that:

1. The `register-personal` command finished without errors
2. The folder contains `kimi.plugin.json` at the root
3. Refresh the Plugins page (close and reopen)

### "GRAPH ERROR 401 / 403" or consent errors

Missing delegated permission in Azure AD. The error message indicates the scope. Path: **Azure Portal → Microsoft Entra ID → App registrations → (your app) → API permissions → Add a permission → Microsoft Graph → Delegated**. For an organizational tenant, the **"Grant admin consent"** button requires administrator role.

### Token expired or authentication fails

Run `python3 scripts/teams_cli.py auth` (device flow) again to regenerate it.

---

## 🚧 Current Development Status

**Version: 0.1.0** — functional, in initial validation phase.

### ✅ Implemented and tested

- Scaffolding, manifest, and skill complete; local validation (`validate_plugin.py`) without errors
- CLI `teams_cli.py` with 22 subcommands covering: auth (device flow, `--no-wait`/`auth-complete`), profile, teams, channels (list/create), channel messages (read/search/send), chats (list/read/send/create), calendar (list/create event with Teams link), online meetings, OneDrive (list/upload/download), presence and To Do (lists/tasks/create/complete)
- Registration and installation in Kimi's personal marketplace (full flow tested on Windows)
- Automatic name → id resolution for teams, channels, and chats
- Automatic token renewal (silent + force refresh with retry on 401)

### ⚠️ Pending live validation

- **Complete authentication with real account**: the device flow was generated but user approval and token acquisition have not yet been verified end-to-end
- Delegated permission consent: pending confirmation of which scopes are blocked by tenant policies (especially `ChannelMessage.Read.All`, `Chat.Read.All`, `OnlineMeetings.ReadWrite`)
- Write commands (send messages, create channels/meetings/tasks, upload files) not yet tested against real Graph
- Message search (`--query`): depends on `ConsistencyLevel: eventual` header and may require additional permissions

### 🔜 Next steps (roadmap)

1. **Validate real auth** — complete the device flow, run `teams` and `chats`, note which permissions require admin consent and document them in this README
2. **E2E tests by area** — test sheet: read/send in channel and chat, create meeting and check joinUrl, upload/download file, create/complete task
3. **Pagination** — `@odata.nextLink` is not yet followed; lists are limited to the first page (top 100/200)
4. **Message attachments** — send/receive files in channel and chat messages (upload session for >4 MB)
5. **Reactions and replies** — add/remove reactions, reply in thread (`replyToId`)
6. **Notifications** — evaluate Microsoft Graph change notifications (requires webhook or Azure, probably out of scope for local use)
7. **Calls / Teams Phone** — via Cloud Communications API (elevated permissions; evaluate feasibility)
8. **Robust error handling** — retries with backoff, translated error messages by HTTP code, `--json` mode for structured output
9. **Unit tests** — Graph mocks for read-only commands
10. **Basic CI** — GitHub Actions: lint + validate_plugin.py on every push

---

## 📄 Changelog

### v0.1.0

- Initial plugin: messaging (channels and chats), teams/channels, calendar and meetings, OneDrive, presence and To Do
- MSAL device flow authentication with local token cache
- 22 subcommands in `teams_cli.py`
- Registration in Kimi's personal marketplace (installation without restart)

---

## 📜 License

MIT — see [LICENSE](LICENSE)
