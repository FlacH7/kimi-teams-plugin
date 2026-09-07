---
name: teams
description: Manage Microsoft Teams from Kimi. Use it when the user wants to read or send messages in channels or chats, list teams/channels, create channels, view or create meetings and calendar events with Teams link, manage OneDrive files, check presence, or manage Microsoft To Do tasks on their Microsoft 365 account. Requires prior authentication with `auth`.
---

# Microsoft Teams Manager

Operates the user's Microsoft Teams account (Microsoft 365) through Microsoft Graph API with delegated permissions. Everything goes through the local CLI:

```bash
python3 scripts/teams_cli.py <command> [options]
```

The script lives in the plugin root directory (`plugins/teams/scripts/teams_cli.py`); resolves the path relative to this SKILL.md if needed. Dependencies: `msal` and `requests` (already available in the runtime).

## Authentication

- First time: `python3 scripts/teams_cli.py auth --no-wait` → generates a **device code** and saves the flow. Tell the user to open `https://login.microsoft.com/device`, enter the code and sign in with their Microsoft 365 account; when confirmed, run `python3 scripts/teams_cli.py auth-complete` to finish. Alternative in one step: `auth` (blocks until approved). The token is saved to `~/.kimi-teams/msal_cache.json` and auto-refreshes.
- The app requests Graph permissions by groups (teams, chats, calendar, files, presence, To Do). If any command fails with a **consent/admin** error, the corresponding delegated permissions must be granted in Azure AD (some, like `ChannelMessage.Read.All` or `Chat.Read.All`, usually require administrator consent). Communicate this clearly to the user.
- Reuses the same tenant/client as the neuro-email plugin (`pedro.valdes@neuroinformatics-collaboratory.org`); can be overridden with `TEAMS_CLIENT_ID` / `TEAMS_TENANT_ID`.

## Usage rules

- **Destructive or visible write operations (send messages, create channels/meetings/tasks, upload files): always confirm with the user before executing**, showing recipient and content.
- Team/channel/chat names accept display name or id; the CLI resolves them automatically.
- Dates in format `YYYY-MM-DD HH:MM` (interpreted as UTC).

## Commands

| Area | Command | Description |
|---|---|---|
| Auth | `auth` | Device flow authentication (do first) |
| Profile | `me` | Current user data |
| Teams | `teams` | List joined teams |
| Channels | `channels --team <team>` | List channels |
| Channels | `channel-create --team <t> --name <n> [--description d] [--private]` | Create channel |
| Messages | `messages --team <t> --channel <c> [--limit N] [--query text]` | Read channel messages (search with `--query`) |
| Messages | `send --team <t> --channel <c> --text "..."` | Send message to channel |
| Chats | `chats [--limit N]` | List chats |
| Chats | `chat-messages --chat <id or member> [--limit N]` | Read messages from a chat |
| Chats | `chat-send --chat <id or member> --text "..."` | Send message to a chat |
| Chats | `chat-create --users a@x.com,b@x.com [--topic t]` | Create 1:1 or group chat |
| Calendar | `events [--days N]` | Upcoming events (includes Teams meetings) |
| Calendar | `event-create --title T --start "..." --end "..." [--attendees a,b] [--body html]` | Create event with Teams link |
| Meetings | `meeting-create --title T --start "..." --end "..." [--attendees a,b]` | Create Teams meeting (returns joinUrl) |
| Files | `files [--path path]` | List OneDrive |
| Files | `file-upload --local <file> [--remote path]` | Upload file |
| Files | `file-download --remote <path> --local <file>` | Download file |
| Presence | `presence [--user <upn>]` | My presence or another user's |
| Tasks | `todo-lists` | List To Do lists |
| Tasks | `todo-tasks --list <id>` | List tasks from a list |
| Tasks | `todo-add --list <id> --title T [--due "YYYY-MM-DD HH:MM"]` | Create task |
| Tasks | `todo-complete --list <id> --task <id>` | Complete task |

## Example user requests

- "Read messages from the General channel of the Marketing team"
- "Send a message to Ana on Teams saying I'm running 10 minutes late"
- "What meetings do I have this week?" / "Create a meeting with Juan tomorrow at 10"
- "Upload this report to OneDrive and send the link to María"
- "Create a To Do task for Friday: prepare the presentation"
