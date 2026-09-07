#!/usr/bin/env python3
"""Microsoft Teams CLI for Kimi — Microsoft Graph API (delegated, MSAL device flow).

Auth data lives in ~/.kimi-teams/ (token cache). Reuses the public client
registration already used by the neuro-email plugin for the same tenant.
"""

import argparse
import json
import os
import sys
import time
from datetime import datetime, timedelta, timezone

import msal
import requests

CLIENT_ID = os.environ.get("TEAMS_CLIENT_ID", "e2469caf-4b8d-4c34-8af4-02a09b6ed6d3")
TENANT_ID = os.environ.get("TEAMS_TENANT_ID", "b0d60629-de22-4251-96f3-1d43de1f3201")
GRAPH = "https://graph.microsoft.com/v1.0"
AUTH_DIR = os.path.expanduser(os.environ.get("TEAMS_AUTH_DIR", "~/.kimi-teams"))
CACHE_FILE = os.path.join(AUTH_DIR, "msal_cache.json")

FLOW_FILE = os.path.join(AUTH_DIR, "device_flow.json")

# Scopes requested incrementally per command group.
SCOPE_GROUPS = {
    "base": ["User.Read"],
    "teams": ["Team.ReadBasic.All", "Channel.ReadBasic.All",
              "ChannelMessage.Read.All", "ChannelMessage.Send",
              "Group.Read.All"],
    "chats": ["Chat.ReadWrite", "Chat.Create", "ChatMessage.Send", "Chat.Read.All"],
    "calendar": ["Calendars.ReadWrite", "OnlineMeetings.ReadWrite"],
    "files": ["Files.ReadWrite"],
    "presence": ["Presence.Read"],
    "todo": ["Tasks.ReadWrite"],
}


def ensure_dir():
    os.makedirs(AUTH_DIR, exist_ok=True)


def get_app():
    ensure_dir()
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE) as f:
            cache.deserialize(f.read())
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache,
    )
    return app, cache


def save_cache(cache):
    with open(CACHE_FILE, "w") as f:
        f.write(cache.serialize())


def cmd_auth(args):
    """Authenticate (device flow). Run this first.

    Without flags it blocks until approval. Use --no-wait to print the code
    and exit, then `auth-complete` afterwards to finish the flow.
    """
    scopes = sorted({s for group in SCOPE_GROUPS.values() for s in group})
    app, cache = get_app()
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
        if result and "access_token" in result:
            print("Already authenticated as:", result.get("id_token_claims", {}).get("preferred_username", "?"))
            save_cache(cache)
            return
    flow = app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        raise RuntimeError(str(flow))
    if getattr(args, "no_wait", False):
        with open(FLOW_FILE, "w") as f:
            json.dump(flow, f)
        print(flow["message"])
        print("(flow saved; run `auth-complete` once you have approved)")
        return
    print(flow["message"])
    print("Waiting for approval...", flush=True)
    result = app.acquire_token_by_device_flow(flow)
    save_cache(cache)
    if "access_token" not in result:
        print("AUTH FAILED:", result.get("error_description", result.get("error")), file=sys.stderr)
        sys.exit(1)
    print("Authenticated as:", result.get("id_token_claims", {}).get("preferred_username", "?"))
    print("NOTE: if a scope failed with admin-consent errors, ask your M365 admin to grant")


def cmd_auth_complete(_args):
    """Complete a device flow started with `auth --no-wait`."""
    if not os.path.exists(FLOW_FILE):
        raise SystemExit("No saved device flow. Run `auth --no-wait` first.")
    with open(FLOW_FILE) as f:
        flow = json.load(f)
    app, cache = get_app()
    print("Waiting for approval...", flush=True)
    result = app.acquire_token_by_device_flow(flow)
    save_cache(cache)
    os.remove(FLOW_FILE)
    if "access_token" not in result:
        print("AUTH FAILED:", result.get("error_description", result.get("error")), file=sys.stderr)
        sys.exit(1)
    print("Authenticated as:", result.get("id_token_claims", {}).get("preferred_username", "?"))
    print("NOTE: if a scope failed with admin-consent errors, ask your M365 admin to grant")


def get_token(scope_group):
    scopes = SCOPE_GROUPS[scope_group]
    app, cache = get_app()
    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
    if not result or "access_token" not in result:
        save_cache(cache)
        print("Not authenticated or consent missing for scopes:", ", ".join(scopes), file=sys.stderr)
        print("Run: python3 scripts/teams_cli.py auth", file=sys.stderr)
        sys.exit(2)
    save_cache(cache)
    return result["access_token"]


def graph(scope_group, method, path, params=None, json_body=None, data=None, headers=None):
    token = get_token(scope_group)
    url = path if path.startswith("http") else GRAPH + path
    h = {"Authorization": f"Bearer {token}"}
    if headers:
        h.update(headers)
    r = requests.request(method, url, params=params, json=json_body, data=data, headers=h)
    if r.status_code == 401:
        # force refresh once
        app, cache = get_app()
        result = app.acquire_token_silent(SCOPE_GROUPS[scope_group],
                                          account=app.get_accounts()[0], force_refresh=True)
        save_cache(cache)
        if result and "access_token" in result:
            h["Authorization"] = f"Bearer {result['access_token']}"
            r = requests.request(method, url, params=params, json=json_body, data=data, headers=h)
    if r.status_code >= 400:
        msg = r.text
        try:
            err = r.json().get("error", {})
            msg = err.get("message", r.text)
        except Exception:
            pass
        print(f"GRAPH ERROR {r.status_code}: {msg}", file=sys.stderr)
        if "consent" in msg.lower() or r.status_code in (401, 403):
            print("Tip: puede faltar consentimiento del administrador para este alcance en Azure AD.", file=sys.stderr)
        sys.exit(1)
    if r.status_code == 204:
        return {}
    return r.json() if r.content else {}


def fmt_user(u):
    if not u:
        return "?"
    return u.get("displayName") or u.get("userPrincipalName") or u.get("id", "?")


def fmt_msg(m):
    from_ = fmt_user(m.get("from", {}).get("user"))
    ts = m.get("createdDateTime", "")
    body = (m.get("body", {}) or {}).get("content", "")
    import re
    body = re.sub(r"<[^>]+>", " ", body)
    body = " ".join(body.split())
    return f"[{ts}] {from_}: {body}"


def print_table(rows, cols):
    widths = {c: max(len(c), max((len(str(r.get(c, ""))) for r in rows), default=0)) for c in cols}
    print("  ".join(c.ljust(widths[c]) for c in cols))
    print("  ".join("-" * widths[c] for c in cols))
    for r in rows:
        print("  ".join(str(r.get(c, "")).ljust(widths[c]) for c in cols))


# ---------------- Teams / Channels / Messages ----------------

def list_teams(_args):
    data = graph("teams", "GET", "/me/joinedTeams")
    rows = [{"id": t["id"], "name": t["displayName"], "desc": t.get("description", "")}
            for t in data.get("value", [])]
    print_table(rows, ["id", "name", "desc"]) if rows else print("(no teams)")


def resolve_team(name_or_id):
    if name_or_id.startswith("https://") or "-" in name_or_id and len(name_or_id) > 30:
        return name_or_id
    data = graph("teams", "GET", "/me/joinedTeams")
    for t in data.get("value", []):
        if t["id"] == name_or_id or t["displayName"].lower() == name_or_id.lower():
            return t["id"]
    print(f"Team not found: {name_or_id}", file=sys.stderr)
    sys.exit(1)


def resolve_channel(team_id, name_or_id):
    data = graph("teams", "GET", f"/teams/{team_id}/channels")
    for c in data.get("value", []):
        if c["id"] == name_or_id or c["displayName"].lower() == name_or_id.lower():
            return c["id"]
    print(f"Channel not found: {name_or_id}", file=sys.stderr)
    sys.exit(1)


def list_channels(args):
    tid = resolve_team(args.team)
    data = graph("teams", "GET", f"/teams/{tid}/channels")
    rows = [{"id": c["id"], "name": c["displayName"], "desc": c.get("description", ""),
             "folder": (c.get("filesFolder", {}) or {}).get("webUrl", "")}
            for c in data.get("value", [])]
    print_table(rows, ["id", "name", "desc"]) if rows else print("(no channels)")


def create_channel(args):
    tid = resolve_team(args.team)
    body = {"displayName": args.name, "description": args.description or ""}
    if args.private:
        body["membershipType"] = "private"
    c = graph("teams", "POST", f"/teams/{tid}/channels", json_body=body)
    print("Channel created:", c.get("id"), c.get("displayName"))


def list_messages(args):
    tid = resolve_team(args.team)
    cid = resolve_channel(tid, args.channel)
    params = {"$top": str(args.limit), "$orderby": "createdDateTime desc"}
    if args.query:
        params["$search"] = f'"{args.query}"'
    data = graph("teams", "GET", f"/teams/{tid}/channels/{cid}/messages", params=params,
                 headers={"ConsistencyLevel": "eventual"} if args.query else None)
    msgs = data.get("value", [])
    for m in reversed(msgs):
        print(fmt_msg(m))


def send_message(args):
    tid = resolve_team(args.team)
    cid = resolve_channel(tid, args.channel)
    body = {"body": {"contentType": "html", "content": args.text.replace("\n", "<br>")}}
    m = graph("teams", "POST", f"/teams/{tid}/channels/{cid}/messages", json_body=body)
    print("Message sent:", m.get("id"))


# ---------------- Chats ----------------

def list_chats(args):
    data = graph("chats", "GET", "/me/chats", params={"$top": str(args.limit), "$expand": "members"})
    rows = []
    for c in data.get("value", []):
        members = ", ".join(sorted({fmt_user(m.get("user")) for m in c.get("members", [])}))
        rows.append({"id": c["id"], "type": c.get("chatType", ""), "members": members})
    print_table(rows, ["id", "type", "members"]) if rows else print("(no chats)")


def resolve_chat(name_or_id):
    if len(name_or_id) > 30 and "-" in name_or_id:
        return name_or_id
    data = graph("chats", "GET", "/me/chats", params={"$top": "100", "$expand": "members"})
    needle = name_or_id.lower()
    for c in data.get("value", []):
        members = ", ".join(sorted({fmt_user(m.get("user")) for m in c.get("members", [])}))
        if c["id"] == name_or_id or needle in members.lower():
            return c["id"]
    print(f"Chat not found: {name_or_id}", file=sys.stderr)
    sys.exit(1)


def chat_messages(args):
    cid = resolve_chat(args.chat)
    data = graph("chats", "GET", f"/me/chats/{cid}/messages",
                 params={"$top": str(args.limit), "$orderby": "createdDateTime desc"})
    for m in reversed(data.get("value", [])):
        print(fmt_msg(m))


def chat_send(args):
    cid = resolve_chat(args.chat)
    body = {"body": {"contentType": "html", "content": args.text.replace("\n", "<br>")}}
    m = graph("chats", "POST", f"/me/chats/{cid}/messages", json_body=body)
    print("Message sent:", m.get("id"))


def chat_create(args):
    members = [{"@odata.type": "#microsoft.graph.aadUserConversationMember",
                "roles": ["owner"],
                "user@odata.bind": f"{GRAPH}/users('{u}')"} for u in args.users]
    body = {"chatType": "oneOnOne" if len(members) == 1 else "group", "members": members}
    if args.topic:
        body["chatType"] = "group"
        body["topic"] = args.topic
    c = graph("chats", "POST", "/chats", json_body=body)
    print("Chat created:", c.get("id"))


# ---------------- Calendar / Meetings ----------------

def events_list(args):
    start = datetime.now(timezone.utc)
    end = start + timedelta(days=args.days)
    params = {
        "startDateTime": start.isoformat(timespec="seconds").replace("+00:00", "Z"),
        "endDateTime": end.isoformat(timespec="seconds").replace("+00:00", "Z"),
        "$orderby": "start/dateTime",
        "$top": "100",
    }
    data = graph("calendar", "GET", "/me/calendarView", params=params)
    rows = []
    for e in data.get("value", []):
        rows.append({
            "id": e["id"],
            "start": e.get("start", {}).get("dateTime", ""),
            "subject": e.get("subject", ""),
            "location": e.get("location", {}).get("displayName", "") or "Teams",
            "online": "yes" if e.get("isOnlineMeeting") or e.get("onlineMeeting") else "",
        })
    print_table(rows, ["start", "subject", "location", "online", "id"]) if rows else print("(no events)")


def _parse_dt(s):
    s = s.strip()
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(s, fmt)
            return dt.replace(tzinfo=timezone.utc)
        except ValueError:
            continue
    raise SystemExit(f"Bad datetime (use YYYY-MM-DD HH:MM): {s}")


def event_create(args):
    start = _parse_dt(args.start)
    end = _parse_dt(args.end)
    tz = "UTC"
    body = {
        "subject": args.title,
        "start": {"dateTime": start.isoformat(timespec="seconds").replace("+00:00", "Z"), "timeZone": tz},
        "end": {"dateTime": end.isoformat(timespec="seconds").replace("+00:00", "Z"), "timeZone": tz},
        "isOnlineMeeting": True,
        "onlineMeetingProvider": "teamsForBusiness",
        "attendees": [{"emailAddress": {"address": a}, "type": "required"}
                      for a in (args.attendees.split(",") if args.attendees else [])],
        "body": {"contentType": "html", "content": args.body or ""},
    }
    e = graph("calendar", "POST", "/me/events", json_body=body)
    print("Event created:", e.get("id"))
    print("Subject:", e.get("subject"))
    print("Join URL:", (e.get("onlineMeeting") or {}).get("joinUrl", ""))


def meeting_create(args):
    start = _parse_dt(args.start)
    end = _parse_dt(args.end)
    body = {
        "subject": args.title,
        "startDateTime": start.isoformat(timespec="seconds").replace("+00:00", "Z"),
        "endDateTime": end.isoformat(timespec="seconds").replace("+00:00", "Z"),
    }
    if args.attendees:
        body["participants"] = {"attendees": [{"identity": {"user": {"id": a}} if "@" not in a
                                               else {"user": {"userPrincipalName": a}}}
                                               for a in args.attendees.split(",")]}
    m = graph("calendar", "POST", "/me/onlineMeetings", json_body=body)
    print("Meeting created:", m.get("id"))
    print("Join URL:", m.get("joinUrl"))
    print("Passcode:", m.get("passcode", ""), "| Toll:", (m.get("audioConferencing") or {}).get("tollNumber", ""))


# ---------------- Files (OneDrive / SharePoint) ----------------

def files_list(args):
    path = f"/me/drive/root:/{args.path}:/children" if args.path else "/me/drive/root/children"
    data = graph("files", "GET", path, params={"$top": "200"})
    rows = [{"name": i["name"], "type": "dir" if "folder" in i else "file",
             "size": i.get("size", 0),
             "modified": i.get("lastModifiedDateTime", ""),
             "id": i["id"]} for i in data.get("value", [])]
    print_table(rows, ["type", "name", "size", "modified"]) if rows else print("(empty)")


def file_upload(args):
    local = args.local
    remote = args.remote or os.path.basename(local)
    if not os.path.exists(local):
        raise SystemExit(f"Local file not found: {local}")
    with open(local, "rb") as f:
        content = f.read()
    item = graph("files", "PUT", f"/me/drive/root:/{remote}:/content",
                 data=content, headers={"Content-Type": "application/octet-stream"})
    print("Uploaded:", (item.get("webUrl") or item.get("id")))


def file_download(args):
    item = graph("files", "GET", f"/me/drive/root:/{args.remote}")
    dl = item.get("@microsoft.graph.downloadUrl")
    if not dl:
        raise SystemExit("No download URL (file missing?)")
    r = requests.get(dl)
    with open(args.local, "wb") as f:
        f.write(r.content)
    print(f"Saved {len(r.content)} bytes -> {args.local}")


# ---------------- Presence ----------------

def presence(args):
    if args.user:
        p = graph("presence", "GET", f"/users/{args.user}/presence")
    else:
        p = graph("presence", "GET", "/me/presence")
    print(json.dumps(p, indent=2, ensure_ascii=False))


# ---------------- To Do tasks ----------------

def todo_lists(_args):
    data = graph("todo", "GET", "/me/todo/lists")
    rows = [{"id": l["id"], "name": l["displayName"]} for l in data.get("value", [])]
    print_table(rows, ["id", "name"]) if rows else print("(no lists)")


def todo_tasks(args):
    data = graph("todo", "GET", f"/me/todo/lists/{args.list}/tasks",
                 params={"$top": "100", "$orderby": "createdDateTime desc"})
    rows = [{"status": t.get("status", ""), "title": t.get("title", ""),
             "due": (t.get("dueDateTime") or {}).get("dateTime", ""), "id": t["id"]}
            for t in data.get("value", [])]
    print_table(rows, ["status", "title", "due", "id"]) if rows else print("(no tasks)")


def todo_add(args):
    body = {"title": args.title}
    if args.due:
        body["dueDateTime"] = {"dateTime": _parse_dt(args.due).isoformat(timespec="seconds").replace("+00:00", "Z"),
                               "timeZone": "UTC"}
    t = graph("todo", "POST", f"/me/todo/lists/{args.list}/tasks", json_body=body)
    print("Task created:", t.get("id"), t.get("title"))


def todo_complete(args):
    graph("todo", "PATCH", f"/me/todo/lists/{args.list}/tasks/{args.task}",
          json_body={"status": "completed"})
    print("Task completed.")


# ---------------- Main ----------------

def main():
    p = argparse.ArgumentParser(description="Microsoft Teams CLI (Graph API, delegated)")
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("auth", help="Authenticate via device flow")
    s.add_argument("--no-wait", action="store_true",
                   help="print code and exit; finish later with auth-complete")
    s.set_defaults(fn=cmd_auth)
    sub.add_parser("auth-complete", help="Complete device flow started with auth --no-wait").set_defaults(fn=cmd_auth_complete)
    sub.add_parser("me", help="Show my profile").set_defaults(
        fn=lambda _a: print(json.dumps(graph("base", "GET", "/me"), indent=2, ensure_ascii=False)))

    s = sub.add_parser("teams", help="List joined teams"); s.set_defaults(fn=list_teams)

    s = sub.add_parser("channels", help="List channels of a team")
    s.add_argument("--team", required=True); s.set_defaults(fn=list_channels)

    s = sub.add_parser("channel-create", help="Create a channel")
    s.add_argument("--team", required=True); s.add_argument("--name", required=True)
    s.add_argument("--description"); s.add_argument("--private", action="store_true")
    s.set_defaults(fn=create_channel)

    s = sub.add_parser("messages", help="Read channel messages")
    s.add_argument("--team", required=True); s.add_argument("--channel", required=True)
    s.add_argument("--limit", type=int, default=20); s.add_argument("--query")
    s.set_defaults(fn=list_messages)

    s = sub.add_parser("send", help="Send a message to a channel")
    s.add_argument("--team", required=True); s.add_argument("--channel", required=True)
    s.add_argument("--text", required=True); s.set_defaults(fn=send_message)

    s = sub.add_parser("chats", help="List my chats")
    s.add_argument("--limit", type=int, default=30); s.set_defaults(fn=list_chats)

    s = sub.add_parser("chat-messages", help="Read messages of a chat")
    s.add_argument("--chat", required=True); s.add_argument("--limit", type=int, default=20)
    s.set_defaults(fn=chat_messages)

    s = sub.add_parser("chat-send", help="Send a message to a chat")
    s.add_argument("--chat", required=True); s.add_argument("--text", required=True)
    s.set_defaults(fn=chat_send)

    s = sub.add_parser("chat-create", help="Create 1:1 or group chat")
    s.add_argument("--users", required=True, help="comma-separated UPNs or user ids")
    s.add_argument("--topic"); s.set_defaults(fn=chat_create)

    s = sub.add_parser("events", help="List calendar events / meetings")
    s.add_argument("--days", type=int, default=7); s.set_defaults(fn=events_list)

    s = sub.add_parser("event-create", help="Create calendar event with Teams link")
    s.add_argument("--title", required=True)
    s.add_argument("--start", required=True, help="YYYY-MM-DD HH:MM")
    s.add_argument("--end", required=True)
    s.add_argument("--attendees", help="comma-separated emails")
    s.add_argument("--body"); s.set_defaults(fn=event_create)

    s = sub.add_parser("meeting-create", help="Create a Teams online meeting")
    s.add_argument("--title", required=True)
    s.add_argument("--start", required=True); s.add_argument("--end", required=True)
    s.add_argument("--attendees"); s.set_defaults(fn=meeting_create)

    s = sub.add_parser("files", help="List OneDrive files")
    s.add_argument("--path", default=""); s.set_defaults(fn=files_list)

    s = sub.add_parser("file-upload", help="Upload a file to OneDrive")
    s.add_argument("--local", required=True); s.add_argument("--remote")
    s.set_defaults(fn=file_upload)

    s = sub.add_parser("file-download", help="Download a OneDrive file")
    s.add_argument("--remote", required=True); s.add_argument("--local", required=True)
    s.set_defaults(fn=file_download)

    s = sub.add_parser("presence", help="Show presence (me or a user)")
    s.add_argument("--user"); s.set_defaults(fn=presence)

    s = sub.add_parser("todo-lists", help="List To Do lists"); s.set_defaults(fn=todo_lists)

    s = sub.add_parser("todo-tasks", help="List tasks in a To Do list")
    s.add_argument("--list", required=True); s.set_defaults(fn=todo_tasks)

    s = sub.add_parser("todo-add", help="Add a task to a To Do list")
    s.add_argument("--list", required=True); s.add_argument("--title", required=True)
    s.add_argument("--due"); s.set_defaults(fn=todo_add)

    s = sub.add_parser("todo-complete", help="Mark a task as completed")
    s.add_argument("--list", required=True); s.add_argument("--task", required=True)
    s.set_defaults(fn=todo_complete)

    args = p.parse_args()
    args.fn(args)


if __name__ == "__main__":
    main()
