"""
Azure Message Center Monitor
Polls Microsoft Graph API for new Message Center posts and shows Windows toast notifications.
"""

import html
import json
import re
import time
import threading
import sys
import logging
import tkinter as tk
from tkinter import ttk
from pathlib import Path

import msal
import requests
import pystray
from PIL import Image, ImageDraw, ImageTk
from winotify import Notification, audio

# ── Paths ────────────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
CONFIG_FILE = BASE_DIR / "config.json"
SEEN_FILE = BASE_DIR / "seen_messages.json"
MESSAGES_CACHE_FILE = BASE_DIR / "messages_cache.json"
TOKEN_CACHE_FILE = BASE_DIR / "token_cache.bin"
LOG_FILE = BASE_DIR / "app.log"

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)

# ── Graph API ────────────────────────────────────────────────────────────────
GRAPH_SCOPES = ["https://graph.microsoft.com/ServiceMessage.Read.All"]
MESSAGES_URL = (
    "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages"
    "?$orderby=lastModifiedDateTime desc&$top=50"
    "&$select=id,title,lastModifiedDateTime,startDateTime,services,"
    "isMajorChange,severity,tags,body"
)

APP_ID = "Azure Message Center Monitor"


# ── Config & storage ──────────────────────────────────────────────────────────

def load_config() -> dict:
    with open(CONFIG_FILE, encoding="utf-8") as f:
        cfg = json.load(f)
    if cfg.get("client_id") == "YOUR_CLIENT_ID_HERE":
        raise ValueError(
            "client_id not set in config.json. "
            "Register an Azure AD app with ServiceMessage.Read.All permission "
            "and paste the Application (client) ID into config.json."
        )
    return cfg


def load_seen() -> set:
    if SEEN_FILE.exists():
        with open(SEEN_FILE, encoding="utf-8") as f:
            return set(json.load(f))
    return set()


def save_seen(seen: set):
    with open(SEEN_FILE, "w", encoding="utf-8") as f:
        json.dump(list(seen), f)


def load_messages_cache() -> list:
    if MESSAGES_CACHE_FILE.exists():
        with open(MESSAGES_CACHE_FILE, encoding="utf-8") as f:
            return json.load(f)
    return []


def save_messages_cache(messages: list):
    with open(MESSAGES_CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(messages, f, ensure_ascii=False, indent=2)


# ── Auth ──────────────────────────────────────────────────────────────────────

def build_msal_app(cfg: dict):
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_FILE.exists():
        cache.deserialize(TOKEN_CACHE_FILE.read_text(encoding="utf-8"))
    app = msal.PublicClientApplication(
        cfg["client_id"],
        authority=f"https://login.microsoftonline.com/{cfg['tenant_id']}",
        token_cache=cache,
    )
    return app, cache


def get_token(cfg: dict) -> str:
    msal_app, cache = build_msal_app(cfg)
    accounts = msal_app.get_accounts()

    result = None
    if accounts:
        result = msal_app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])

    if not result:
        log.info("No cached token — opening browser for interactive sign-in...")
        result = msal_app.acquire_token_interactive(
            scopes=GRAPH_SCOPES,
            login_hint=cfg.get("login_hint"),
        )

    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description', result)}")

    if cache.has_state_changed:
        TOKEN_CACHE_FILE.write_text(cache.serialize(), encoding="utf-8")

    return result["access_token"]


# ── Notifications ─────────────────────────────────────────────────────────────

def notify(title: str, message: str, duration: str = "short"):
    try:
        toast = Notification(
            app_id=APP_ID,
            title=title,
            msg=message[:256],
            duration=duration,
        )
        toast.set_audio(audio.Default, loop=False)
        toast.show()
    except Exception as exc:
        log.warning("Toast notification failed: %s", exc)


# ── Graph fetch ───────────────────────────────────────────────────────────────

def fetch_messages(token: str) -> list:
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    resp = requests.get(MESSAGES_URL, headers=headers, timeout=30)
    if not resp.ok:
        log.error("Graph API %s: %s", resp.status_code, resp.text)
    resp.raise_for_status()
    return resp.json().get("value", [])


def html_to_text(raw: str) -> str:
    """Strip HTML tags and decode entities for plain-text display."""
    if not raw:
        return ""
    text = re.sub(r"<br\s*/?>", "\n", raw, flags=re.IGNORECASE)
    text = re.sub(r"</p>", "\n\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<li>", "\n• ", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = html.unescape(text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def severity_emoji(msg: dict) -> str:
    s = (msg.get("severity") or "").lower()
    if s == "critical":
        return "🔴"
    if s == "high":
        return "🟠"
    if msg.get("isMajorChange"):
        return "🟡"
    return "🔵"


def fmt_date(iso: str) -> str:
    if not iso:
        return ""
    try:
        dt = iso[:10]  # YYYY-MM-DD
        return dt
    except Exception:
        return iso


# ── Polling ───────────────────────────────────────────────────────────────────

def poll_once(cfg: dict, seen: set, first_run: bool, cache_ref: list) -> set:
    try:
        token = get_token(cfg)
        messages = fetch_messages(token)
        log.info("Fetched %d messages from Message Center", len(messages))
    except Exception as exc:
        log.error("Poll failed: %s", exc)
        notify("Azure Message Center – Error", str(exc)[:200])
        return seen

    # Update local message cache (keep latest 50)
    cache_ref.clear()
    cache_ref.extend(messages)
    save_messages_cache(messages)

    new_ids = [m["id"] for m in messages if m["id"] not in seen]

    if not new_ids:
        log.info("No new messages.")
        return seen

    log.info("%d new message(s) found.", len(new_ids))

    if first_run:
        log.info("First run — marking %d existing messages as seen.", len(new_ids))
        return seen | set(new_ids)

    max_notify = cfg.get("max_notifications_per_poll", 5)
    new_messages = [m for m in messages if m["id"] in new_ids]

    if len(new_messages) > max_notify:
        notify(
            title=f"Azure Message Center – {len(new_messages)} New Posts",
            message=f"{len(new_messages)} new posts. Open the viewer to review.",
        )
    else:
        for msg in new_messages[:max_notify]:
            emoji = severity_emoji(msg)
            services = ", ".join(msg.get("services", [])) or "General"
            tags = ", ".join(msg.get("tags", []))
            detail = services + (f" | {tags}" if tags else "")
            notify(
                title=f"{emoji} {msg.get('title', 'New Message Center Post')}",
                message=detail,
            )
            time.sleep(0.5)

    return seen | set(new_ids)


# ── Message Viewer Window ─────────────────────────────────────────────────────

class MessageViewer:
    # column index → message field key for sorting
    _SORT_KEY = {
        "date":     lambda m: m.get("lastModifiedDateTime", ""),
        "title":    lambda m: m.get("title", "").lower(),
        "services": lambda m: ", ".join(m.get("services", [])).lower(),
        "tags":     lambda m: ", ".join(m.get("tags", [])).lower(),
    }

    def __init__(self, messages: list):
        self.messages = messages
        self._filtered: list = list(messages)
        self._sort_col: str = "date"
        self._sort_asc: bool = False
        self.root = tk.Tk()
        self._build()

    def _build(self):
        root = self.root
        root.title("Azure Message Center")
        root.geometry("1100x720")
        root.configure(bg="#1e1e2e")

        # Title-bar icon — must be kept alive on self to prevent GC
        self._tk_icon = ImageTk.PhotoImage(make_azure_icon(32))
        root.iconphoto(True, self._tk_icon)

        style = ttk.Style(root)
        style.theme_use("clam")
        style.configure("Treeview",
                        background="#2a2a3e", foreground="#e0e0f0",
                        rowheight=28, fieldbackground="#2a2a3e",
                        font=("Segoe UI", 9))
        style.configure("Treeview.Heading",
                        background="#0078d4", foreground="white",
                        font=("Segoe UI", 9, "bold"))
        style.map("Treeview", background=[("selected", "#0078d4")])
        style.configure("TPanedwindow", background="#0078d4")
        style.configure("Filter.TEntry", fieldbackground="#2a2a3e",
                        foreground="#e0e0f0", insertcolor="#e0e0f0")

        # ── Filter bar ────────────────────────────────────────────────────────
        filter_bar = tk.Frame(root, bg="#1e1e2e")
        filter_bar.pack(fill=tk.X, padx=10, pady=(8, 2))

        lbl_kw = {"bg": "#1e1e2e", "fg": "#a0a0c0", "font": ("Segoe UI", 9)}
        ent_kw = {"bg": "#2a2a3e", "fg": "#e0e0f0", "insertbackground": "#e0e0f0",
                  "relief": "flat", "font": ("Segoe UI", 9), "width": 28}

        tk.Label(filter_bar, text="Title:", **lbl_kw).pack(side=tk.LEFT, padx=(0, 4))
        self._f_title = tk.StringVar()
        tk.Entry(filter_bar, textvariable=self._f_title, **ent_kw).pack(side=tk.LEFT, padx=(0, 12))

        tk.Label(filter_bar, text="Services:", **lbl_kw).pack(side=tk.LEFT, padx=(0, 4))
        self._f_services = tk.StringVar()
        tk.Entry(filter_bar, textvariable=self._f_services, **ent_kw).pack(side=tk.LEFT, padx=(0, 12))

        tk.Label(filter_bar, text="Tags:", **lbl_kw).pack(side=tk.LEFT, padx=(0, 4))
        self._f_tags = tk.StringVar()
        tk.Entry(filter_bar, textvariable=self._f_tags, **ent_kw).pack(side=tk.LEFT, padx=(0, 12))

        tk.Button(filter_bar, text="Clear", bg="#3a3a5e", fg="#e0e0f0",
                  relief="flat", font=("Segoe UI", 9), padx=8,
                  command=self._clear_filters).pack(side=tk.LEFT)

        self._count_var = tk.StringVar()
        tk.Label(filter_bar, textvariable=self._count_var, **lbl_kw).pack(side=tk.RIGHT)

        for var in (self._f_title, self._f_services, self._f_tags):
            var.trace_add("write", lambda *_: self._apply_filters())

        # ── Vertical PanedWindow ──────────────────────────────────────────────
        paned = ttk.PanedWindow(root, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=(4, 10))

        # ── Top pane: message list ────────────────────────────────────────────
        top = tk.Frame(paned, bg="#1e1e2e")
        paned.add(top, weight=1)

        cols = ("sev", "date", "title", "services", "tags")
        self.tree = ttk.Treeview(top, columns=cols, show="headings", selectmode="browse")

        headings = {"sev": "", "date": "Date ▼", "title": "Title",
                    "services": "Services", "tags": "Tags"}
        for col, text in headings.items():
            self.tree.heading(col, text=text,
                              command=lambda c=col: self._sort_by(c))
        self.tree.column("sev",      width=30,  stretch=False, anchor="center")
        self.tree.column("date",     width=90,  stretch=False)
        self.tree.column("title",    width=480)
        self.tree.column("services", width=220)
        self.tree.column("tags",     width=200)

        vsb = ttk.Scrollbar(top, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # ── Bottom pane: detail ───────────────────────────────────────────────
        bottom = tk.Frame(paned, bg="#1e1e2e")
        paned.add(bottom, weight=1)

        self.title_var = tk.StringVar()
        tk.Label(bottom, textvariable=self.title_var, bg="#1e1e2e",
                 fg="#60cdff", font=("Segoe UI", 11, "bold"),
                 anchor="w", wraplength=1060, justify="left").pack(fill=tk.X, pady=(4, 2))

        text_frame = tk.Frame(bottom, bg="#1e1e2e")
        text_frame.pack(fill=tk.BOTH, expand=True)

        self.detail_text = tk.Text(
            text_frame, bg="#2a2a3e", fg="#e0e0f0",
            font=("Segoe UI", 9), relief="flat", wrap=tk.WORD,
            padx=8, pady=6, state=tk.DISABLED,
        )
        detail_vsb = ttk.Scrollbar(text_frame, orient="vertical",
                                   command=self.detail_text.yview)
        self.detail_text.configure(yscrollcommand=detail_vsb.set)
        self.detail_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        detail_vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # Initial population
        self._apply_filters()

    # ── Filter & sort ─────────────────────────────────────────────────────────

    def _apply_filters(self):
        ft = self._f_title.get().lower()
        fs = self._f_services.get().lower()
        fg = self._f_tags.get().lower()

        self._filtered = [
            m for m in self.messages
            if ft in m.get("title", "").lower()
            and fs in ", ".join(m.get("services", [])).lower()
            and fg in ", ".join(m.get("tags", [])).lower()
        ]
        self._sort_col  # keep current sort
        self._refresh_tree()

    def _clear_filters(self):
        self._f_title.set("")
        self._f_services.set("")
        self._f_tags.set("")

    def _sort_by(self, col: str):
        if col == "sev":
            return
        if self._sort_col == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = col
            self._sort_asc = True
        self._refresh_tree()

    def _refresh_tree(self):
        key_fn = self._SORT_KEY.get(self._sort_col)
        if key_fn:
            self._filtered.sort(key=key_fn, reverse=not self._sort_asc)

        # Update heading arrows
        arrows = {True: " ▲", False: " ▼"}
        for col in ("date", "title", "services", "tags"):
            label = {"date": "Date", "title": "Title",
                     "services": "Services", "tags": "Tags"}[col]
            suffix = arrows[self._sort_asc] if col == self._sort_col else ""
            self.tree.heading(col, text=label + suffix)

        # Repopulate treeview
        self.tree.delete(*self.tree.get_children())
        for msg in self._filtered:
            self.tree.insert("", tk.END, iid=msg["id"], values=(
                severity_emoji(msg),
                fmt_date(msg.get("lastModifiedDateTime", "")),
                msg.get("title", ""),
                ", ".join(msg.get("services", [])),
                ", ".join(msg.get("tags", [])),
            ))

        self._count_var.set(f"{len(self._filtered)} / {len(self.messages)} messages")

        # Auto-select first visible row
        children = self.tree.get_children()
        if children:
            self.tree.selection_set(children[0])
            self.tree.focus(children[0])
            msg = next((m for m in self._filtered if m["id"] == children[0]), None)
            if msg:
                self._show_message(msg)
        else:
            self.title_var.set("")
            self.detail_text.configure(state=tk.NORMAL)
            self.detail_text.delete("1.0", tk.END)
            self.detail_text.configure(state=tk.DISABLED)

    # ── Selection & detail ────────────────────────────────────────────────────

    def _on_select(self, _event):
        sel = self.tree.selection()
        if not sel:
            return
        msg = next((m for m in self._filtered if m["id"] == sel[0]), None)
        if msg:
            self._show_message(msg)

    def _show_message(self, msg: dict):
        self.title_var.set(
            f"{severity_emoji(msg)}  {msg.get('title', '')}  "
            f"({fmt_date(msg.get('lastModifiedDateTime', ''))})"
        )
        body_html = (msg.get("body") or {}).get("content", "")
        body_text = html_to_text(body_html) if body_html else "(No content)"

        self.detail_text.configure(state=tk.NORMAL)
        self.detail_text.delete("1.0", tk.END)
        self.detail_text.insert(tk.END, body_text)
        self.detail_text.configure(state=tk.DISABLED)
        self.detail_text.yview_moveto(0)

    def show(self):
        self.root.mainloop()


def open_viewer(messages: list):
    """Open the viewer in its own thread (tkinter needs the main thread or its own)."""
    def _run():
        try:
            viewer = MessageViewer(list(messages))
            viewer.show()
        except Exception as exc:
            log.error("Viewer error: %s", exc)
    t = threading.Thread(target=_run, daemon=True)
    t.start()


# ── System Tray ───────────────────────────────────────────────────────────────

def make_azure_icon(size: int = 64) -> Image.Image:
    """
    Draws the Microsoft Azure 'A' logo: two overlapping trapezoids
    (dark-left / light-right) with a reflected bottom strip, matching
    the real Azure brand geometry.
    """
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    s = size

    # Colours from the official Azure brand palette
    DARK  = (0,  72, 153, 255)   # #004899 – left face
    LIGHT = (0, 163, 237, 255)   # #00A3ED – right face
    MID   = (0, 114, 198, 255)   # #0072C6 – bottom reflection

    # Left trapezoid — leans right toward the apex
    left = [
        (s * 0.08, s * 0.88),   # bottom-left
        (s * 0.36, s * 0.12),   # apex-left
        (s * 0.52, s * 0.12),   # apex-right
        (s * 0.29, s * 0.88),   # bottom-right
    ]
    d.polygon(left, fill=DARK)

    # Right trapezoid — mirror image of the left
    right = [
        (s * 0.52, s * 0.12),   # apex-left
        (s * 0.92, s * 0.88),   # bottom-right
        (s * 0.71, s * 0.88),   # bottom-left
        (s * 0.36, s * 0.12),   # apex-right  (shared with left apex)
    ]
    d.polygon(right, fill=LIGHT)

    # Reflection strip across the lower portion (gives depth)
    strip = [
        (s * 0.17, s * 0.70),
        (s * 0.50, s * 0.52),
        (s * 0.71, s * 0.70),
        (s * 0.50, s * 0.70),
    ]
    d.polygon(strip, fill=MID)

    return img


# Keep old name as alias so the rest of the code doesn't need changing
def make_icon() -> Image.Image:
    return make_azure_icon(64)


class MonitorApp:
    def __init__(self):
        self.cfg = load_config()
        self.seen = load_seen()
        self.running = True
        self.first_run = len(self.seen) == 0
        self.icon = None
        self._messages: list = load_messages_cache()
        self._poll_thread = None

    def start(self):
        self._poll_thread = threading.Thread(target=self._poll_loop, daemon=True)
        self._poll_thread.start()

        icon_image = make_icon()
        self.icon = pystray.Icon(
            "AzureMsgCenter",
            icon_image,
            "Azure Message Center Monitor",
            menu=pystray.Menu(
                pystray.MenuItem("View Messages", self._view_messages, default=True),
                pystray.MenuItem("Check Now", self._check_now),
                pystray.MenuItem("Open in Portal", self._open_portal),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Quit", self._quit),
            ),
        )
        self.icon.run()

    def _poll_loop(self):
        log.info("Polling thread started (interval: %d min)", self.cfg["poll_interval_minutes"])
        while self.running:
            self.seen = poll_once(self.cfg, self.seen, self.first_run, self._messages)
            save_seen(self.seen)
            self.first_run = False
            interval = self.cfg["poll_interval_minutes"] * 60
            for _ in range(interval // 5):
                if not self.running:
                    break
                time.sleep(5)

    def _check_now(self, icon, item):
        threading.Thread(target=self._do_check_now, daemon=True).start()

    def _do_check_now(self):
        log.info("Manual check triggered.")
        self.seen = poll_once(self.cfg, self.seen, False, self._messages)
        save_seen(self.seen)

    def _view_messages(self, icon, item):
        if not self._messages:
            notify("Azure Message Center", "No messages loaded yet. Try Check Now first.")
            return
        open_viewer(self._messages)

    def _open_portal(self, icon, item):
        import subprocess
        subprocess.Popen(
            ["rundll32", "url.dll,FileProtocolHandler",
             "https://admin.microsoft.com/Adminportal/Home#/MessageCenter"],
            shell=False,
        )

    def _quit(self, icon, item):
        log.info("Quit requested.")
        self.running = False
        icon.stop()


# ── Entry Point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    try:
        app = MonitorApp()
        log.info("Starting Azure Message Center Monitor...")
        app.start()
    except ValueError as e:
        print(f"\nConfiguration error: {e}\n")
        input("Press Enter to exit...")
        sys.exit(1)
