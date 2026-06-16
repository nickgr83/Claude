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

from datetime import datetime
from bs4 import BeautifulSoup
import feedparser
import msal
import requests
import pystray
from PIL import Image, ImageDraw, ImageTk
from winotify import Notification, audio

# ── Paths ────────────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
CONFIG_FILE = BASE_DIR / "config.json"
SEEN_FILE = BASE_DIR / "seen_messages.json"
SEEN_BLOG_FILE = BASE_DIR / "seen_blog.json"
MESSAGES_CACHE_FILE = BASE_DIR / "messages_cache.json"
TOKEN_CACHE_FILE = BASE_DIR / "token_cache.bin"
LOG_FILE = BASE_DIR / "app.log"

# Each entry: (source_label, [feed_urls_in_priority_order], filter_tag_or_None)
# filter_tag=None means accept all entries (feed is already pre-filtered).
BLOG_SOURCES = [
    (
        "Azure Blog",
        [
            "https://azure.microsoft.com/en-us/blog/content-type/announcements/feed/",
            "https://azure.microsoft.com/en-us/blog/feed/",
        ],
        "announcements",   # filter tag when using the main feed
    ),
    (
        "Windows Blog",
        [
            "https://blogs.windows.com/feed/",
        ],
        None,              # accept all posts from this feed
    ),
]

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


def load_seen_blog() -> set:
    if SEEN_BLOG_FILE.exists():
        with open(SEEN_BLOG_FILE, encoding="utf-8") as f:
            return set(json.load(f))
    return set()


def save_seen_blog(seen: set):
    with open(SEEN_BLOG_FILE, "w", encoding="utf-8") as f:
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

def notify(title: str, message: str, duration: str = "short", launch_url: str = ""):
    try:
        toast = Notification(
            app_id=APP_ID,
            title=title,
            msg=message[:256],
            duration=duration,
            launch=launch_url or "https://admin.microsoft.com/Adminportal/Home#/MessageCenter",
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
    msgs = resp.json().get("value", [])
    for m in msgs:
        m.setdefault("source", "Message Center")
        m.setdefault("link", "https://admin.microsoft.com/Adminportal/Home#/MessageCenter")
    return msgs


def _fetch_one_blog(source_label: str, feed_urls: list, filter_tag: str | None) -> list:
    """Try each URL in order, return parsed entries for the first that succeeds."""
    headers = {"User-Agent": "AzureMessageCenterMonitor/1.0"}
    last_exc = None

    for feed_url in feed_urls:
        try:
            feed = feedparser.parse(feed_url, request_headers=headers)
        except Exception as exc:
            last_exc = exc
            log.warning("%s feed %s failed: %s", source_label, feed_url, exc)
            continue

        if feed.get("bozo") and not feed.entries:
            last_exc = feed.bozo_exception
            log.warning("%s feed %s bozo (no entries): %s", source_label, feed_url, last_exc)
            continue

        results = []
        for entry in feed.entries:
            if filter_tag:
                entry_tag_vals = []
                for t in entry.get("tags", []):
                    for field in ("term", "label", "scheme"):
                        val = (t.get(field) or "").lower()
                        if val:
                            entry_tag_vals.append(val)
                if not any(filter_tag in v for v in entry_tag_vals):
                    continue

            published = entry.get("published", entry.get("updated", ""))
            entry_tags_clean = [
                t.get("term") or t.get("label", "")
                for t in entry.get("tags", [])
                if (t.get("term") or t.get("label", "")).lower() != (filter_tag or "")
            ]
            results.append({
                "id": f"blog:{source_label}:{entry.get('id', entry.get('link', ''))}",
                "title": entry.get("title", "(no title)"),
                "lastModifiedDateTime": published,
                "services": [source_label],
                "tags": entry_tags_clean,
                "severity": None,
                "isMajorChange": False,
                "body": {"content": entry.get("summary", ""), "contentType": "html"},
                "link": entry.get("link", ""),
                "source": source_label,
            })

        log.info("Fetched %d post(s) from %s (%s)", len(results), source_label, feed_url)
        return results

    log.warning("All feeds for %s failed. Last error: %s", source_label, last_exc)
    return []


def fetch_blog_announcements() -> list:
    """Fetch all configured blog sources and return combined results."""
    results = []
    for source_label, feed_urls, filter_tag in BLOG_SOURCES:
        try:
            results.extend(_fetch_one_blog(source_label, feed_urls, filter_tag))
        except Exception as exc:
            log.warning("Blog source %s error: %s", source_label, exc)
    return results


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


_ARTICLE_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AzureMessageCenterMonitor/1.0"}
_ARTICLE_SELECTORS = [
    "article",
    "[class*='article-body']",
    "[class*='post-content']",
    "[class*='entry-content']",
    "[class*='blog-content']",
    "[class*='article-content']",
    "main",
    "body",
]


def fetch_full_article(url: str) -> str:
    """Download a blog post and return plain text of the article body."""
    resp = requests.get(url, headers=_ARTICLE_HEADERS, timeout=20)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "lxml")
    for tag in soup(["script", "style", "nav", "header", "footer",
                     "aside", "form", "noscript", "iframe", "img"]):
        tag.decompose()

    content_tag = next(
        (soup.select_one(sel) for sel in _ARTICLE_SELECTORS if soup.select_one(sel)),
        soup,
    )
    return html_to_text(str(content_tag))


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

def poll_once(cfg: dict, seen: set, seen_blog: set, first_run: bool, cache_ref: list) -> tuple:
    source_counts = {}   # source_label -> count (for status bar)

    # ── Message Center ────────────────────────────────────────────────────────
    mc_messages = []
    try:
        token = get_token(cfg)
        mc_messages = fetch_messages(token)
        log.info("Fetched %d messages from Message Center", len(mc_messages))
        source_counts["Message Center"] = len(mc_messages)
    except Exception as exc:
        log.error("Message Center poll failed: %s", exc)
        notify("Azure Message Center – Error", str(exc)[:200])
        source_counts["Message Center"] = "ERR"

    # ── Blogs ─────────────────────────────────────────────────────────────────
    blog_messages = []
    try:
        blog_messages = fetch_blog_announcements()
        for label, _, _ in BLOG_SOURCES:
            count = sum(1 for m in blog_messages if m.get("source") == label)
            source_counts[label] = count
    except Exception as exc:
        log.warning("Blog fetch failed (non-fatal): %s", exc)
        for label, _, _ in BLOG_SOURCES:
            source_counts[label] = "ERR"

    # Update combined cache sorted by date descending
    all_messages = mc_messages + blog_messages
    all_messages.sort(key=lambda m: m.get("lastModifiedDateTime", ""), reverse=True)
    cache_ref.clear()
    cache_ref.extend(all_messages)
    save_messages_cache(all_messages)

    max_notify = cfg.get("max_notifications_per_poll", 5)

    # ── Notify: Message Center new items ─────────────────────────────────────
    new_mc_ids = [m["id"] for m in mc_messages if m["id"] not in seen]
    if new_mc_ids and not first_run:
        new_mc = [m for m in mc_messages if m["id"] in new_mc_ids]
        portal = "https://admin.microsoft.com/Adminportal/Home#/MessageCenter"
        if len(new_mc) > max_notify:
            notify(f"Azure Message Center – {len(new_mc)} New Posts",
                   f"{len(new_mc)} new posts. Click to open the portal.",
                   launch_url=portal)
        else:
            for msg in new_mc[:max_notify]:
                emoji = severity_emoji(msg)
                services = ", ".join(msg.get("services", [])) or "General"
                tags = ", ".join(msg.get("tags", []))
                detail = services + (f" | {tags}" if tags else "")
                notify(f"{emoji} {msg.get('title', 'New Message Center Post')}", detail,
                       launch_url=msg.get("link", portal))
                time.sleep(0.5)
    elif new_mc_ids and first_run:
        log.info("First run — marking %d MC messages as seen.", len(new_mc_ids))

    # ── Notify: Blog new items ────────────────────────────────────────────────
    new_blog_ids = [m["id"] for m in blog_messages if m["id"] not in seen_blog]
    if new_blog_ids and not first_run:
        new_blog = [m for m in blog_messages if m["id"] in new_blog_ids]
        blog_index = "https://azure.microsoft.com/en-us/blog/content-type/announcements/"
        if len(new_blog) > max_notify:
            notify(f"Blog – {len(new_blog)} New Posts",
                   f"{len(new_blog)} new blog posts. Click to open.",
                   launch_url=blog_index)
        else:
            for msg in new_blog[:max_notify]:
                source = msg.get("source", "Blog")
                notify(f"📢 {msg.get('title', f'New {source} Post')}",
                       f"{source} — click to read",
                       launch_url=msg.get("link", blog_index))
                time.sleep(0.5)
    elif new_blog_ids and first_run:
        log.info("First run — marking %d blog posts as seen.", len(new_blog_ids))

    if not new_mc_ids and not new_blog_ids:
        log.info("No new messages.")

    last_checked = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    status = "Last checked: " + last_checked + "  |  " + "  ".join(
        f"{lbl}: {cnt}" for lbl, cnt in source_counts.items()
    )
    return seen | set(new_mc_ids), seen_blog | set(new_blog_ids), last_checked, status


# ── Message Viewer Window ─────────────────────────────────────────────────────

class MessageViewer:
    # column index → message field key for sorting
    _SORT_KEY = {
        "date":     lambda m: m.get("lastModifiedDateTime", ""),
        "title":    lambda m: m.get("title", "").lower(),
        "services": lambda m: ", ".join(m.get("services", [])).lower(),
        "tags":     lambda m: ", ".join(m.get("tags", [])).lower(),
        "source":   lambda m: m.get("source", "").lower(),
    }

    def __init__(self, messages: list, last_checked: str = ""):
        self.messages = messages
        self.last_checked = last_checked
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

        cols = ("sev", "date", "source", "title", "services", "tags")
        self.tree = ttk.Treeview(top, columns=cols, show="headings", selectmode="browse")

        headings = {"sev": "", "date": "Date ▼", "source": "Source",
                    "title": "Title", "services": "Services", "tags": "Tags"}
        for col, text in headings.items():
            self.tree.heading(col, text=text,
                              command=lambda c=col: self._sort_by(c))
        self.tree.column("sev",      width=30,  stretch=False, anchor="center")
        self.tree.column("date",     width=90,  stretch=False)
        self.tree.column("source",   width=120, stretch=False)
        self.tree.column("title",    width=400)
        self.tree.column("services", width=200)
        self.tree.column("tags",     width=170)

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

        # ── Status bar ────────────────────────────────────────────────────────
        status_bar = tk.Frame(self.root, bg="#111122", pady=3)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)
        self._status_var = tk.StringVar()
        checked_txt = f"Last checked: {self.last_checked}" if self.last_checked else "Not yet checked"
        self._status_var.set(checked_txt)
        tk.Label(status_bar, textvariable=self._status_var, bg="#111122",
                 fg="#6080a0", font=("Segoe UI", 8), anchor="w",
                 padx=8).pack(side=tk.LEFT)

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
        self._refresh_tree()

    def _clear_filters(self):
        self._f_title.set("")
        self._f_services.set("")
        self._f_tags.set("")

    def _sort_by(self, col: str):
        if col in ("sev",):
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
        labels = {"date": "Date", "source": "Source", "title": "Title",
                  "services": "Services", "tags": "Tags"}
        for col, label in labels.items():
            suffix = arrows[self._sort_asc] if col == self._sort_col else ""
            self.tree.heading(col, text=label + suffix)

        # Repopulate treeview
        self.tree.delete(*self.tree.get_children())
        for msg in self._filtered:
            self.tree.insert("", tk.END, iid=msg["id"], values=(
                severity_emoji(msg),
                fmt_date(msg.get("lastModifiedDateTime", "")),
                msg.get("source", "Message Center"),
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
        is_blog = msg.get("source", "Message Center") != "Message Center"
        link = msg.get("link", "")

        if is_blog and link:
            if msg.get("_full_body"):
                self._render_body(msg["_full_body"], link)
            else:
                body_html = (msg.get("body") or {}).get("content", "")
                preview = html_to_text(body_html) if body_html else ""
                self._render_body("⏳ Loading full article…\n\n" + preview, link)
                threading.Thread(target=self._load_full_article,
                                 args=(msg, link), daemon=True).start()
        else:
            body_html = (msg.get("body") or {}).get("content", "")
            body_text = html_to_text(body_html) if body_html else "(No content)"
            self._render_body(body_text, None)

    def _load_full_article(self, msg: dict, url: str):
        try:
            msg["_full_body"] = fetch_full_article(url)
        except Exception as exc:
            log.warning("Full article fetch failed for %s: %s", url, exc)
            body_html = (msg.get("body") or {}).get("content", "")
            msg["_full_body"] = html_to_text(body_html) if body_html else "(Could not load article)"
        self.root.after(0, lambda: self._render_body(msg["_full_body"], url))

    def _render_body(self, body_text: str, link: str | None):
        self.detail_text.configure(state=tk.NORMAL)
        self.detail_text.delete("1.0", tk.END)
        self.detail_text.insert(tk.END, body_text)

        if link:
            self.detail_text.insert(tk.END, "\n\n")
            self.detail_text.insert(tk.END, "─" * 60 + "\n")
            self.detail_text.insert(tk.END, "Read original post: ", "link_label")
            self.detail_text.insert(tk.END, link, ("link", link))
            self.detail_text.tag_configure("link_label", foreground="#a0a0c0")
            self.detail_text.tag_configure("link", foreground="#60cdff", underline=True)
            self.detail_text.tag_bind("link", "<Button-1>",
                                      lambda e, u=link: self._open_link(u))
            self.detail_text.tag_bind("link", "<Enter>",
                                      lambda e: self.detail_text.configure(cursor="hand2"))
            self.detail_text.tag_bind("link", "<Leave>",
                                      lambda e: self.detail_text.configure(cursor=""))

        self.detail_text.configure(state=tk.DISABLED)
        self.detail_text.yview_moveto(0)

    def _open_link(self, url: str):
        import subprocess
        subprocess.Popen(["rundll32", "url.dll,FileProtocolHandler", url], shell=False)

    def show(self):
        self.root.mainloop()


def open_viewer(messages: list, last_checked: str = ""):
    """Open the viewer in its own thread (tkinter needs the main thread or its own)."""
    def _run():
        try:
            viewer = MessageViewer(list(messages), last_checked)
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
        self.seen_blog = load_seen_blog()
        self.running = True
        self.first_run = len(self.seen) == 0 and len(self.seen_blog) == 0
        self.icon = None
        self._messages: list = load_messages_cache()
        self._last_checked: str = ""
        self._status: str = ""
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
            self.seen, self.seen_blog, self._last_checked, self._status = poll_once(
                self.cfg, self.seen, self.seen_blog, self.first_run, self._messages)
            save_seen(self.seen)
            save_seen_blog(self.seen_blog)
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
        self.seen, self.seen_blog, self._last_checked, self._status = poll_once(
            self.cfg, self.seen, self.seen_blog, False, self._messages)
        save_seen(self.seen)
        save_seen_blog(self.seen_blog)

    def _view_messages(self, icon, item):
        if not self._messages:
            notify("Azure Message Center", "No messages loaded yet. Try Check Now first.")
            return
        open_viewer(self._messages, self._status or self._last_checked)

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
