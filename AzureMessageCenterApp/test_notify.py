"""
Simulates a new message notification using the first entry from messages_cache.json.
Run this while app.py is NOT running (or in a separate terminal).
"""
import json
import time
from pathlib import Path
from winotify import Notification, audio

BASE_DIR = Path(__file__).parent
CACHE = BASE_DIR / "messages_cache.json"

APP_ID = "Azure Message Center Monitor"

def severity_emoji(msg):
    s = (msg.get("severity") or "").lower()
    if s == "critical": return "🔴"
    if s == "high":     return "🟠"
    if msg.get("isMajorChange"): return "🟡"
    return "🔵"

def notify(title, message, duration="short"):
    toast = Notification(app_id=APP_ID, title=title, msg=message[:256], duration=duration)
    toast.set_audio(audio.Default, loop=False)
    toast.show()

if not CACHE.exists():
    print("No messages_cache.json found — run app.py and do a Check Now first.")
    raise SystemExit

messages = json.loads(CACHE.read_text(encoding="utf-8"))
if not messages:
    print("Cache is empty.")
    raise SystemExit

msg = messages[0]
emoji   = severity_emoji(msg)
services = ", ".join(msg.get("services", [])) or "General"
tags     = ", ".join(msg.get("tags", []))
detail   = services + (f" | {tags}" if tags else "")

print(f"Simulating notification for: {msg.get('title')}")
notify(
    title=f"{emoji} {msg.get('title', 'New Message Center Post')}",
    message=detail,
)
print("Toast sent. Check your notification area.")

# Also simulate the "many new messages" banner
time.sleep(3)
print("Simulating bulk notification (>5 new messages)...")
notify(
    title="Azure Message Center – 8 New Posts",
    message="8 new Message Center posts. Open the viewer to review.",
)
print("Done.")
