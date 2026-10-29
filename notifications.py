"""Activity log shown to the admin (notifications.json).

Every entry is stored in the repo next to the other data files (GitHub
Contents API) with a local mirror, so the log survives redeploys like the
rest of the data. Writing an entry must never break or slow down the action
that triggered it: all failures are swallowed and the GitHub write runs on a
background thread.
"""
import json
import os
import threading
import time
import uuid
from datetime import datetime

FILE_NAME = "notifications.json"
MAX_EVENTS = 500
CACHE_TTL = 45

_lock = threading.Lock()
_recent = {}
_cache = {"at": 0.0, "data": {"events": [], "seen": {}}}

KIND_META = {
    "invoice": ("🧾", "#fd79a8"),
    "template": ("📄", "#6c5ce7"),
    "codes": ("🏷️", "#00cec9"),
    "form41": ("📋", "#74b9ff"),
    "vat": ("💰", "#fdcb6e"),
    "portal": ("📄", "#a29bfe"),
    "receipts": ("🧾", "#fd79a8"),
    "barcodes": ("📦", "#00b894"),
    "users": ("👥", "#e84393"),
    "gs1": ("📦", "#00b894"),
    "supplier": ("🏭", "#74b9ff"),
    "vat10": ("📑", "#fdcb6e"),
    "requests": ("📨", "#a29bfe"),
    "submit": ("🚀", "#00cec9"),
    "login": ("🔐", "#7878a0"),
    "other": ("⚙️", "#7878a0"),
}


def _path():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), FILE_NAME)


def _normalize(d):
    if isinstance(d, list):
        return {"events": [e for e in d if isinstance(e, dict)], "seen": {}}
    if isinstance(d, dict):
        ev = d.get("events")
        seen = d.get("seen")
        return {
            "events": [e for e in ev if isinstance(e, dict)] if isinstance(ev, list) else [],
            "seen": seen if isinstance(seen, dict) else {},
        }
    return {"events": [], "seen": {}}


def _local_read():
    try:
        with open(_path(), encoding="utf-8") as f:
            return _normalize(json.load(f))
    except Exception:
        return {"events": [], "seen": {}}


def _local_write(d):
    try:
        tmp = _path() + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2, default=str)
        os.replace(tmp, _path())
    except Exception:
        pass


def _gh_read():
    try:
        from github_storage import gh_read
        d = gh_read(FILE_NAME)
        return None if d is None else _normalize(d)
    except Exception:
        return None


def _gh_write(d):
    try:
        from github_storage import gh_write
        return bool(gh_write(FILE_NAME, d))
    except Exception:
        return False


def load(force=False):
    now = time.time()
    if not force and now - _cache["at"] < CACHE_TTL:
        return _cache["data"]
    d = _gh_read()
    if d is None:
        d = _local_read()
    _cache["at"] = now
    _cache["data"] = d
    return d


def _now():
    return datetime.now().isoformat(timespec="seconds")


def _current_user():
    try:
        import streamlit as st
        cu = st.session_state.get("current_user") or {}
        return cu.get("display_name") or cu.get("username") or "", cu.get("username", ""), cu.get("role", "")
    except Exception:
        return "", "", ""


def _page():
    try:
        import streamlit as st
        return st.session_state.get("selected_page", "") or ""
    except Exception:
        return ""


def _describe(items, names=()):
    out = []
    for it in items:
        if isinstance(it, dict):
            for k in names:
                v = str(it.get(k, "") or "").strip()
                if v:
                    out.append(v)
                    break
        elif str(it).strip():
            out.append(str(it).strip())
    txt = "، ".join(out)
    return txt[:160]


def _append(entry, sync=False):
    def work():
        try:
            with _lock:
                d = _gh_read()
                if d is None:
                    d = _local_read()
                d["events"].append(entry)
                d["events"] = d["events"][-MAX_EVENTS:]
                if not _gh_write(d):
                    d2 = _gh_read()
                    if d2 is not None:
                        if not any(e.get("id") == entry.get("id") for e in d2["events"]):
                            d2["events"].append(entry)
                        d2["events"] = d2["events"][-MAX_EVENTS:]
                        if _gh_write(d2):
                            d = d2
                d["events"] = d["events"][-MAX_EVENTS:]
                _local_write(d)
                _cache["data"] = d
                _cache["at"] = time.time()
        except Exception:
            pass

    if sync:
        work()
    else:
        threading.Thread(target=work, daemon=True).start()


def log(title, detail="", kind="other", action="create", page=None, user=None,
        username=None, role=None, dedupe_minutes=0, sync=False):
    """Record one event. Never raises."""
    try:
        if user is None:
            user, username, role = _current_user()
        key = f"{kind}|{action}|{username or user}|{title}"
        if dedupe_minutes:
            last = _recent.get(key, 0)
            if time.time() - last < dedupe_minutes * 60:
                return None
        _recent[key] = time.time()
        if len(_recent) > 2000:
            _recent.clear()
        entry = {
            "id": uuid.uuid4().hex[:12],
            "ts": _now(),
            "user": str(user or "غير معروف"),
            "username": str(username or ""),
            "role": str(role or ""),
            "kind": kind if kind in KIND_META else "other",
            "action": action,
            "title": str(title),
            "detail": str(detail or "")[:400],
            "page": page if page is not None else _page(),
        }
        _append(entry, sync=sync)
        return entry
    except Exception:
        return None


def change(title_prefix, label, kind, old, new, names=(), page=None, sync=False):
    """Log an add/update/delete by comparing the stored list before and after."""
    try:
        o = old if isinstance(old, list) else []
        n = new if isinstance(new, list) else []
        no, nn = len(o), len(n)
        if nn > no:
            action, verb = "create", "أضاف"
            added = n[len(o):] if no else n
            det = f"{no} ← {nn} سجل"
            extra = _describe(added, names)
            if extra:
                det += f" — {extra}"
        elif nn < no:
            action, verb = "delete", "حذف"
            det = f"{no} ← {nn} سجل"
            gone = []
            key = names[0] if names else None
            if key and no - nn <= 200:
                nset = {str(x.get(key, "")) for x in n if isinstance(x, dict)}
                for x in o:
                    if isinstance(x, dict) and str(x.get(key, "")) not in nset:
                        gone.append(x)
                        if len(gone) >= 3:
                            break
            elif nn == 0:
                gone = o[:3]
            extra = _describe(gone, names)
            if extra:
                det += f" — {extra}"
        else:
            action, verb = "update", "حدّث"
            det = f"{nn} سجل"
        return log(f"{verb} {label}", det, kind=kind, action=action, page=page, sync=sync)
    except Exception:
        return None


def unread(username):
    try:
        d = load()
        last = d.get("seen", {}).get(str(username or ""), "")
        return sum(1 for e in d["events"] if str(e.get("ts", "")) > last)
    except Exception:
        return 0


def mark_seen(username, data=None):
    try:
        if not username:
            return
        d = load(force=True)
        last = d.get("seen", {}).get(username, "")
        if not any(str(e.get("ts", "")) > last for e in d["events"]):
            _cache["at"] = time.time()
            return
        d["seen"][username] = _now()
        _local_write(d)
        _cache["data"] = d

        def work():
            try:
                with _lock:
                    fresh = _gh_read()
                    if fresh is None:
                        return
                    fresh["seen"][username] = _now()
                    _gh_write(fresh)
            except Exception:
                pass

        threading.Thread(target=work, daemon=True).start()
    except Exception:
        pass


def clear_cache():
    _cache["at"] = 0.0
