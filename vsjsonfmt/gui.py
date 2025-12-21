from __future__ import annotations

import os, sys, re, json, time, threading, tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from typing import Callable, Optional, List, Tuple
import tkinter.font as tkfont

# ------------- Johnny deps -------------
try:
    import customtkinter as ctk
except Exception:
    messagebox.showerror(
        "Missing dependency",
        "customtkinter is not installed.\n\nInstall:\n  pip install customtkinter pyperclip pillow",
    )
    raise SystemExit(1)

try:
    import pyperclip
except Exception:
    pyperclip = None  # we’ll fall back to tk’s clipboard or sumthin i hope

# optional: fast clipboard on Windows
_HAVE_WIN32 = False
try:
    import win32clipboard as wcb, win32con
    def _copy_win32(text: str, retries=4, delay=0.02):
        payload = text.replace("\n","\r\n")
        last=None
        for _ in range(retries):
            try:
                wcb.OpenClipboard(None)
                try:
                    wcb.EmptyClipboard()
                    wcb.SetClipboardData(win32con.CF_UNICODETEXT, payload)
                finally:
                    wcb.CloseClipboard()
                return
            except Exception as ex:
                last=ex; time.sleep(delay)
        if last: raise last
    _HAVE_WIN32=True
except Exception:
    pass

# ------------- App Constants -------------
APP_TITLE = "Prettiest — Vintage Story JSON Formatter"
DARK_BG   = "#0f0b1a"
CARD_BG   = "#1a1230"
ACCENT_2  = "#a78bfa"
TEXT_MUT  = "#cbd5e1"
ERR_LINE_BG = "#2f2969"

ROOT_DIR  = Path(__file__).resolve().parents[2]  # project root
ICO_PATH  = ROOT_DIR / "assets" / "PrettiestIcon.ico"
PNG_PATH  = ROOT_DIR / "assets" / "PrettiestIcon.png"
# ------------- User Preferences / Themes -------------
PREF_PATH = Path.home() / ".prettiest_prefs.json"

# Button look (requested: square + purple)
BTN_PURPLE = "#6d28d9"
BTN_PURPLE_HOVER = "#7c3aed"

THEMES = {
    "Night Owl": {
        "dark_bg": "#011627",
        "card_bg": "#0b253a",
        "accent":  "#c792ea",
        "text_mut":"#9aa7bd",
        "err_bg":  "#233a57",
        "tokens": {
            "key_imp":      {"foreground": "#ecc48d"},
            "key_imp_child":{"foreground": "#ecc48d"},
            "key_block":    {"foreground": "#ff7ab2"},
            "key_reg":      {"foreground": "#c792ea"},
            "str":          {"foreground": "#ecc48d"},
            "num":          {"foreground": "#addb67"},
            "bool":         {"foreground": "#82aaff"},
            "null":         {"foreground": "#82aaff"},
            "brace":        {"foreground": "#d6deeb"},
            "comm":         {"foreground": "#637777"},
            "colon":        {"foreground": "#d6deeb"},
        },
    },
    "One Dark": {
        "dark_bg": "#1e2127",
        "card_bg": "#23272e",
        "accent":  "#c678dd",
        "text_mut":"#abb2bf",
        "err_bg":  "#3a3f4b",
        "tokens": {
            "key_imp":      {"foreground": "#e5c07b"},
            "key_imp_child":{"foreground": "#e5c07b"},
            "key_block":    {"foreground": "#ff7ab2"},
            "key_reg":      {"foreground": "#c678dd"},
            "str":          {"foreground": "#98c379"},
            "num":          {"foreground": "#61afef"},
            "bool":         {"foreground": "#56b6c2"},
            "null":         {"foreground": "#56b6c2"},
            "brace":        {"foreground": "#e6e6e6"},
            "comm":         {"foreground": "#7f848e"},
            "colon":        {"foreground": "#e6e6e6"},
        },
    },
    "Dracula": {
        "dark_bg": "#282a36",
        "card_bg": "#2f3142",
        "accent":  "#bd93f9",
        "text_mut":"#b0b7d6",
        "err_bg":  "#44475a",
        "tokens": {
            "key_imp":      {"foreground": "#ffb86c"},
            "key_imp_child":{"foreground": "#ffb86c"},
            "key_block":    {"foreground": "#ff7ab2"},
            "key_reg":      {"foreground": "#bd93f9"},
            "str":          {"foreground": "#f1fa8c"},
            "num":          {"foreground": "#8be9fd"},
            "bool":         {"foreground": "#ff79c6"},
            "null":         {"foreground": "#ff79c6"},
            "brace":        {"foreground": "#f8f8f2"},
            "comm":         {"foreground": "#6272a4"},
            "colon":        {"foreground": "#f8f8f2"},
        },
    },
    "Nord": {
        "dark_bg": "#2e3440",
        "card_bg": "#3b4252",
        "accent":  "#b48ead",
        "text_mut":"#d8dee9",
        "err_bg":  "#4c566a",
        "tokens": {
            "key_imp":      {"foreground": "#ebcb8b"},
            "key_imp_child":{"foreground": "#ebcb8b"},
            "key_block":    {"foreground": "#ff7ab2"},
            "key_reg":      {"foreground": "#b48ead"},
            "str":          {"foreground": "#a3be8c"},
            "num":          {"foreground": "#88c0d0"},
            "bool":         {"foreground": "#81a1c1"},
            "null":         {"foreground": "#81a1c1"},
            "brace":        {"foreground": "#eceff4"},
            "comm":         {"foreground": "#7f8ea3"},
            "colon":        {"foreground": "#eceff4"},
        },
    },
}

def _load_prefs() -> dict:
    try:
        if PREF_PATH.exists():
            return json.loads(PREF_PATH.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}

def _save_prefs(prefs: dict) -> None:
    try:
        PREF_PATH.write_text(json.dumps(prefs, indent=2), encoding="utf-8")
    except Exception:
        pass

def _set_window_icon(win: tk.Misc) -> None:
    # iconbitmap (ico) + iconphoto (png) covers most platforms/taskbars
    try:
        if ICO_PATH.exists():
            win.iconbitmap(str(ICO_PATH))
    except Exception:
        pass
    try:
        if PNG_PATH.exists():
            img = tk.PhotoImage(file=str(PNG_PATH))
            setattr(win, "_prettiest_iconphoto", img)  # keep alive
            win.iconphoto(True, img)
    except Exception:
        pass

def _apply_theme(theme: dict) -> None:
    global DARK_BG, CARD_BG, ACCENT_2, TEXT_MUT, ERR_LINE_BG, TOKENS
    DARK_BG = theme["dark_bg"]
    CARD_BG = theme["card_bg"]
    ACCENT_2 = theme["accent"]
    TEXT_MUT = theme["text_mut"]
    ERR_LINE_BG = theme["err_bg"]
    TOKENS = theme["tokens"]

# ------------- Robust Local Imports (Giggidy) -------------
def _robust_imports():
    try:
        from vsjsonfmt.api import format_text as _fmt
        from vsjsonfmt import cli as _cli
        return _fmt, _cli._strip_comments, _cli._convert_single_quoted_strings
    except Exception:
        here = Path(__file__).parent
        sys.path.insert(0, str(here))
        sys.path.insert(0, str(here.parent))
        from api import format_text as _fmt   # type: ignore
        from cli import _strip_comments as _sc  # type: ignore
        from cli import _convert_single_quoted_strings as _csq  # type: ignore
        return _fmt, _sc, _csq

format_text, _strip_comments, _convert_single_quoted_strings = _robust_imports()

def _normalize_eol(s: str) -> str:
    return s.replace("\r\n","\n").replace("\r","\n")

def _strip_bom(s: str) -> str:
    """Strip a leading UTF-8 BOM (U+FEFF) if present."""
    return s[1:] if s.startswith("\ufeff") else s

# =======================================================
#      Visible-Range Syntax Highlighter (Fast & Light Baby!!)
# =======================================================
TOKENS = {
    "key_imp":      {"foreground": "#ffd166"},
    "key_imp_child":{"foreground": "#f3c56b"},
    "key_block":    {"foreground": "#ff7ab2"},
    "key_reg":      {"foreground": "#a78bfa"},
    "str":          {"foreground": "#f0abfc"},
    "num":          {"foreground": "#bae6fd"},
    "bool":         {"foreground": "#c4b5fd"},
    "null":         {"foreground": "#c4b5fd"},
    "brace":        {"foreground": "#e5e7eb"},
    "comm":         {"foreground": "#94a3b8"},
    "colon":        {"foreground": "#e5e7eb"},
}

IMPORTANT_KEYS = {
    "attributes","attribute","behaviors","behaviours","textures","texturesbytype",
    "shape","shapes","shapebytype","shapesbytype","elements","elementgroups",
    "variantgroups","variants","recipes","ingredients","outputs","drops",
    "entityproperties","creativeinventory"
}
IMPORTANT_SUBSTRINGS = ("texture","shape","behav","attrib")

# Keys that should use the special "block header" color (key_block) when their
# value opens an object/array. This is a small whitelist for Vintage Story
# model JSON.
BLOCK_HEADER_KEYS = {
    "children",
    "faces",
}

# Visible-range highlighter: fast on big files; bolds VS-relevant keys.
class AsyncHighlighter:
    STR_RE   = re.compile(r'(["\'])(?:\\.|(?!\1)[^\\\n])*\1')
    # Capture the *key string* separately so we can tag exactly "'key'" / "\"key\""
    # (including the closing quote) without any trailing whitespace/colon.
    KEY_RE   = re.compile(r'(?P<key>(?P<q>["\'])(?:\\.|(?!\2)[^\\\n])*\2)\s*:(?!:)')

    def __init__(self, root: tk.Misc, perf_threshold=120_000, margin_lines=80):
        self.root = root
        self.perf_threshold = perf_threshold
        self.margin_lines = margin_lines
        self._lock = threading.Lock()
        self._state = {}
        self._enabled = True
        self.bold_font = tkfont.Font(family="Consolas", size=12, weight="bold")

    def set_enabled(self, on: bool):
        self._enabled = on

    def _tw(self, w: tk.Text) -> tk.Text:
        return getattr(w, "_textbox", w)

    def _state_for(self, widget: tk.Text):
        k = id(widget)
        with self._lock:
            st = self._state.get(k)
            if not st:
                st = {"pending": None, "cancel": threading.Event(), "version": 0, "bound": False}
                self._state[k] = st
            return st

    def watch(self, widget: tk.Text, side: str):
        st = self._state_for(widget)
        if st["bound"]: return
        tw = self._tw(widget)
        def _bump(_=None): self.debounce(widget, side, 25, 80)
        for ev in ("<MouseWheel>","<ButtonRelease-1>","<Configure>"): tw.bind(ev, _bump, add="+")
        for ev in ("<KeyRelease-Up>","<KeyRelease-Down>","<KeyRelease-Prior>","<KeyRelease-Next>",
                   "<KeyRelease-Home>","<KeyRelease-End>"): tw.bind(ev, _bump, add="+")
        st["bound"] = True

    def debounce(self, widget: tk.Text, side: str, delay_ms_small=180, delay_ms_big=580):
        if not self._enabled: return
        tw = self._tw(widget)
        try: n = int(tw.count("1.0","end-1c","chars")[0])
        except Exception: n = len(tw.get("1.0","end-1c"))
        delay = delay_ms_big if n >= self.perf_threshold else delay_ms_small

        st = self._state_for(widget)
        st["version"] += 1
        st["cancel"].set()
        st["cancel"] = threading.Event()
        if st["pending"]:
            try: self.root.after_cancel(st["pending"])
            except Exception: pass
            st["pending"] = None
        st["pending"] = self.root.after(delay, lambda: self.highlight(widget, side))

    def _visible_range(self, tw: tk.Text) -> tuple[str,str]:
        try:
            start = tw.index("@0,0")
            end   = tw.index(f"@{tw.winfo_width()-1},{tw.winfo_height()-1}")
        except Exception:
            return "1.0","end-1c"
        start = tw.index(f"{start} linestart - {self.margin_lines} lines")
        end   = tw.index(f"{end} lineend + {self.margin_lines} lines")
        return start, end

    def _apply_styles(self, tw: tk.Text):
        for tag, style in TOKENS.items():
            kw = dict(style)
            if tag in ("key_imp","key_imp_child","key_block","key_reg"):
                kw["font"] = self.bold_font
            try: tw.tag_configure(tag, **kw)
            except Exception: pass
        try:
            tw.tag_raise("key_imp"); tw.tag_raise("key_imp_child"); tw.tag_raise("key_block"); tw.tag_raise("key_reg")
            tw.tag_lower("str")
        except Exception: pass

    def _is_important_key(self, key_name: str) -> bool:
        k = key_name.lower()
        return (k in IMPORTANT_KEYS) or any(sub in k for sub in IMPORTANT_SUBSTRINGS)

    def _tokenize(self, s: str, want_full: bool) -> List[Tuple[str,int,int]]:
        spans : List[Tuple[str,int,int]] = []
        # keys
        for m in self.KEY_RE.finditer(s):
            # Only tag the key string itself (including its closing quote).
            key_start, key_end = m.span("key")
            key_quoted = m.group("key")
            key_text = key_quoted[1:-1] if len(key_quoted) >= 2 else ""

            # Where does the value start?
            colon_pos = m.end("key")
            k = colon_pos
            while k < len(s) and s[k].isspace():
                k += 1

            # Block-header keys: a tiny whitelist (e.g., children/faces) gets a
            # dedicated color when the value opens an object/array.
            is_block = (k < len(s) and s[k] in '{[' and key_text.strip().lower() in BLOCK_HEADER_KEYS)

            if is_block:
                tag = "key_block"
            else:
                tag = "key_imp" if self._is_important_key(key_text) else "key_reg"

            spans.append((tag, key_start, key_end))
        # strings
        for m in self.STR_RE.finditer(s):
            spans.append(("str", m.start(), m.end()))
        # comments (very light)
        for m in re.finditer(r"//.*?$", s, flags=re.M):
            spans.append(("comm", m.start(), m.end()))
        for m in re.finditer(r"/\*.*?\*/", s, flags=re.S):
            spans.append(("comm", m.start(), m.end()))
        if want_full:
            for m in re.finditer(r"[{}\[\]:]", s):
                ch = m.group(0)
                spans.append(("brace" if ch in "{}[]" else "colon", m.start(), m.end()))
            for m in re.finditer(r"(?<![A-Za-z0-9_])(?:-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)", s):
                spans.append(("num", m.start(), m.end()))
            for m in re.finditer(r"(?<![A-Za-z0-9_])(true|false)(?![A-Za-z0-9_])", s):
                spans.append(("bool", m.start(), m.end()))
            for m in re.finditer(r"(?<![A-Za-z0-9_])null(?![A-Za-z0-9_])", s):
                spans.append(("null", m.start(), m.end()))
        spans.sort(key=lambda t: (t[0], t[1], t[2]))
        merged=[]
        for tag,a,b in spans:
            if merged and merged[-1][0]==tag and a <= merged[-1][2]:
                p=merged[-1]; merged[-1]=(tag,p[1],max(p[2],b))
            else:
                merged.append((tag,a,b))
        return merged

    def highlight(self, widget: tk.Text, side: str):
        if not self._enabled: return
        tw = self._tw(widget)
        st = self._state_for(widget)
        st["version"] += 1; version = st["version"]
        st["cancel"].set(); st["cancel"] = threading.Event()
        cancel_evt = st["cancel"]
        self._apply_styles(tw)
        start_idx, end_idx = self._visible_range(tw)
        try: snippet = tw.get(start_idx, end_idx)
        except Exception: return
        try:
            total_len = int(tw.count("1.0","end-1c","chars")[0])
        except Exception:
            total_len = len(tw.get("1.0","end-1c"))
        want_full = total_len < (self.perf_threshold // 2)

        def worker():
            try:
                spans = self._tokenize(snippet, want_full)
                if cancel_evt.is_set() or self._state_for(widget)["version"] != version:
                    return
                def apply_once():
                    try:
                        for tag in ("key_imp","key_imp_child","key_block","key_reg","str","num","bool","null","brace","comm","colon"):
                            tw.tag_remove(tag, start_idx, end_idx)
                        for tag,a,b in spans:
                            tw.tag_add(tag, f"{start_idx} + {a}c", f"{start_idx} + {b}c")
                    except Exception: pass
                self.root.after(0, apply_once)
            except Exception:
                pass
        threading.Thread(target=worker, daemon=True).start()

# =======================================================
#                        Tooltips
# =======================================================
# Tiny tooltip helper; keeps the main UI tidy.
# =======================================================
#                      Line Numbers
# =======================================================
class LineNumberBar(tk.Canvas):
    def __init__(self, parent: tk.Misc, textbox: tk.Text, *, bg: str, fg: str, font=("Consolas", 12), pad_x=6):
        super().__init__(parent, highlightthickness=0, bd=0, background=bg)
        self.textbox = textbox
        self._fg = fg
        self._font = font
        self._pad_x = pad_x
        self._job = None

        # redraw hooks
        for ev in ("<KeyRelease>", "<MouseWheel>", "<ButtonRelease-1>", "<Configure>"):
            self.textbox.bind(ev, self._schedule_redraw, add="+")
        self.textbox.bind("<<Change>>", self._schedule_redraw, add="+")
        self.textbox.bind("<Expose>", self._schedule_redraw, add="+")
        self._last_first = None
        self._last_last = None
        self.after(50, self.redraw)
        self.after(120, self._poll)

    def _poll(self):
        # Periodically check for scroll/size changes that don't emit events reliably on all platforms.
        try:
            first = self.textbox.index("@0,0")
            last = self.textbox.index(f"@0,{self.textbox.winfo_height()}")
        except Exception:
            self.after(120, self._poll)
            return

        if first != self._last_first or last != self._last_last:
            self._last_first = first
            self._last_last = last
            self.redraw()

        self.after(120, self._poll)

    def set_style(self, *, bg: str, fg: str, font):
        self.configure(background=bg)
        self._fg = fg
        self._font = font
        self.redraw()

    def _schedule_redraw(self, _evt=None):
        if self._job is not None:
            try:
                self.after_cancel(self._job)
            except Exception:
                pass
        self._job = self.after(25, self.redraw)

    def redraw(self, _evt=None):
        self._job = None
        self.delete("all")

        # IMPORTANT:
        # We query dlineinfo() from the *internal* tkinter.Text widget that lives
        # inside a CTkTextbox. This LineNumberBar canvas is a sibling of the
        # CTkTextbox container, not of the internal Text widget. That means the
        # coordinate systems can be vertically offset (CTk adds its own padding/
        # border/corner canvas). To make line numbers align with the visible
        # text lines, translate Text-widget coordinates into our canvas coords.
        try:
            y_offset = self.textbox.winfo_rooty() - self.winfo_rooty()
        except Exception:
            y_offset = 0
        try:
            i = self.textbox.index("@0,0")
            last = self.textbox.index(f"@0,{self.textbox.winfo_height()}")
        except Exception:
            return

        first_line = int(i.split(".")[0])
        last_line = int(last.split(".")[0]) + 1
        width = max(36, len(str(last_line)) * 9 + self._pad_x * 2)
        try:
            self.configure(width=width)
        except Exception:
            pass

        for line in range(first_line, last_line + 1):
            y = self.textbox.dlineinfo(f"{line}.0")
            if not y:
                continue
            y0 = y[1] + y_offset
            self.create_text(width - self._pad_x, y0, anchor="ne", text=str(line), fill=self._fg, font=self._font)


class HoverTip:
    def __init__(self, parent: tk.Misc, text: str):
        self.parent = parent; self.text=text; self.top=None
        parent.bind("<Enter>", self._show, add="+")
        parent.bind("<Leave>", self._hide, add="+")
        parent.bind("<ButtonPress>", self._hide, add="+")
    def _show(self, e=None):
        if self.top: return
        self.top = ctk.CTkToplevel(self.parent)
        self.top.overrideredirect(True)
        self.top.configure(fg_color=CARD_BG)
        self.top.attributes("-topmost", True)
        lbl = ctk.CTkLabel(self.top, text=self.text, text_color="#dbeafe",
                           wraplength=380, justify="left")
        lbl.pack(ipadx=10, ipady=6)
        x = self.parent.winfo_rootx()
        y = self.parent.winfo_rooty() + self.parent.winfo_height() + 6
        self.top.geometry(f"+{x}+{y}")
    def _hide(self, e=None):
        if self.top:
            try: self.top.destroy()
            except Exception: pass
            self.top=None

# =======================================================
#            VS Warnings / Errors Helpers
# =======================================================
ERR_TAG = "errline"

def highlight_error_line(textbox: tk.Text, lineno: int) -> None:
    tw = getattr(textbox, "_textbox", textbox)
    start=f"{lineno}.0"; end=f"{lineno+1}.0"
    try:
        tw.tag_remove(ERR_TAG, "1.0", "end")
        tw.tag_configure(ERR_TAG, background=ERR_LINE_BG)
        tw.tag_add(ERR_TAG, start, end); tw.see(start)
    except Exception:
        pass

# Accidental return inside quoted string. (Preserves comments.)
def _detect_accidental_return(src: str) -> tuple[bool,int,int]:
    quote: str|None = None  # '"' or "'"
    esc=False; in_line=False; in_block=False
    i=0; n=len(src); lineno=1; col=0
    while i<n:
        ch=src[i]; col+=1
        if ch=="\n":
            if quote is not None: return True, lineno, col
            in_line=False; lineno+=1; col=0; i+=1; continue
        if ch=="\r":
            if quote is not None: return True, lineno, col
            i+=1; continue
        if in_block:
            if ch=="*" and i+1<n and src[i+1]=="/": in_block=False; i+=2; continue
            i+=1; continue
        if in_line:
            i+=1; continue
        if quote is None and ch=="/" and i+1<n:
            if src[i+1]=="/": in_line=True; i+=2; continue
            if src[i+1]=="*": in_block=True; i+=2; continue
        if quote is not None:
            if esc: esc=False
            else:
                if ch=="\\": esc=True
                elif ch==quote: quote=None
        else:
            if ch in ('"', "'"):
                quote = ch
                esc = False
        i+=1
    return False,0,0

_IDENT = re.compile(r'(?P<pre>[\s{,])(?P<key>[A-Za-z_][A-Za-z0-9_\-]*)\s*:')

def _quote_unquoted_keys_preserving_lines(s: str) -> str:
    # add quotes around bare keys, preserving newlines
    out=[]; i=0; n=len(s); quote: str|None=None; esc=False; in_line=False; in_block=False
    while i<n:
        ch=s[i]
        if quote is not None:
            out.append(ch)
            if esc: esc=False
            elif ch=="\\": esc=True
            elif ch==quote: quote=None
            i+=1; continue
        if in_line:
            out.append(ch)
            if ch=="\n": in_line=False
            i+=1; continue
        if in_block:
            out.append(ch)
            if ch=="*" and i+1<n and s[i+1]=="/": out.append("/"); i+=2; in_block=False; continue
            i+=1; continue
        if ch=='/':
            if i+1<n and s[i+1]=='/': out.append("//"); i+=2; in_line=True; continue
            if i+1<n and s[i+1]=='*': out.append("/*"); i+=2; in_block=True; continue
        if ch in ('"', "'"):
            quote = ch
            out.append(ch)
            i+=1
            continue
        if ch in " \t\r\n{,":
            m=_IDENT.match(s, i)
            if m:
                out.append(m.group("pre"))
                out.append('"'); out.append(m.group("key")); out.append('":')
                i = m.end()
                continue
        out.append(ch); i+=1
    return "".join(out)

def _remove_trailing_commas(s: str) -> str:
    return re.sub(r",(\s*[}\]])", r"\1", s)

def _vs_warnings(src: str) -> list[str]:
    warns=[]
    cleaned = _strip_comments(src)
    if re.search(r'(?m)^[ \t]*[A-Za-z_][A-Za-z0-9_\-]*\s*:', cleaned):
        warns.append("Unquoted property names detected. Allowed by Vintage Story; quoting is recommended for clarity.")
    if re.search(r",\s*[}\]]", cleaned):
        warns.append("Trailing commas detected before '}' or ']'. Allowed by Vintage Story.")
    # multiple top-level values (tolerated by VS)
    top = cleaned.strip()
    if top:
        count = 0; depth = 0; in_str=False; esc=False
        for ch in top:
            if in_str:
                if esc: esc=False
                elif ch=="\\": esc=True
                elif ch=='"': in_str=False
                continue
            if ch=='"': in_str=True; continue
            if ch in "{[": depth += 1
            elif ch in "}]": depth -= 1
            elif ch == "\n" and depth==0: count += 1
        if count>0:
            warns.append("Multiple top-level JSON values detected (Vintage Story allows this).")
    return warns

# --- “missing comma” auto-fix (best-effort, line-preserving) ---
_VALUE_END = re.compile(r'[}\]0-9"truefalsenull\s\]]', re.I)
def _autofix_insert_comma_at(src: str, lineno: int, colno: int) -> tuple[str,int]:
    """
    Insert a comma after the end of the previous JSON value on or above the given line.
    Returns (new_src, fixed_line).
    """
    lines = src.splitlines(True)  # keepends
    ln = max(1, min(lineno, len(lines)))
    # work on current line first
    def last_nonspace_idx(s: str) -> int:
        i = len(s)-1
        while i>=0 and s[i].isspace(): i-=1
        return i
    i = last_nonspace_idx(lines[ln-1][:colno-1] if colno>1 else lines[ln-1])
    if i >= 0 and lines[ln-1][i] != ",":
        lines[ln-1] = lines[ln-1][:i+1] + "," + lines[ln-1][i+1:]
        return ("".join(lines), ln)
    # otherwise use previous line
    p = ln-1
    while p>0:
        j = last_nonspace_idx(lines[p-1])
        if j >= 0:
            if lines[p-1][j] != ",":
                lines[p-1] = lines[p-1][:j+1] + "," + lines[p-1][j+1:]
            return ("".join(lines), p)
        p -= 1
    return (src, ln)

# =======================================================
#                     Progress Modal
# =======================================================
# Progress bar + cancel; used for long ops like huge pastes/opens.
class ProgressModal:
    def __init__(self, root: tk.Misc, title="Working…"):
        self.root = root
        self.top = ctk.CTkToplevel(root)
        self.top.overrideredirect(True)
        self.top.attributes("-topmost", True)
        _set_window_icon(self.top)
        self.top.configure(fg_color=CARD_BG)
        self.top.withdraw()
        root.update_idletasks()
        rx, ry = root.winfo_rootx(), root.winfo_rooty()
        rw, rh = root.winfo_width(), root.winfo_height()
        w, h = 380, 160
        x = rx + (rw - w)//2; y = ry + (rh - h)//2
        self.top.geometry(f"{w}x{h}+{max(0,x)}+{max(0,y)}")

        card = ctk.CTkFrame(self.top, fg_color=CARD_BG, corner_radius=18)
        card.pack(fill="both", expand=True, padx=10, pady=10)
        self._label_var = tk.StringVar(value=title)
        ctk.CTkLabel(card, textvariable=self._label_var).pack(pady=(12,6))
        self._pb = ctk.CTkProgressBar(card, mode="determinate", width=300)
        self._pb.set(0.0); self._pb.pack(pady=(4,10))
        self.cancelled = threading.Event()
        ctk.CTkButton(card, text="Cancel", command=self.cancel, width=120, corner_radius=0,
                      fg_color=BTN_PURPLE, hover_color=BTN_PURPLE_HOVER).pack(pady=(0,12))
        self.top.deiconify(); self.top.grab_set(); self.top.focus_force()

    def set_text(self, t: str): self.root.after(0, lambda: self._label_var.set(t))
    def set_progress(self, p: float): self.root.after(0, lambda: self._pb.set(max(0.0, min(1.0, p))))
    def cancel(self): self.cancelled.set(); self.set_text("Cancelling…")
    def close(self):
        def _do():
            try: self.top.grab_release()
            except Exception: pass
            try: self.top.destroy()
            except Exception: pass
        self.root.after(0, _do)

# =======================================================
#                     Quote Toggles
# =======================================================
KW = {"true","false","null"}
IDENT_BODY = re.compile(r"^[A-Za-z0-9_\-]*$")
def _scan_string(s: str, i: int, quote: str = '"') -> int:
    i += 1; esc=False; n=len(s)
    while i<n:
        c=s[i]
        if esc: esc=False
        elif c=="\\": esc=True
        elif c==quote: return i+1
        i+=1
    return n
def _scan_ident(s: str, i: int) -> int:
    if i>=len(s): return i
    c=s[i]
    if not (c.isalpha() or c=="_"): return i
    j=i+1
    while j<len(s) and (s[j].isalnum() or s[j] in "_-"): j+=1
    return j
# Quote <-> unquote engine with progress and cancel.
def _toggle_quotes_progress(src: str, to_quoted: bool,
                            progress: Callable[[float],None]|None=None,
                            cancelled: threading.Event|None=None) -> str:
    s=src; n=len(s); out=[]; last=0; i=0
    stride=max(65536, n//200); next_mark=stride
    def flush(a,b):
        if b>a: out.append(s[a:b])
    while i<n:
        if cancelled is not None and cancelled.is_set(): return src
        if progress and i>=next_mark:
            progress(min(0.999, i/max(1,n))); next_mark += stride
        ch = s[i]
        if ch in ('"', "'"):
            endq=_scan_string(s,i,ch); j=endq
            while j<n and s[j].isspace(): j+=1
            if j<n and s[j]==":":
                content=s[i+1:endq-1]
                if content and (content[0].isalpha() or content[0]=='_') and IDENT_BODY.match(content[1:] or ""):
                    if not to_quoted:
                        # Unquote "key": or 'key': -> key:
                        flush(last,i); out.append(content); out.append(s[endq:j+1]); last=j+1; i=j+1; continue
                    else:
                        # Convert single-quoted keys 'key': -> "key": (leave double-quoted keys untouched)
                        if ch=="'":
                            flush(last,i); out.append('"'); out.append(content); out.append('"'); out.append(s[endq:j+1]); last=j+1; i=j+1; continue
                i=j+1; continue
            i=endq; continue
        if ch.isalpha() or ch=="_":
            j=_scan_ident(s,i); k=j
            while k<n and s[k].isspace(): k+=1
            if k<n and s[k]==":":
                if to_quoted:
                    flush(last,i); out.append('"'); out.append(s[i:j]); out.append('"'); out.append(s[j:k+1]); last=k+1; i=k+1; continue
                i=k+1; continue
            i=j; continue
        i+=1
    flush(last,n)
    if progress: progress(1.0)
    return "".join(out)
def quote_keys_and_values(s: str, progress=None, cancelled=None) -> str:
    return _toggle_quotes_progress(s, True, progress, cancelled)
def unquote_keys_and_values(s: str, progress=None, cancelled=None) -> str:
    return _toggle_quotes_progress(s, False, progress, cancelled)

# =======================================================
#                           GUI
# =======================================================
def _center(win, w=1320, h=760):
    win.update_idletasks()
    sw,sh = win.winfo_screenwidth(), win.winfo_screenheight()
    x,y = (sw-w)//2, (sh-h)//2
    win.geometry(f"{w}x{h}+{max(0,x)}+{max(0,y)}")

# App entry: layout, bindings, actions, shortcuts. Ctrl+Enter formats.
def run_gui():
    ctk.set_default_color_theme("dark-blue")
    ctk.set_appearance_mode("dark")

    prefs = _load_prefs()
    theme_name = prefs.get("theme", "Night Owl")
    if theme_name not in THEMES:
        theme_name = "Night Owl"
    _apply_theme(THEMES[theme_name])

    root = ctk.CTk()
    root.title(APP_TITLE)
    root.configure(fg_color=DARK_BG)
    _center(root, w=1320, h=780)
    _set_window_icon(root)

    # Top: controls + buttons
    top = ctk.CTkFrame(root, corner_radius=16, fg_color=CARD_BG)
    top.pack(fill="x", padx=16, pady=(16, 8))
    top.grid_columnconfigure(0, weight=1)

    controls = ctk.CTkFrame(top, fg_color="transparent")
    controls.grid(row=0, column=0, sticky="w", padx=10, pady=(10, 2))

    # persistent vars
    safety_backup_var = tk.BooleanVar(value=bool(prefs.get("safety_backup", True)))
    editor_scale_var = tk.DoubleVar(value=float(prefs.get("editor_scale", 1.0)))

    # formatting controls
    lbl_compress = ctk.CTkLabel(controls, text="Compress")
    lbl_compress.grid(row=0, column=0, padx=(0, 6))
    HoverTip(lbl_compress, "Compress: Expand or compress your formatting based on number of lines in a section.")
    width_entry = ctk.CTkEntry(controls, width=72)
    width_entry.insert(0, str(prefs.get("print_width", "120")))
    width_entry.grid(row=0, column=1, padx=(0, 16))

    lbl_header = ctk.CTkLabel(controls, text="Header After")
    lbl_header.grid(row=0, column=2, padx=(0, 6))
    HoverTip(lbl_header, "Header after # of lines: Formats a 'title header' that doesn't collapse in an IDE for visual identifying and organization.")
    long_entry = ctk.CTkEntry(controls, width=72)
    long_entry.insert(0, str(prefs.get("long_block_threshold", "20")))
    long_entry.grid(row=0, column=3, padx=(0, 16))

    # editor text size (bigger + cleaner: +/- + percentage)
    scale_row = ctk.CTkFrame(controls, fg_color="transparent")
    scale_row.grid(row=1, column=0, columnspan=4, sticky="w", pady=(6, 2))

    ctk.CTkLabel(scale_row, text="Editor Text Size").pack(side="left", padx=(0, 8))

    scale_readout = tk.StringVar(value="100%")
    ctk.CTkLabel(scale_row, textvariable=scale_readout, text_color=TEXT_MUT).pack(side="left", padx=(0, 10))

    def _clamp_scale(v: float) -> float:
        return max(0.75, min(2.0, v))

    ui_scale = ctk.CTkSlider(scale_row, from_=0.75, to=2.0, number_of_steps=25, width=240)
    ui_scale.set(_clamp_scale(editor_scale_var.get()))
    ui_scale.pack(side="left")

    def _nudge_scale(delta: float):
        v = _clamp_scale(float(ui_scale.get()) + delta)
        ui_scale.set(v)
        on_scale(v)

    ctk.CTkButton(scale_row, text="−", width=36, height=30, corner_radius=0,
                  fg_color=BTN_PURPLE, hover_color=BTN_PURPLE_HOVER,
                  command=lambda: _nudge_scale(-0.05)).pack(side="left", padx=(10, 6))
    ctk.CTkButton(scale_row, text="+", width=36, height=30, corner_radius=0,
                  fg_color=BTN_PURPLE, hover_color=BTN_PURPLE_HOVER,
                  command=lambda: _nudge_scale(0.05)).pack(side="left")

    backup_row = ctk.CTkFrame(controls, fg_color="transparent")
    backup_row.grid(row=2, column=0, columnspan=4, sticky="w", pady=(2, 2))
    backup_cb = ctk.CTkCheckBox(
        backup_row,
        text="Create safety backup (batch)",
        variable=safety_backup_var,
        onvalue=True,
        offvalue=False,
        fg_color=BTN_PURPLE,
        hover_color=BTN_PURPLE_HOVER,
    )
    backup_cb.pack(side="left")

    # Buttons container
    btns = ctk.CTkFrame(top, fg_color="transparent")
    btns.grid(row=0, column=1, sticky="e", padx=10, pady=(10, 2))

    # Split panes
    split = ctk.CTkFrame(root, corner_radius=16, fg_color="transparent")
    split.pack(fill="both", expand=True, padx=16, pady=(8, 16))
    left = ctk.CTkFrame(split, corner_radius=16, fg_color=CARD_BG)
    right = ctk.CTkFrame(split, corner_radius=16, fg_color=CARD_BG)
    left.pack(side="left", fill="both", expand=True, padx=(0, 8))
    right.pack(side="left", fill="both", expand=True, padx=(8, 0))

    lbl_in = ctk.CTkLabel(left, text="Input (JSON / JSONC)", text_color=ACCENT_2)
    lbl_out = ctk.CTkLabel(right, text="Output (Formatted JSON)", text_color=ACCENT_2)
    lbl_in.pack(anchor="w", padx=12, pady=(12, 4))
    lbl_out.pack(anchor="w", padx=12, pady=(12, 4))

    # Editors: line numbers + textbox
    editor_font_base = max(10, int(12 * ui_scale.get()))
    editor_font = ("Consolas", editor_font_base)

    left_edit = ctk.CTkFrame(left, fg_color="transparent")
    right_edit = ctk.CTkFrame(right, fg_color="transparent")
    left_edit.pack(fill="both", expand=True, padx=12, pady=(0, 12))
    right_edit.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    inp = ctk.CTkTextbox(left_edit, corner_radius=10, wrap="none", font=editor_font)
    out = ctk.CTkTextbox(right_edit, corner_radius=10, wrap="none", font=editor_font)

    inp_tw = getattr(inp, "_textbox", inp)
    out_tw = getattr(out, "_textbox", out)

    gutter_bg = DARK_BG
    gutter_fg = TEXT_MUT

    ln_inp = LineNumberBar(left_edit, inp_tw, bg=gutter_bg, fg=gutter_fg, font=editor_font)
    ln_out = LineNumberBar(right_edit, out_tw, bg=gutter_bg, fg=gutter_fg, font=editor_font)

    ln_inp.pack(side="left", fill="y", padx=(0, 6))
    inp.pack(side="left", fill="both", expand=True)

    ln_out.pack(side="left", fill="y", padx=(0, 6))
    out.pack(side="left", fill="both", expand=True)

    # footer & credit
    status = ctk.CTkFrame(root, corner_radius=12, fg_color=CARD_BG)
    status.pack(fill="x", padx=16, pady=(0, 12))
    status_var = tk.StringVar(value="Ready")
    ctk.CTkLabel(status, textvariable=status_var, text_color=TEXT_MUT).pack(side="left", padx=10, pady=6)

    credit_row = ctk.CTkFrame(status, fg_color="transparent")
    credit_row.pack(side="right", padx=10, pady=6)

    badge_img = None
    try:
        from PIL import Image
        if PNG_PATH.exists():
            badge_img = ctk.CTkImage(Image.open(PNG_PATH), size=(16, 16))
    except Exception:
        badge_img = None

    ctk.CTkLabel(credit_row, image=badge_img, text="  By Pure Winter", compound="left").pack(side="left")

    # theme picker (bottom right)
    ctk.CTkLabel(credit_row, text="  🎨", text_color=TEXT_MUT).pack(side="left", padx=(12, 4))
    theme_var = tk.StringVar(value=theme_name)
    theme_menu = ctk.CTkOptionMenu(
        credit_row,
        values=list(THEMES.keys()),
        variable=theme_var,
        width=140,
        corner_radius=0,
        fg_color=BTN_PURPLE,
        button_color=BTN_PURPLE,
        button_hover_color=BTN_PURPLE_HOVER,
        dropdown_fg_color=CARD_BG,
    )
    theme_menu.pack(side="left")

    # highlighter
    PERF_THRESHOLD = 120_000
    highlighter = AsyncHighlighter(root, PERF_THRESHOLD)
    highlighter.watch(inp, "left")
    highlighter.watch(out, "right")

    # persist helpers
    def _save_pref_now():
        prefs["theme"] = theme_var.get()
        prefs["safety_backup"] = bool(safety_backup_var.get())
        prefs["editor_scale"] = float(ui_scale.get())
        prefs["print_width"] = width_entry.get()
        prefs["long_block_threshold"] = long_entry.get()
        _save_prefs(prefs)

    def _apply_theme_to_widgets():
        root.configure(fg_color=DARK_BG)
        for f in (top, left, right, status):
            f.configure(fg_color=CARD_BG)
        lbl_in.configure(text_color=ACCENT_2)
        lbl_out.configure(text_color=ACCENT_2)
        theme_menu.configure(dropdown_fg_color=CARD_BG)
        try:
            batch_menu.configure(dropdown_fg_color=CARD_BG)
        except Exception:
            pass
        ln_inp.set_style(bg=DARK_BG, fg=TEXT_MUT, font=editor_font)
        ln_out.set_style(bg=DARK_BG, fg=TEXT_MUT, font=editor_font)
        try:
            inp.configure(fg_color=CARD_BG, text_color="#e5e7eb")
            out.configure(fg_color=CARD_BG, text_color="#e5e7eb")
        except Exception:
            pass
        highlighter.debounce(inp, "left")
        highlighter.debounce(out, "right")

    def on_theme_change(_val=None):
        name = theme_var.get()
        if name not in THEMES:
            return
        _apply_theme(THEMES[name])
        _apply_theme_to_widgets()
        _save_pref_now()

    theme_menu.configure(command=on_theme_change)

    # scaled fonts (debounced & no flicker) — debounce + pause highlighter
    _scale_job = {"id": None}

    def set_ui_scale_now(scale: float):
        nonlocal editor_font_base, editor_font
        scale = _clamp_scale(scale)
        editor_scale_var.set(scale)
        editor_font_base = max(10, int(12 * scale))
        editor_font = ("Consolas", editor_font_base)
        scale_readout.set(f"{int(scale * 100)}%")
        inp.configure(font=editor_font)
        out.configure(font=editor_font)
        ln_inp.set_style(bg=DARK_BG, fg=TEXT_MUT, font=editor_font)
        ln_out.set_style(bg=DARK_BG, fg=TEXT_MUT, font=editor_font)
        try:
            highlighter.bold_font.configure(size=editor_font_base, weight="bold")
        except Exception:
            highlighter.bold_font = tkfont.Font(family="Consolas", size=editor_font_base, weight="bold")
        highlighter.set_enabled(True)
        highlighter.debounce(inp, "left")
        highlighter.debounce(out, "right")
        _save_pref_now()

    def on_scale(val):
        highlighter.set_enabled(False)
        if _scale_job["id"] is not None:
            try:
                root.after_cancel(_scale_job["id"])
            except Exception:
                pass
        _scale_job["id"] = root.after(120, lambda v=float(val): set_ui_scale_now(v))

    ui_scale.configure(command=on_scale)
    set_ui_scale_now(float(ui_scale.get()))

    def _numbers() -> tuple[int, int]:
        try:
            w = int(width_entry.get() or "120")
        except Exception:
            w = 120
        try:
            l = int(long_entry.get() or "20")
        except Exception:
            l = 20
        return w, l

    def _set_text(widget: tk.Text, text: str):
        tw = getattr(widget, "_textbox", widget)
        tw.delete("1.0", "end")
        tw.insert("1.0", text)

    # incremental insert with progress
    def _insert_text_incremental(widget: tk.Text, text: str, modal: ProgressModal,
                                 start=0.1, end=1.0, side="left", done_cb: Callable | None = None):
        tw = getattr(widget, "_textbox", widget)
        text = _normalize_eol(text)
        total = len(text)
        if total == 0:
            modal.set_progress(end)
            done_cb and done_cb()
            return
        tw.delete("1.0", "end")
        chunk = max(250_000, total // 120)
        idx = 0

        def step():
            nonlocal idx
            if modal.cancelled.is_set():
                modal.close()
                done_cb and done_cb()
                return
            j = min(idx + chunk, total)
            tw.insert("end", text[idx:j])
            idx = j
            modal.set_progress(start + (end - start) * (idx / total))
            if idx < total:
                tw.update_idletasks()
                root.after(1, step)
            else:
                highlighter.debounce(widget, side)
                done_cb and done_cb()

        root.after(0, step)

    # dialogs
    def _show_text_dialog(title: str, text: str, *,
                          action_title: str | None = None,
                          action: Callable[[], None] | None = None):
        dlg = ctk.CTkToplevel(root)
        dlg.title(title)
        dlg.attributes("-topmost", True)
        dlg.grab_set()
        _set_window_icon(dlg)

        w, h = 820, 420
        rx, ry = root.winfo_rootx(), root.winfo_rooty()
        rw, rh = root.winfo_width(), root.winfo_height()
        x = rx + (rw - w) // 2
        y = ry + (rh - h) // 2
        dlg.geometry(f"{w}x{h}+{max(0, x)}+{max(0, y)}")

        container = ctk.CTkFrame(dlg, fg_color=CARD_BG, corner_radius=16)
        container.pack(fill="both", expand=True, padx=12, pady=12)

        t = ctk.CTkTextbox(container, wrap="word", corner_radius=10)
        t.configure(font=("Consolas", max(12, int(14 * editor_scale_var.get()))))
        t.pack(fill="both", expand=True, padx=8, pady=(8, 8))
        t.insert("1.0", text.strip() + "\n")
        t.configure(state="disabled")

        row = ctk.CTkFrame(container, fg_color="transparent")
        row.pack(fill="x", padx=8, pady=(0, 8))
        if action_title and action:
            ctk.CTkButton(row, text=action_title, width=190, command=lambda: (dlg.destroy(), action()),
                          corner_radius=0, fg_color=BTN_PURPLE, hover_color=BTN_PURPLE_HOVER).pack(side="right", padx=8)
        ctk.CTkButton(row, text="OK", width=100, command=dlg.destroy,
                      corner_radius=0, fg_color=BTN_PURPLE, hover_color=BTN_PURPLE_HOVER).pack(side="right")

    # actions
    def do_format():
        w, l = _numbers()
        src_text = getattr(inp, "_textbox", inp).get("1.0", "end-1c")
        if not src_text.strip():
            out.delete("1.0", "end")
            status_var.set("Nothing to format")
            return
        modal = ProgressModal(root, title="Formatting…")
        modal.set_progress(0.02)

        def work():
            try:
                formatted = _normalize_eol(format_text(src_text, print_width=w, long_block_threshold=l, aggressive_inline=True))

                def after_insert():
                    highlighter.debounce(out, "right")
                    modal.close()
                    status_var.set("Formatted ✓")

                _insert_text_incremental(out, formatted, modal, start=0.1, end=1.0, side="right", done_cb=after_insert)
            except Exception as ex:
                root.after(0, lambda ex=ex: (modal.close(), messagebox.showerror("Format", str(ex))))

        threading.Thread(target=work, daemon=True).start()

    def _fast_copy(text: str) -> bool:
        if _HAVE_WIN32:
            try:
                _copy_win32(text)
                return True
            except Exception:
                pass
        if pyperclip:
            try:
                pyperclip.copy(text)
                return True
            except Exception:
                pass
        try:
            root.clipboard_clear()
            root.clipboard_append(text)
            return True
        except Exception:
            return False

    def copy_right():
        text = getattr(out, "_textbox", out).get("1.0", "end-1c")
        status_var.set("Copied ✓" if _fast_copy(text) else "Copy failed")

    def paste_left():
        modal = ProgressModal(root, title="Pasting…")
        modal.set_progress(0.05)

        def bg():
            try:
                if pyperclip:
                    text = _strip_bom(_normalize_eol(pyperclip.paste()))
                else:
                    text = _strip_bom(_normalize_eol(root.clipboard_get()))

                def after_insert():
                    highlighter.debounce(inp, "left")
                    modal.close()
                    status_var.set("Pasted ✓")

                _insert_text_incremental(inp, text, modal, start=0.1, end=1.0, side="left", done_cb=after_insert)
            except Exception as ex:
                root.after(0, lambda ex=ex: (modal.close(), messagebox.showerror("Paste", str(ex))))

        threading.Thread(target=bg, daemon=True).start()

    def open_left():
        p = filedialog.askopenfilename(filetypes=[("JSON / JSONC", "*.json *.jsonc"), ("All", "*.*")])
        if not p:
            return
        try:
            # utf-8-sig removes an optional BOM; we also strip defensively.
            data = Path(p).read_text(encoding="utf-8-sig")
            _set_text(inp, _strip_bom(_normalize_eol(data)))
            status_var.set(f"Opened: {Path(p).name}")
            highlighter.debounce(inp, "left")
        except Exception as ex:
            messagebox.showerror("Open", str(ex))

    def save_right():
        p = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json"), ("All", "*.*")])
        if not p:
            return
        try:
            data = getattr(out, "_textbox", out).get("1.0", "end-1c")
            with open(p, "w", encoding="utf-8", newline="\n") as f:
                f.write(data if data.endswith("\n") else data + "\n")
            status_var.set(f"Saved: {Path(p).name}")
        except Exception as ex:
            messagebox.showerror("Save", str(ex))

    # Batch helpers
    def _create_backup(path: Path) -> Path | None:
        if not safety_backup_var.get():
            return None
        # Prefer a deterministic name, but avoid clobbering a previous backup.
        candidate = path.with_suffix(path.suffix + ".bak")
        if candidate.exists():
            candidate = path.with_suffix(path.suffix + f".bak.{int(time.time())}")
        try:
            candidate.write_bytes(path.read_bytes())
            return candidate
        except Exception:
            return None

    def _atomic_write(path: Path, text: str) -> None:
        tmp = path.with_suffix(path.suffix + ".prettiest.tmp")
        with open(tmp, "w", encoding="utf-8", newline="\n") as f:
            f.write(text)
        os.replace(str(tmp), str(path))

    def _batch_process(kind: str):
        files = filedialog.askopenfilenames(
            title=f"Select files for {kind}",
            filetypes=[("JSON / JSONC", "*.json *.jsonc"), ("All", "*.*")],
        )
        if not files:
            return
        paths = [Path(p) for p in files]
        modal = ProgressModal(root, title=f"{kind}…")
        modal.set_progress(0.02)
        w, l = _numbers()

        def work():
            ok = 0
            failed: list[str] = []
            for i, fp in enumerate(paths):
                if modal.cancelled.is_set():
                    break
                try:
                    # utf-8-sig removes an optional BOM; we also strip defensively.
                    original = _strip_bom(fp.read_text(encoding="utf-8-sig"))
                    _create_backup(fp)

                    if kind == "Batch - Format":
                        new_text = _normalize_eol(format_text(original, print_width=w, long_block_threshold=l, aggressive_inline=True))
                    elif kind == "Batch - Quotes":
                        # Batch Quotes should always FORCE quotes ON (match the intent of the Quotes tool
                        # for vanilla-style JSON with lots of unquoted keys).
                        new_text = _normalize_eol(quote_keys_and_values(original, cancelled=modal.cancelled))
                    else:
                        raise RuntimeError("Unknown batch action")

                    if not new_text.endswith("\n"):
                        new_text += "\n"
                    _atomic_write(fp, new_text)
                    ok += 1
                except Exception as ex:
                    failed.append(f"{fp.name}: {ex}")
                modal.set_progress(0.02 + 0.98 * ((i + 1) / max(1, len(paths))))

            def finish():
                modal.close()
                # Keep UI state consistent: if Batch Quotes fully succeeds, Quotes are ON.
                if kind == "Batch - Quotes" and (not modal.cancelled.is_set()) and (not failed) and ok == len(paths):
                    quotes_state["on"] = True
                if failed:
                    status_var.set(f"{kind}: {ok}/{len(paths)} ok")
                    _show_text_dialog(kind, "Some files failed:\n\n" + "\n".join(failed[:120]))
                else:
                    status_var.set(f"{kind}: {ok} file(s) ✓")

            root.after(0, finish)

        threading.Thread(target=work, daemon=True).start()

    # Warnings / Errors
    def _show_warnings():
        src_text = getattr(inp, "_textbox", inp).get("1.0", "end-1c")
        if not src_text.strip():
            messagebox.showinfo("Vintage Story JSON check", "Left pane is empty.")
            return
        warns = _vs_warnings(src_text)
        body = "No warnings." if not warns else ("Warnings:\n• " + "\n• ".join(warns))
        _show_text_dialog("Vintage Story JSON check", body)

    def _strictish(src_text: str) -> str:
        # Vintage Story is fairly permissive (unquoted keys, trailing commas, single-quoted strings).
        # For a helpful “strict-ish” check we normalize those into valid JSON, preserving line layout
        # so the reported error line stays accurate.
        src_text = _strip_bom(src_text)
        return _convert_single_quoted_strings(
            _remove_trailing_commas(
                _quote_unquoted_keys_preserving_lines(
                    _strip_comments(src_text)
                )
            )
        )

    def _show_errors():
        src_text = getattr(inp, "_textbox", inp).get("1.0", "end-1c")
        if not src_text.strip():
            messagebox.showinfo("Vintage Story JSON check", "Left pane is empty.")
            return

        brk, ln, col = _detect_accidental_return(src_text)
        if brk:
            line = src_text.splitlines()[ln - 1] if 1 <= ln <= len(src_text.splitlines()) else ""
            pointer = " " * (max(0, col - 1)) + "^"
            msg = (
                f"Invalid control character at\n(Line {ln})\n\n{line}\n{pointer}\n\n"
                "Hint: Detected a line break inside a quoted string (accidental Return)."
            )
            highlight_error_line(inp, ln)
            _show_text_dialog("Vintage Story JSON check", msg)
            return

        strictish = _strictish(src_text)
        try:
            json.loads(strictish)
            _show_text_dialog("Vintage Story JSON check", "No errors found.")
        except json.JSONDecodeError as e:
            ln = max(1, e.lineno)
            col = max(1, e.colno)
            line = strictish.splitlines()[ln - 1] if 1 <= ln <= len(strictish.splitlines()) else ""
            pointer = " " * (col - 1) + "^"
            msg = (
                f"{e.msg}\n\n(Line {ln}, column {col})\n{line}\n{pointer}\n\n"
                "Tip: This is likely a missing comma between items or a broken structure."
            )
            highlight_error_line(inp, ln)

            if "Expecting ',' delimiter" in e.msg:
                def do_fix():
                    new_src, fixed_ln = _autofix_insert_comma_at(src_text, ln, col)
                    _set_text(inp, new_src)
                    highlight_error_line(inp, fixed_ln)
                    status_var.set("Inserted missing comma")
                    highlighter.debounce(inp, "left")

                _show_text_dialog("Vintage Story JSON check", msg, action_title="Auto-fix: Insert comma", action=do_fix)
            else:
                _show_text_dialog("Vintage Story JSON check", msg)

    # quotes toggle (single button)
    quotes_state = {"on": False}

    def on_toggle_quotes():
        target = not quotes_state["on"]
        fn = quote_keys_and_values if target else unquote_keys_and_values
        title = "Quoting keys & values…" if target else "Unquoting keys & values…"
        modal = ProgressModal(root, title=title)
        modal.set_progress(0.02)

        def work():
            try:
                ltxt = getattr(inp, "_textbox", inp).get("1.0", "end-1c")
                rtxt = getattr(out, "_textbox", out).get("1.0", "end-1c")

                def p1(v): modal.set_progress(0.05 + v * 0.45)
                def p2(v): modal.set_progress(0.55 + v * 0.45)

                new_l = _normalize_eol(fn(ltxt, progress=p1, cancelled=modal.cancelled))
                if modal.cancelled.is_set():
                    raise RuntimeError("Cancelled")
                new_r = rtxt.strip() and _normalize_eol(fn(rtxt, progress=p2, cancelled=modal.cancelled)) or rtxt
                if modal.cancelled.is_set():
                    raise RuntimeError("Cancelled")

                def apply_res():
                    _set_text(inp, new_l)
                    if isinstance(new_r, str):
                        _set_text(out, new_r)
                    highlighter.debounce(inp, "left")
                    highlighter.debounce(out, "right")
                    modal.close()
                    quotes_state["on"] = target
                    status_var.set("Quoted" if target else "Unquoted")

                root.after(0, apply_res)
            except Exception as ex:
                root.after(0, lambda ex=ex: (modal.close(), messagebox.showerror("Quotes", str(ex))))

        threading.Thread(target=work, daemon=True).start()

    def _style_action_widget(w):
        try:
            w.configure(corner_radius=0, fg_color=BTN_PURPLE, hover_color=BTN_PURPLE_HOVER)
        except Exception:
            pass
        return w

    # Batch dropdown (requested)
    batch_var = tk.StringVar(value="Batch ▾")

    def _on_batch_select(val: str):
        # run and then reset back to "Batch ▾"
        if val in ("Batch - Format", "Batch - Quotes"):
            _batch_process(val)
        batch_var.set("Batch ▾")

    batch_menu = ctk.CTkOptionMenu(
        btns,
        values=["Batch - Format", "Batch - Quotes"],
        variable=batch_var,
        command=_on_batch_select,
        width=160,
        corner_radius=0,
        fg_color=BTN_PURPLE,
        button_color=BTN_PURPLE,
        button_hover_color=BTN_PURPLE_HOVER,
        dropdown_fg_color=CARD_BG,
    )

    # buttons (auto-wrap) - order: Paste Left then Copy Right (requested)
    btn_widgets = [
        _style_action_widget(ctk.CTkButton(btns, text="Format", command=do_format)),
        _style_action_widget(ctk.CTkButton(btns, text="Paste Left", command=paste_left)),
        _style_action_widget(ctk.CTkButton(btns, text="Copy Right", command=copy_right)),
        _style_action_widget(ctk.CTkButton(btns, text="Open…", command=open_left)),
        _style_action_widget(ctk.CTkButton(btns, text="Save Right…", command=save_right)),
        _style_action_widget(ctk.CTkButton(btns, text="Error Check", command=_show_errors)),
        _style_action_widget(ctk.CTkButton(btns, text="Warnings", command=_show_warnings)),
        _style_action_widget(ctk.CTkButton(btns, text="Quotes", command=on_toggle_quotes, width=120)),
        batch_menu,
    ]
    for w in btn_widgets:
        w.grid_forget()

    def layout_buttons(_evt=None):
        btns.update_idletasks()
        top.update_idletasks()
        controls.update_idletasks()
        avail = max(60, top.winfo_width() - controls.winfo_width() - 48)
        x = 0
        row = 0
        col = 0
        pad = 8
        for w in btn_widgets:
            w.grid_forget()
        for w in btn_widgets:
            w_w = w.winfo_reqwidth()
            if x and x + w_w + pad > avail:
                row += 1
                col = 0
                x = 0
            w.grid(row=row, column=col, padx=4, pady=4, sticky="e")
            x += w_w + pad
            col += 1

    top.bind("<Configure>", layout_buttons)
    layout_buttons()

    # Save prefs when backup checkbox changes
    def _on_backup_toggle(*_):
        _save_pref_now()
    safety_backup_var.trace_add("write", _on_backup_toggle)

    # shortcuts
    root.bind("<Control-Return>", lambda e: (do_format(), "break"))
    root.bind("<Control-Shift-C>", lambda e: (copy_right(), "break"))
    root.bind("<Control-Shift-V>", lambda e: (paste_left(), "break"))

    # first paint
    highlighter.debounce(inp, "left")
    highlighter.debounce(out, "right")
    root.mainloop()

if __name__ == "__main__":
    run_gui()
