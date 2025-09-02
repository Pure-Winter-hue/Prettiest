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

# ------------- Robust Local Imports (Giggidy) -------------
def _robust_imports():
    try:
        from vsjsonfmt.api import format_text as _fmt
        from vsjsonfmt import cli as _cli
        return _fmt, _cli._strip_comments
    except Exception:
        here = Path(__file__).parent
        sys.path.insert(0, str(here))
        sys.path.insert(0, str(here.parent))
        from api import format_text as _fmt   # type: ignore
        from cli import _strip_comments as _sc  # type: ignore
        return _fmt, _sc

format_text, _strip_comments = _robust_imports()

def _normalize_eol(s: str) -> str:
    return s.replace("\r\n","\n").replace("\r","\n")

# =======================================================
#      Visible-Range Syntax Highlighter (Fast & Light Baby!!)
# =======================================================
TOKENS = {
    "key_imp":      {"foreground": "#ffd166"},
    "key_imp_child":{"foreground": "#f3c56b"},
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

# Visible-range highlighter: fast on big files; bolds VS-relevant keys.
class AsyncHighlighter:
    STR_RE   = re.compile(r'"(?:\\.|[^"\\])*"')
    KEY_RE   = re.compile(r'"(?:\\.|[^"\\])*"\s*:(?!:)')

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
            if tag in ("key_imp","key_imp_child","key_reg"):
                kw["font"] = self.bold_font
            try: tw.tag_configure(tag, **kw)
            except Exception: pass
        try:
            tw.tag_raise("key_imp"); tw.tag_raise("key_imp_child"); tw.tag_raise("key_reg")
            tw.tag_lower("str")
        except Exception: pass

    def _is_important_key(self, key_name: str) -> bool:
        k = key_name.lower()
        return (k in IMPORTANT_KEYS) or any(sub in k for sub in IMPORTANT_SUBSTRINGS)

    def _tokenize(self, s: str, want_full: bool) -> List[Tuple[str,int,int]]:
        spans : List[Tuple[str,int,int]] = []
        # keys
        for m in self.KEY_RE.finditer(s):
            a,b = m.span()
            colon_off = m.group(0).rfind(":")
            colon_pos = a + colon_off
            key_text  = s[a+1:colon_pos-1] if colon_pos-a>=2 else ""
            tag = "key_imp" if self._is_important_key(key_text) else "key_reg"
            spans.append((tag, a, colon_pos-1))
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
                        for tag in ("key_imp","key_imp_child","key_reg","str","num","bool","null","brace","comm","colon"):
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
    in_str=False; esc=False; in_line=False; in_block=False
    i=0; n=len(src); lineno=1; col=0
    while i<n:
        ch=src[i]; col+=1
        if ch=="\n":
            if in_str: return True, lineno, col
            in_line=False; lineno+=1; col=0; i+=1; continue
        if ch=="\r":
            if in_str: return True, lineno, col
            i+=1; continue
        if in_block:
            if ch=="*" and i+1<n and src[i+1]=="/": in_block=False; i+=2; continue
            i+=1; continue
        if in_line:
            i+=1; continue
        if not in_str and ch=="/" and i+1<n:
            if src[i+1]=="/": in_line=True; i+=2; continue
            if src[i+1]=="*": in_block=True; i+=2; continue
        if in_str:
            if esc: esc=False
            else:
                if ch=="\\": esc=True
                elif ch=='"': in_str=False
        else:
            if ch=='"': in_str=True; esc=False
        i+=1
    return False,0,0

_IDENT = re.compile(r'(?P<pre>[\s{,])(?P<key>[A-Za-z_][A-Za-z0-9_\-]*)\s*:')

def _quote_unquoted_keys_preserving_lines(s: str) -> str:
    # add quotes around bare keys, preserving newlines
    out=[]; i=0; n=len(s); in_str=False; esc=False; in_line=False; in_block=False
    while i<n:
        ch=s[i]
        if in_str:
            out.append(ch)
            if esc: esc=False
            elif ch=="\\": esc=True
            elif ch=='"': in_str=False
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
        if ch=='"':
            in_str=True; out.append(ch); i+=1; continue
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
        try:
            if ICO_PATH.exists(): self.top.iconbitmap(str(ICO_PATH))
        except Exception: pass
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
        ctk.CTkButton(card, text="Cancel", command=self.cancel, width=120).pack(pady=(0,12))
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
def _scan_string(s: str, i: int) -> int:
    i += 1; esc=False; n=len(s)
    while i<n:
        c=s[i]
        if esc: esc=False
        elif c=="\\": esc=True
        elif c=='"': return i+1
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
        if ch=='"':
            endq=_scan_string(s,i); j=endq
            while j<n and s[j].isspace(): j+=1
            if j<n and s[j]==":":
                if not to_quoted:
                    content=s[i+1:endq-1]
                    if content and (content[0].isalpha() or content[0]=='_') and IDENT_BODY.match(content[1:] or ""):
                        flush(last,i); out.append(content); out.append(s[endq:j+1]); last=j+1; i=j+1; continue
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

    root = ctk.CTk()
    root.title(APP_TITLE)
    root.configure(fg_color=DARK_BG)
    _center(root, w=1320, h=780)
    try:
        if ICO_PATH.exists(): root.iconbitmap(str(ICO_PATH))
    except Exception: pass

    # Top: controls + buttons
    top = ctk.CTkFrame(root, corner_radius=16, fg_color=CARD_BG)
    top.pack(fill="x", padx=16, pady=(16,8))
    top.grid_columnconfigure(0, weight=1)

    controls = ctk.CTkFrame(top, fg_color="transparent")
    controls.grid(row=0, column=0, sticky="w", padx=10, pady=(10,2))

    lbl_compress = ctk.CTkLabel(controls, text="Compress")
    lbl_compress.grid(row=0, column=0, padx=(0,6))
    HoverTip(lbl_compress, "Compress: Expand or compress your formatting based on number of lines in a section.")
    width_entry = ctk.CTkEntry(controls, width=72); width_entry.insert(0,"120")
    width_entry.grid(row=0, column=1, padx=(0,16))

    lbl_header = ctk.CTkLabel(controls, text="Header After")
    lbl_header.grid(row=0, column=2, padx=(0,6))
    HoverTip(lbl_header, "Header after # of lines: Formats a 'title header' that doesn't collapse in an IDE for visual identifying and organization.")
    long_entry = ctk.CTkEntry(controls, width=72); long_entry.insert(0,"20")
    long_entry.grid(row=0, column=3, padx=(0,16))

    # compact UI text size slider (debounced & no flicker)
    ui_scale_frame = ctk.CTkFrame(controls, fg_color="transparent")
    ui_scale_frame.grid(row=1, column=0, columnspan=4, sticky="w", pady=(4,4))
    ctk.CTkLabel(ui_scale_frame, text="UI Text Size").pack(side="left", padx=(0,8))
    ui_scale = ctk.CTkSlider(ui_scale_frame, from_=0.85, to=1.35, number_of_steps=10, width=260)
    ui_scale.set(1.0); ui_scale.pack(side="left")

    # Buttons container
    btns = ctk.CTkFrame(top, fg_color="transparent")
    btns.grid(row=0, column=1, sticky="e", padx=10, pady=(10,2))

    # Split panes
    split = ctk.CTkFrame(root, corner_radius=16, fg_color="transparent")
    split.pack(fill="both", expand=True, padx=16, pady=(8,16))
    left = ctk.CTkFrame(split, corner_radius=16, fg_color=CARD_BG)
    right = ctk.CTkFrame(split, corner_radius=16, fg_color=CARD_BG)
    left.pack(side="left", fill="both", expand=True, padx=(0,8))
    right.pack(side="left", fill="both", expand=True, padx=(8,0))

    ctk.CTkLabel(left, text="Input (JSON / JSONC)", text_color=ACCENT_2).pack(anchor="w", padx=12, pady=(12,4))
    ctk.CTkLabel(right, text="Output (Formatted JSON)", text_color=ACCENT_2).pack(anchor="w", padx=12, pady=(12,4))

    # text widgets
    inp = ctk.CTkTextbox(left, corner_radius=10, wrap="none", font=("Consolas", 12))
    out = ctk.CTkTextbox(right, corner_radius=10, wrap="none", font=("Consolas", 12))
    inp.pack(fill="both", expand=True, padx=12, pady=(0,12))
    out.pack(fill="both", expand=True, padx=12, pady=(0,12))

    # footer & credit
    status = ctk.CTkFrame(root, corner_radius=12, fg_color=CARD_BG)
    status.pack(fill="x", padx=16, pady=(0,12))
    status_var = tk.StringVar(value="Ready")
    ctk.CTkLabel(status, textvariable=status_var, text_color=TEXT_MUT).pack(side="left", padx=10, pady=6)

    credit_row = ctk.CTkFrame(status, fg_color="transparent")
    credit_row.pack(side="right", padx=10, pady=6)
    badge_img = None
    try:
        from PIL import Image
        if PNG_PATH.exists():
            badge_img = ctk.CTkImage(Image.open(PNG_PATH), size=(16,16))
    except Exception:
        badge_img = None
    ctk.CTkLabel(credit_row, image=badge_img, text="  By Pure Winter", compound="left").pack()

    # highlighter
    PERF_THRESHOLD = 120_000
    highlighter = AsyncHighlighter(root, PERF_THRESHOLD)
    highlighter.watch(inp,"left"); highlighter.watch(out,"right")

    # scaled fonts (no flicker) — debounce + pause highlighter
    _scale_job = {"id": None}
    def set_ui_scale_now(scale: float):
        base = max(10, int(12 * scale))
        inp.configure(font=("Consolas", base))
        out.configure(font=("Consolas", base))
        # small re-highlight when done
        highlighter.set_enabled(True)
        highlighter.debounce(inp,"left"); highlighter.debounce(out,"right")

    def on_scale(val):
        # pause highlighting while the user drags
        highlighter.set_enabled(False)
        if _scale_job["id"] is not None:
            try: root.after_cancel(_scale_job["id"])
            except Exception: pass
        _scale_job["id"] = root.after(120, lambda v=float(val): set_ui_scale_now(v))

    ui_scale.configure(command=on_scale)

    # helpers
    def _numbers() -> tuple[int,int]:
        try: w = int(width_entry.get() or "120")
        except Exception: w = 120
        try: l = int(long_entry.get() or "20")
        except Exception: l = 20
        return w, l

    def _set_text(widget: tk.Text, text: str):
        tw = getattr(widget,"_textbox",widget)
        tw.delete("1.0","end"); tw.insert("1.0", text)

    # incremental insert with progress
    def _insert_text_incremental(widget: tk.Text, text: str, modal: ProgressModal,
                                 start=0.1, end=1.0, side="left", done_cb: Callable|None=None):
        tw = getattr(widget, "_textbox", widget)
        text = _normalize_eol(text); total = len(text)
        if total == 0:
            modal.set_progress(end); done_cb and done_cb(); return
        tw.delete("1.0","end")
        chunk = max(250_000, total//120); idx = 0
        def step():
            nonlocal idx
            if modal.cancelled.is_set(): modal.close(); done_cb and done_cb(); return
            j = min(idx+chunk, total)
            tw.insert("end", text[idx:j]); idx=j
            modal.set_progress(start + (end-start) * (idx/total))
            if idx < total:
                tw.update_idletasks(); root.after(1, step)
            else:
                highlighter.debounce(widget, side); done_cb and done_cb()
        root.after(0, step)

    # dialogs
    def _show_text_dialog(title: str, text: str, *,
                          action_title: str|None=None,
                          action: Callable[[], None] | None = None):
        dlg = ctk.CTkToplevel(root); dlg.title(title); dlg.attributes("-topmost", True); dlg.grab_set()
        try:
            if ICO_PATH.exists(): dlg.iconbitmap(str(ICO_PATH))
        except Exception: pass
        w,h = 820, 420
        rx,ry = root.winfo_rootx(), root.winfo_rooty()
        rw,rh = root.winfo_width(), root.winfo_height()
        x = rx + (rw - w)//2; y = ry + (rh - h)//2
        dlg.geometry(f"{w}x{h}+{max(0,x)}+{max(0,y)}")
        container = ctk.CTkFrame(dlg, fg_color=CARD_BG, corner_radius=16)
        container.pack(fill="both", expand=True, padx=12, pady=12)
        t = ctk.CTkTextbox(container, wrap="word", corner_radius=10)
        t.configure(font=("Consolas", 14))
        t.pack(fill="both", expand=True, padx=8, pady=(8,8))
        t.insert("1.0", text.strip()+"\n")
        t.configure(state="disabled")
        row = ctk.CTkFrame(container, fg_color="transparent")
        row.pack(fill="x", padx=8, pady=(0,8))
        if action_title and action:
            ctk.CTkButton(row, text=action_title, width=160, command=lambda: (dlg.destroy(), action())).pack(side="right", padx=8)
        ctk.CTkButton(row, text="OK", width=100, command=dlg.destroy).pack(side="right")

    # actions
    def do_format():
        w, l = _numbers()
        src = getattr(inp,"_textbox",inp).get("1.0","end-1c")
        if not src.strip():
            out.delete("1.0","end"); status_var.set("Nothing to format"); return
        modal = ProgressModal(root, title="Formatting…"); modal.set_progress(0.02)
        def work():
            try:
                formatted = _normalize_eol(format_text(src, print_width=w, long_block_threshold=l, aggressive_inline=True))
                def after_insert():
                    highlighter.debounce(out,"right"); modal.close(); status_var.set("Formatted ✓")
                _insert_text_incremental(out, formatted, modal, start=0.1, end=1.0, side="right", done_cb=after_insert)
            except Exception as ex:
                root.after(0, lambda ex=ex: (modal.close(), messagebox.showerror("Format", str(ex))))
        threading.Thread(target=work, daemon=True).start()

    def _fast_copy(text: str) -> bool:
        if _HAVE_WIN32:
            try: _copy_win32(text); return True
            except Exception: pass
        if pyperclip:
            try: pyperclip.copy(text); return True
            except Exception: pass
        try:
            root.clipboard_clear(); root.clipboard_append(text); return True
        except Exception: return False

    def copy_right():
        text = getattr(out,"_textbox",out).get("1.0","end-1c")
        status_var.set("Copied ✓" if _fast_copy(text) else "Copy failed")

    def paste_left():
        modal = ProgressModal(root, title="Pasting…"); modal.set_progress(0.05)
        def bg():
            try:
                if pyperclip: text = _normalize_eol(pyperclip.paste())
                else: text = root.clipboard_get()
                def after_insert():
                    highlighter.debounce(inp, "left"); modal.close(); status_var.set("Pasted ✓")
                _insert_text_incremental(inp, text, modal, start=0.1, end=1.0, side="left", done_cb=after_insert)
            except Exception as ex:
                root.after(0, lambda ex=ex: (modal.close(), messagebox.showerror("Paste", str(ex))))
        threading.Thread(target=bg, daemon=True).start()

    def _read_file_chunked(path: str, modal: ProgressModal) -> Optional[str]:
        try: size = max(1, os.path.getsize(path))
        except Exception: size = 1
        buf = bytearray(); read = 0
        with open(path, "rb") as f:
            while True:
                if modal.cancelled.is_set(): return None
                chunk = f.read(4<<20)
                if not chunk: break
                buf.extend(chunk); read += len(chunk)
                modal.set_progress(min(0.49, 0.49 * read / size))
        return buf.decode("utf-8-sig")

    def open_left():
        p = filedialog.askopenfilename(filetypes=[
            ("JSON/JSONC","*.json;*.jsonc;*.patch.json;*.patch"),
            ("All files","*.*"),
        ])
        if not p: return
        modal = ProgressModal(root, title="Opening…"); modal.set_progress(0.02)
        def work():
            try:
                raw = _read_file_chunked(p, modal)
                text = None if raw is None else _normalize_eol(raw)
                def after_insert():
                    modal.close()
                    if text is None: status_var.set("Open cancelled")
                    else: status_var.set(f"Loaded: {Path(p).name}")
                if text is None: root.after(0, after_insert)
                else: _insert_text_incremental(inp, text, modal, start=0.5, end=1.0, side="left", done_cb=after_insert)
            except Exception as ex:
                root.after(0, lambda ex=ex: (modal.close(), messagebox.showerror("Open", str(ex))))
        threading.Thread(target=work, daemon=True).start()

    def save_right():
        p = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON","*.json")])
        if not p: return
        try:
            data = _normalize_eol(getattr(out,"_textbox",out).get("1.0","end-1c"))
            with open(p, "w", encoding="utf-8", newline="\n") as f:
                f.write(data)
            status_var.set(f"Saved: {Path(p).name}")
        except Exception as ex:
            messagebox.showerror("Save", str(ex))

    # Warnings / Errors
    def _show_warnings():
        src = getattr(inp,"_textbox",inp).get("1.0","end-1c")
        if not src.strip():
            messagebox.showinfo("Vintage Story JSON check", "Left pane is empty."); return
        warns = _vs_warnings(src)
        body = "No warnings." if not warns else ("Warnings:\n• " + "\n• ".join(warns))
        _show_text_dialog("Vintage Story JSON check", body)

    def _strictish(src: str) -> str:
        # quote unquoted keys but keep line numbers stable; remove trailing commas
        return _remove_trailing_commas(_quote_unquoted_keys_preserving_lines(_strip_comments(src)))

    def _show_errors():
        src = getattr(inp,"_textbox",inp).get("1.0","end-1c")
        if not src.strip():
            messagebox.showinfo("Vintage Story JSON check", "Left pane is empty."); return

        # accidental return inside string?
        brk, ln, col = _detect_accidental_return(src)
        if brk:
            line = src.splitlines()[ln-1] if 1<=ln<=len(src.splitlines()) else ""
            pointer = " "*(max(0,col-1)) + "^"
            msg = (f"Invalid control character at\n(Line {ln})\n\n{line}\n{pointer}\n\n"
                   "Hint: Detected a line break inside a quoted string (accidental Return).")
            highlight_error_line(inp, ln)
            _show_text_dialog("Vintage Story JSON check", msg)
            return

        strictish = _strictish(src)
        try:
            json.loads(strictish)
            _show_text_dialog("Vintage Story JSON check", "No errors found.")
        except json.JSONDecodeError as e:
            ln = max(1, e.lineno); col = max(1, e.colno)
            line = strictish.splitlines()[ln-1] if 1<=ln<=len(strictish.splitlines()) else ""
            pointer = " "*(col-1)+"^"
            msg = (f"{e.msg}\n\n(Line {ln}, column {col})\n{line}\n{pointer}\n\n"
                   "Tip: This is likely a missing comma between items or a broken structure.")
            highlight_error_line(inp, ln)

            # offer auto-fix when the parser says “Expecting ',' delimiter”
            if "Expecting ',' delimiter" in e.msg:
                def do_fix():
                    new_src, fixed_ln = _autofix_insert_comma_at(src, ln, col)
                    _set_text(inp, new_src)
                    highlight_error_line(inp, fixed_ln)
                    status_var.set("Inserted missing comma")
                    highlighter.debounce(inp,"left")
                _show_text_dialog("Vintage Story JSON check", msg, action_title="Auto-fix: Insert comma", action=do_fix)
            else:
                _show_text_dialog("Vintage Story JSON check", msg)

    # quotes toggle (single button)
    quotes_state = {"on": False}
    def on_toggle_quotes():
        target = not quotes_state["on"]
        fn = quote_keys_and_values if target else unquote_keys_and_values
        title = "Quoting keys & values…" if target else "Unquoting keys & values…"
        modal = ProgressModal(root, title=title); modal.set_progress(0.02)
        def work():
            try:
                ltxt = getattr(inp,"_textbox",inp).get("1.0","end-1c")
                rtxt = getattr(out,"_textbox",out).get("1.0","end-1c")
                def p1(v): modal.set_progress(0.05 + v*0.45)
                def p2(v): modal.set_progress(0.55 + v*0.45)
                new_l = _normalize_eol(fn(ltxt, progress=p1, cancelled=modal.cancelled))
                if modal.cancelled.is_set(): raise RuntimeError("Cancelled")
                new_r = rtxt.strip() and _normalize_eol(fn(rtxt, progress=p2, cancelled=modal.cancelled)) or rtxt
                if modal.cancelled.is_set(): raise RuntimeError("Cancelled")
                def apply_res():
                    _set_text(inp, new_l)
                    if isinstance(new_r,str): _set_text(out, new_r)
                    highlighter.debounce(inp,"left"); highlighter.debounce(out,"right")
                    modal.close(); quotes_state["on"] = target
                    status_var.set("Quoted" if target else "Unquoted")
                root.after(0, apply_res)
            except Exception as ex:
                root.after(0, lambda ex=ex: (modal.close(), messagebox.showerror("Quotes", str(ex))))
        threading.Thread(target=work, daemon=True).start()

    # buttons (auto-wrap)
    btn_widgets = [
        ctk.CTkButton(btns, text="Format",      command=do_format,   corner_radius=12),
        ctk.CTkButton(btns, text="Copy Right",  command=copy_right,  corner_radius=12),
        ctk.CTkButton(btns, text="Paste Left",  command=paste_left,  corner_radius=12),
        ctk.CTkButton(btns, text="Open...",     command=open_left,   corner_radius=12),
        ctk.CTkButton(btns, text="Save Right...", command=save_right, corner_radius=12),
        ctk.CTkButton(btns, text="Error Check", command=_show_errors, corner_radius=12),
        ctk.CTkButton(btns, text="Warnings",    command=_show_warnings, corner_radius=12),
        ctk.CTkButton(btns, text="Quotes",      command=on_toggle_quotes, corner_radius=12, width=120),
    ]
    for b in btn_widgets: b.grid_forget()

    def layout_buttons(_evt=None):
        btns.update_idletasks(); top.update_idletasks(); controls.update_idletasks()
        avail = max(60, top.winfo_width() - controls.winfo_width() - 48)
        x = 0; row = 0; col = 0; pad = 8
        for b in btn_widgets: b.grid_forget()
        for b in btn_widgets:
            w = b.winfo_reqwidth()
            if x and x + w + pad > avail:
                row += 1; col = 0; x = 0
            b.grid(row=row, column=col, padx=4, pady=4, sticky="e")
            x += w + pad; col += 1
    top.bind("<Configure>", layout_buttons); layout_buttons()

    # shortcuts
    root.bind("<Control-Return>", lambda e: (do_format(), "break"))
    root.bind("<Control-Shift-C>", lambda e: (copy_right(), "break"))
    root.bind("<Control-Shift-V>", lambda e: (paste_left(), "break"))

    # first paint
    highlighter.debounce(inp,"left"); highlighter.debounce(out,"right")
    root.mainloop()

if __name__ == "__main__":
    run_gui()
