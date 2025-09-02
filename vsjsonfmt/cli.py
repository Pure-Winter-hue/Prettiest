from __future__ import annotations
import argparse, json, re, sys, difflib
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Tuple, Optional

# ------------------------------ Utilities ------------------------------
def _read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")

def _write_text(path: Path, data: str) -> None:
    # Normalize to LF
    data = data.replace("\r\n", "\n").replace("\r", "\n")
    if not data.endswith("\n"):
        data += "\n"
    path.write_text(data, encoding="utf-8")

# ------------------------------ JSONC support ------------------------------
_STRING_RE = re.compile(r'"(?:\\.|[^"\\])*"')
_SLASH_SLASH_COMMENT_RE = re.compile(r"//.*?(?=\n|$)")
_SLASH_STAR_COMMENT_RE = re.compile(r"/\*.*?\*/", re.S)

def _strip_comments(jsonc: str) -> str:
    parts: List[str] = []
    i = 0
    for m in _STRING_RE.finditer(jsonc):
        gap = jsonc[i:m.start()]
        gap = _SLASH_SLASH_COMMENT_RE.sub("", gap)
        gap = _SLASH_STAR_COMMENT_RE.sub("", gap)
        parts.append(gap)
        parts.append(jsonc[m.start():m.end()])
        i = m.end()
    tail = jsonc[i:]
    tail = _SLASH_SLASH_COMMENT_RE.sub("", tail)
    tail = _SLASH_STAR_COMMENT_RE.sub("", tail)
    parts.append(tail)
    return "".join(parts)

def _strip_trailing_commas(s: str) -> str:
    out: List[str] = []
    i = 0; in_string = False; escape = False
    while i < len(s):
        ch = s[i]
        if in_string:
            out.append(ch)
            if escape: escape = False
            elif ch == "\\": escape = True
            elif ch == '"': in_string = False
            i += 1; continue
        if ch == '"':
            in_string = True; escape = False; out.append(ch); i += 1; continue
        if ch in "}]":
            j = len(out) - 1
            while j >= 0 and out[j].isspace(): j -= 1
            if j >= 0 and out[j] == ",": out.pop(j)
            out.append(ch); i += 1; continue
        out.append(ch); i += 1
    return "".join(out)

def parse_json_lenient(text: str) -> Any:
    no_comments = _strip_comments(text)
    no_trailing = _strip_trailing_commas(no_comments)
    return json.loads(no_trailing)  # key order preserved (3.7+)

# ------------------------------ Formatting ------------------------------
@dataclass
class FormatConfig:
    indent: int = 2
    print_width: int = 120
    array_wrap: str = "collapse"       # collapse | preserve
    object_wrap: str = "collapse"      # collapse | preserve
    aggressive_inline: bool = True
    long_block_threshold: int = 20
    eol: str = "\n"

    # Arrays under these keys try to stay compact (inline or one-line-per-element)
    array_collapse_keys: Tuple[str, ...] = (
        "requireStacks", "addElements", "removeElements",
        "ingredients", "outputs", "drops"
    )

    # Global rule: arrays of homogeneous small objects → one compact object per line
    inline_homogeneous_objects: bool = True
    inline_homogeneous_min: int = 3
    inline_string_max: int = 120
    allow_overflow_inline_arrays: bool = True  # allow slight overflow for aesthetics

def _is_prim(x: Any) -> bool:
    return isinstance(x, (str, int, float, bool)) or x is None

def _dump_oneline(val: Any) -> str:
    return json.dumps(val, ensure_ascii=False, separators=(",", ": "))

def _fits_width(s: str, cfg: FormatConfig, level: int) -> bool:
    return len(s) + level * cfg.indent <= cfg.print_width

def _line_count(s: str) -> int:
    return s.count("\n") + 1

def _keys_tuple(d: Dict[str, Any]) -> Tuple[str, ...]:
    return tuple(d.keys())

def _is_small_prim(v: Any, cfg: FormatConfig) -> bool:
    if not _is_prim(v): return False
    if isinstance(v, str) and len(v) > cfg.inline_string_max: return False
    return True

def _array_is_homog_small_objs(arr: List[Any], cfg: FormatConfig) -> Tuple[bool, Tuple[str, ...]]:
    if not arr or len(arr) < cfg.inline_homogeneous_min: return (False, ())
    if not all(isinstance(x, dict) for x in arr): return (False, ())
    key_order = _keys_tuple(arr[0])
    if not key_order: return (False, ())
    for d in arr:
        if _keys_tuple(d) != key_order: return (False, ())
        for v in d.values():
            if not _is_small_prim(v, cfg): return (False, ())
    return (True, key_order)

# ---- compact helpers -----------------------------------------------------
def _compact_object(d: Dict[str, Any], cfg: FormatConfig) -> Optional[str]:
    # only primitives/small strings allowed
    for v in d.values():
        if not _is_small_prim(v, cfg):
            return None
    return "{ " + ", ".join(f"{json.dumps(k)}: {_dump_oneline(v)}" for k, v in d.items()) + " }"

def _compact_list(lst: List[Any], cfg: FormatConfig, level: int) -> Optional[str]:
    parts: List[str] = []
    for el in lst:
        if _is_prim(el):
            parts.append(_dump_oneline(el)); continue
        if isinstance(el, dict):
            c = _compact_object(el, cfg)
            if c is None: return None
            parts.append(c); continue
        if isinstance(el, list):
            c = _compact_list(el, cfg, level)  # nested list compact
            if c is None: return None
            parts.append(c); continue
        return None
    one = "[ " + ", ".join(parts) + " ]"
    return one if (_fits_width(one, cfg, level) or cfg.allow_overflow_inline_arrays) else None

def _format_value(val: Any, cfg: FormatConfig, level: int, parent_key: Optional[str] = None) -> str:
    indent = " " * (cfg.indent * level)
    child  = " " * (cfg.indent * (level + 1))

    # Primitive
    if _is_prim(val):
        return _dump_oneline(val)

    # List (array)
    if isinstance(val, list):
        # If parent is a collapse-key, try very compact forms first.
        if isinstance(parent_key, str) and parent_key in cfg.array_collapse_keys:
            # 1) Entire array inline if it fits (handles arrays and arrays-of-arrays)
            whole_inline = _compact_list(val, cfg, level)
            if whole_inline is not None:
                return whole_inline
            # 2) Otherwise: keep each element compact on a single line
            #    (e.g., each inner array or object rendered inline)
            elem_strings: List[str] = []
            for el in val:
                if _is_prim(el):
                    elem_strings.append(_dump_oneline(el))
                elif isinstance(el, dict):
                    c = _compact_object(el, cfg)
                    if c is None: break
                    elem_strings.append(c)
                elif isinstance(el, list):
                    c = _compact_list(el, cfg, level + 1)
                    if c is None: break
                    elem_strings.append(c)
                else:
                    break
            else:
                parts = [child + s for s in elem_strings]
                return "[\n" + ",\n".join(parts) + "\n" + indent + "]"
            # If we couldn't keep it compact, fall through to normal formatting.

        # Global rule: homogeneous small objects → one compact object per line
        if cfg.inline_homogeneous_objects:
            ok, key_order = _array_is_homog_small_objs(val, cfg)
            if ok:
                parts: List[str] = []
                for el in val:
                    el_sorted = {k: el[k] for k in key_order}
                    one = _compact_object(el_sorted, cfg) or _dump_oneline(el_sorted)
                    parts.append(child + one)
                return "[\n" + ",\n".join(parts) + "\n" + indent + "]"

        # Collapse list of primitives when it fits
        if cfg.array_wrap == "collapse" and val and all(_is_prim(x) for x in val):
            one = "[" + ", ".join(_dump_oneline(x) for x in val) + "]"
            if _fits_width(one, cfg, level) and _line_count(one) <= cfg.long_block_threshold:
                return one

        # Default multiline (recurse)
        parts: List[str] = []
        for el in val:
            parts.append(child + _format_value(el, cfg, level + 1))
        return "[\n" + ",\n".join(parts) + "\n" + indent + "]"

    # Dict
    if isinstance(val, dict):
        items = list(val.items())

        # Inline compact object when possible
        if cfg.object_wrap == "collapse":
            if (all(_is_prim(v) for _, v in items) or cfg.aggressive_inline) and items:
                inner_chunks: Optional[List[str]] = []
                for k, v in items:
                    if _is_prim(v):
                        inner_chunks.append(f"{json.dumps(k)}: {_dump_oneline(v)}")
                    else:
                        candidate = _format_value(v, cfg, level, parent_key=k)
                        if "\n" in candidate:
                            inner_chunks = None
                            break
                        inner_chunks.append(f"{json.dumps(k)}: {candidate}")
                if inner_chunks is not None:
                    one = "{ " + ", ".join(inner_chunks) + " }"
                    if _fits_width(one, cfg, level) and _line_count(one) <= cfg.long_block_threshold:
                        return one

        # Multiline object
        lines: List[str] = []
        for k, v in items:
            formatted = _format_value(v, cfg, level + 1, parent_key=k)
            lines.append(f"{child}{json.dumps(k)}: {formatted}")
        return "{\n" + ",\n".join(lines) + "\n" + indent + "}"

    # Fallback
    return _dump_oneline(val)

def format_json(data: Any, cfg: FormatConfig) -> str:
    return _format_value(data, cfg, 0, parent_key=None) + cfg.eol

# ------------------------------ CLI ------------------------------
DEFAULT_CONFIG: Dict[str, Any] = {
    "indent": 2,
    "print_width": 120,
    "array_wrap": "collapse",
    "object_wrap": "collapse",
    "aggressive_inline": True,
    "long_block_threshold": 20,

    "array_collapse_keys": ("requireStacks", "addElements", "removeElements",
                            "ingredients", "outputs", "drops"),

    "inline_homogeneous_objects": True,
    "inline_homogeneous_min": 3,
    "inline_string_max": 120,
    "allow_overflow_inline_arrays": True,
}

def _load_project_config(start_dir: Path) -> Dict[str, Any]:
    cur = start_dir
    root = Path(cur.anchor)
    while True:
        p = cur / "vsjsonfmt.config.json"
        if p.exists():
            try:
                return json.loads(p.read_text(encoding="utf-8"))
            except Exception:
                return {}
        if cur == root:
            break
        cur = cur.parent
    return {}

def _build_config(args) -> FormatConfig:
    cfg = DEFAULT_CONFIG.copy()
    cfg.update(_load_project_config(Path.cwd()))

    if args.indent is not None: cfg["indent"] = args.indent
    if args.print_width is not None: cfg["print_width"] = args.print_width
    if args.long_block is not None: cfg["long_block_threshold"] = args.long_block
    if args.no_aggressive: cfg["aggressive_inline"] = False
    if args.no_inline_homogeneous: cfg["inline_homogeneous_objects"] = False
    if args.collapse_keys:
        cfg["array_collapse_keys"] = tuple(k.strip() for k in args.collapse_keys.split(",") if k.strip())
    return FormatConfig(**cfg)

def _expand_files(patterns: List[str]) -> List[Path]:
    out: List[Path] = []
    for pat in patterns:
        if any(ch in pat for ch in "*?[]"):
            out.extend([p for p in Path().glob(pat) if p.is_file()])
        else:
            p = Path(pat)
            if p.is_file():
                out.append(p)
    seen, uniq = set(), []
    for p in out:
        rp = p.resolve()
        if rp not in seen:
            seen.add(rp); uniq.append(p)
    return uniq

def _process_text(text: str, cfg: FormatConfig) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    data = parse_json_lenient(text)
    return format_json(data, cfg)

def main(argv: List[str] | None = None) -> int:
    if argv is None:
        argv = sys.argv[1:]
    ap = argparse.ArgumentParser(description="Vintage Story JSON formatter (vsjsonfmt)")
    ap.add_argument("paths", nargs="*", help="Files or globs to format")
    ap.add_argument("--write", "-w", action="store_true", help="Write changes to files")
    ap.add_argument("--check", action="store_true", help="Exit 1 if any files would be changed")
    ap.add_argument("--diff", action="store_true", help="Show unified diff for changes")
    ap.add_argument("--stdin", action="store_true", help="Read from stdin and write to stdout")
    ap.add_argument("--indent", type=int, help="Indent size (default 2)")
    ap.add_argument("--print-width", type=int, help="Max line width (default 120)")
    ap.add_argument("--long-block", type=int, help="Long block threshold in lines (default 20)")
    ap.add_argument("--no-aggressive", action="store_true", help="Disable aggressive inlining of short objects")
    ap.add_argument("--no-inline-homogeneous", action="store_true",
                    help="Disable the one-per-line behavior for homogeneous object arrays")
    ap.add_argument("--collapse-keys",
                    help="Comma-separated keys whose arrays try to stay compact (e.g. 'requireStacks,addElements')")
    args = ap.parse_args(argv)

    cfg = _build_config(args)

    if args.stdin:
        sys.stdout.write(_process_text(sys.stdin.read(), cfg))
        return 0

    files = _expand_files(args.paths) if args.paths else []
    if not files:
        print("vsjsonfmt: No input files. Provide paths or use --stdin.", file=sys.stderr)
        return 2

    changed = 0
    for f in files:
        original = _read_text(f)
        try:
            formatted = _process_text(original, cfg)
        except Exception as ex:
            print(f"{f}: error: {ex}", file=sys.stderr)
            continue

        norm_original = original.replace("\r\n", "\n").replace("\r", "\n")
        if not norm_original.endswith("\n"):
            norm_original += "\n"

        if formatted != norm_original:
            changed += 1
            if args.diff and not args.write:
                diff = difflib.unified_diff(
                    norm_original.splitlines(keepends=True),
                    formatted.splitlines(keepends=True),
                    fromfile=str(f),
                    tofile=str(f) + " (formatted)",
                )
                sys.stdout.writelines(diff)
            if args.write:
                _write_text(f, formatted)

    if args.check and changed:
        return 1
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
