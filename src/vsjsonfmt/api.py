from __future__ import annotations
from typing import Any, Optional
from .cli import parse_json_lenient, format_json, FormatConfig

# Default config mirrors your preferences
DEFAULT_CFG = FormatConfig()

def format_text(
    text: str,
    *,
    indent: Optional[int] = None,
    print_width: Optional[int] = None,
    aggressive_inline: Optional[bool] = None,
    long_block_threshold: Optional[int] = None,
) -> str:
    """Format JSON/JSONC string with pretty, width-aware layout (preserving key order)."""
    try:
        data = parse_json_lenient(text)
    except Exception as ex:
        raise ValueError(f"Parse error: {ex}")

    cfg = DEFAULT_CFG
    if any(v is not None for v in (indent, print_width, aggressive_inline, long_block_threshold)):
        cfg = FormatConfig(
            indent = indent if indent is not None else DEFAULT_CFG.indent,
            print_width = print_width if print_width is not None else DEFAULT_CFG.print_width,
            array_wrap = DEFAULT_CFG.array_wrap,
            object_wrap = DEFAULT_CFG.object_wrap,
            aggressive_inline = aggressive_inline if aggressive_inline is not None else DEFAULT_CFG.aggressive_inline,
            long_block_threshold = long_block_threshold if long_block_threshold is not None else DEFAULT_CFG.long_block_threshold,
            eol = DEFAULT_CFG.eol,
        )
    return format_json(data, cfg)
