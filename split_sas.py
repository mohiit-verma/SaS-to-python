# Databricks/Notebook-safe: no file I/O required unless you call split_sas_file.
from __future__ import annotations

import re
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional, Tuple

# --------------------------------------------------------------------------------------
# Configuration and Logging
# --------------------------------------------------------------------------------------

@dataclass(frozen=True)
class SplitConfig:
    """
    Configuration for splitting SAS code into chunks.

    Parameters
    ----------
    max_lines : int
        Maximum number of lines allowed in each output chunk. Oversized logical
        blocks (e.g., a very large PROC step) will be split safely on soft
        boundaries when possible.
    prefer_blank_split : bool
        When splitting an oversized logical block, prefer to cut on blank lines
        (for readability) before falling back to statement boundaries or hard
        limits.
    keep_empty_trailing_lines : bool
        If True, preserves trailing blank lines within chunks; otherwise trims them.
    filename_prefix : Optional[str]
        Optional prefix used by `split_sas_file` when writing output files. If not
        provided, the input file’s stem is used.
    """
    max_lines: int = 200
    prefer_blank_split: bool = True
    keep_empty_trailing_lines: bool = False
    filename_prefix: Optional[str] = None


logger = logging.getLogger("sas_splitter")
if not logger.handlers:
    # Light default logging; users can override in their notebooks.
    handler = logging.StreamHandler()
    formatter = logging.Formatter("%(levelname)s | %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)
logger.setLevel(logging.INFO)

# --------------------------------------------------------------------------------------
# Regexes (compiled once) — case-insensitive and whitespace-tolerant
# --------------------------------------------------------------------------------------

START_STEP_RE = re.compile(r'^\s*(?:proc\s+\w+|data\s+\S+)\b', re.IGNORECASE)
END_STEP_RE   = re.compile(r'^\s*(?:run|quit)\s*;\s*(?:\*.*)?$', re.IGNORECASE)

MACRO_START_RE = re.compile(r'^\s*%macro\b', re.IGNORECASE)
MACRO_END_RE   = re.compile(r'^\s*%mend\b', re.IGNORECASE)

# SAS statement comment "* blah blah ;" (must end with semicolon on same line)
STMT_COMMENT_RE = re.compile(r'^\s*\*[^;]*;\s*$')

# --------------------------------------------------------------------------------------
# Comment handling and normalization
# --------------------------------------------------------------------------------------

def strip_for_detection(line: str, in_block_comment: bool) -> Tuple[str, bool]:
    """
    Remove SAS comments purely for boundary detection (not for output).

    This function preserves the original `line` for output elsewhere, but returns a
    "detection-safe" version with comments removed, plus the updated block-comment
    state. It supports:
      - Multi-line block comments: /* ... */
      - Statement comments: * ... ;

    Parameters
    ----------
    line : str
        The raw input line from a SAS file.
    in_block_comment : bool
        Whether the parser is currently inside a multi-line /* ... */ comment.

    Returns
    -------
    Tuple[str, bool]
        A tuple of (normalized_line_for_detection, new_in_block_comment_state).

    Notes
    -----
    - This is intentionally minimal and conservative: content is only removed for
      detection purposes. The original line should always be written to output.
    - Nested /* */ comments are not standard SAS; if they appear, they will be
      treated as linear "first close wins".
    """
    original = line

    # Handle statement-level comments like: * this is a comment ;
    if STMT_COMMENT_RE.match(line):
        return ("", in_block_comment)

    s = line
    i = 0
    out = []
    while i < len(s):
        if in_block_comment:
            end = s.find("*/", i)
            if end == -1:
                # Entire remainder is in comment
                i = len(s)
                continue
            else:
                i = end + 2
                in_block_comment = False
                continue
        else:
            start = s.find("/*", i)
            if start == -1:
                out.append(s[i:])
                break
            else:
                out.append(s[i:start])
                i = start + 2
                in_block_comment = True

    cleaned = "".join(out)
    # Normalize whitespace for simpler regex matches
    cleaned = cleaned.strip()

    return (cleaned, in_block_comment)

# --------------------------------------------------------------------------------------
# Logical block detection
# --------------------------------------------------------------------------------------

def iter_logical_blocks(lines: Iterable[str]) -> List[List[str]]:
    """
    Group SAS source lines into logical blocks: DATA/PROC steps and %MACRO blocks.

    This routine scans the input line-by-line, ignoring comments for the purposes of
    detection, and yields blocks that are *semantically* coherent for SAS:
      - A PROC/DATA step starts at `proc <name>` or `data <name>` and ends at
        `run;` or `quit;` (case-insensitive), unless inside a %macro.
      - A %macro block starts at `%macro` and ends at `%mend`. Macro depth is
        tracked so nested macros are handled safely.
      - Lines outside of any step/macro are grouped into "loose" blocks (e.g.,
        OPTIONS, %INCLUDE, LIBNAME, etc.).

    Parameters
    ----------
    lines : Iterable[str]
        Iterable of raw SAS lines.

    Returns
    -------
    List[List[str]]
        A list of blocks, each block being a list of original lines (unmodified).

    Design & Robustness
    -------------------
    - Ignores multi-line and statement-level comments when detecting starts/ends.
    - Keeps `%macro` content intact—even if it contains RUN/QUIT.
    - Handles nested `%macro` definitions by tracking depth.
    """
    blocks: List[List[str]] = []
    buf: List[str] = []

    in_step = False
    macro_depth = 0
    in_block_comment = False

    def flush_buffer():
        nonlocal buf
        if buf:
            blocks.append(buf)
            buf = []

    for raw in lines:
        det, in_block_comment = strip_for_detection(raw, in_block_comment)

        # Macro start/end take precedence
        if MACRO_START_RE.match(det):
            # Current loose or step buffer ends here
            flush_buffer()
            macro_depth += 1
            buf.append(raw)
            continue

        if macro_depth > 0:
            buf.append(raw)
            if MACRO_END_RE.match(det):
                macro_depth -= 1
                if macro_depth == 0:
                    flush_buffer()
            continue

        # Outside of macros, watch for PROC/DATA boundaries
        if not in_step and START_STEP_RE.match(det):
            flush_buffer()
            in_step = True
            buf.append(raw)
            continue

        if in_step:
            buf.append(raw)
            if END_STEP_RE.match(det):
                flush_buffer()
                in_step = False
            continue

        # Loose text (options, %include, libname, comments, whitespace, etc.)
        buf.append(raw)

    # Any remainder
    if buf:
        blocks.append(buf)

    return blocks

# --------------------------------------------------------------------------------------
# Block splitting and packing
# --------------------------------------------------------------------------------------

def split_oversized_block(block: List[str], cfg: SplitConfig) -> List[List[str]]:
    """
    Split a single logical block if it exceeds `cfg.max_lines`.

    Splitting strategy (in order of preference):
      1. Split on blank lines near the boundary (if `prefer_blank_split`).
      2. Split on a statement boundary (a line that ends with ';').
      3. Hard split at `max_lines`.

    Parameters
    ----------
    block : List[str]
        The logical block to evaluate.
    cfg : SplitConfig
        Configuration options.

    Returns
    -------
    List[List[str]]
        One or more sub-blocks whose lengths are <= `cfg.max_lines`.

    Notes
    -----
    - We never drop or rewrite lines; only the *cut point* is chosen.
    - This is intentionally conservative to avoid breaking statements.
    """
    n = len(block)
    if n <= cfg.max_lines:
        return [block]

    parts: List[List[str]] = []
    start = 0
    while start < n:
        end = min(start + cfg.max_lines, n)

        # 1) Prefer blank line just before the end
        cut = None
        if cfg.prefer_blank_split:
            for i in range(end - 1, start, -1):
                if not block[i].strip():
                    cut = i + 1
                    break

        # 2) Otherwise, look for the nearest statement boundary (line ending with ;)
        if cut is None:
            for i in range(end - 1, start, -1):
                if block[i].rstrip().endswith(";"):
                    cut = i + 1
                    break

        # 3) Hard cut
        if cut is None or cut == start:
            cut = end

        parts.append(block[start:cut])
        start = cut

    return parts

def pack_blocks(blocks: List[List[str]], cfg: SplitConfig) -> List[List[str]]:
    """
    Pack logical (sub)blocks into chunks whose total line count does not exceed
    `cfg.max_lines`.

    Parameters
    ----------
    blocks : List[List[str]]
        Blocks of SAS code, typically from `iter_logical_blocks`. Any oversized
        block should be pre-split with `split_oversized_block` before packing.
    cfg : SplitConfig
        Configuration controlling the maximum lines and trimming behavior.

    Returns
    -------
    List[List[str]]
        A list of chunks, each chunk a list of lines ready to be written.

    Implementation details
    ----------------------
    - Trims trailing blank lines if `keep_empty_trailing_lines` is False.
    """
    chunks: List[List[str]] = []
    current: List[str] = []

    def flush():
        nonlocal current
        if not cfg.keep_empty_trailing_lines:
            # Trim trailing empties for a tidy chunk
            while current and not current[-1].strip():
                current.pop()
        if current:
            chunks.append(current)
            current = []

    for block in blocks:
        # Ensure no single block exceeds max_lines
        subblocks = split_oversized_block(block, cfg)
        for sub in subblocks:
            if len(current) + len(sub) <= cfg.max_lines:
                current.extend(sub)
            else:
                flush()
                if len(sub) > cfg.max_lines:
                    # This shouldn’t happen thanks to split_oversized_block,
                    # but guard anyway.
                    logger.warning("Encountered a subblock > max_lines during packing; writing as-is.")
                current.extend(sub)

    flush()
    return chunks

# --------------------------------------------------------------------------------------
# High-level helpers (text and file I/O)
# --------------------------------------------------------------------------------------

def split_sas_text(sas_text: str, cfg: Optional[SplitConfig] = None) -> List[str]:
    """
    Split a SAS program (provided as a single string) into chunk strings.

    Parameters
    ----------
    sas_text : str
        Entire SAS program contents.
    cfg : Optional[SplitConfig]
        Configuration for splitting. If None, defaults are used.

    Returns
    -------
    List[str]
        Chunked SAS text segments as strings. Each element corresponds to one
        output block (≤ `max_lines`).

    Examples
    --------
    >>> chunks = split_sas_text(some_sas_text, SplitConfig(max_lines=200))
    >>> len(chunks)
    3

    Extensibility
    -------------
    - To support more SAS constructs, extend the regexes or the detection logic
      in `iter_logical_blocks`. Because we separate *detection* from *output*,
      changes there won’t risk altering user source code.
    """
    cfg = cfg or SplitConfig()
    lines = sas_text.splitlines(keepends=True)
    logical_blocks = iter_logical_blocks(lines)

    # Pre-split any oversized logical blocks, then pack
    flattened: List[List[str]] = []
    for b in logical_blocks:
        flattened.extend(split_oversized_block(b, cfg))

    chunks = pack_blocks(flattened, cfg)
    return ["".join(chunk) for chunk in chunks]

def split_sas_file(
    input_path: str | Path,
    out_dir: str | Path = "chunks",
    cfg: Optional[SplitConfig] = None,
) -> List[Path]:
    """
    Split a SAS file into multiple <= `max_lines` parts and write them to disk.

    Parameters
    ----------
    input_path : str | Path
        Path to the source `.sas` file.
    out_dir : str | Path
        Output directory. It will be created if it does not exist.
    cfg : Optional[SplitConfig]
        Configuration for splitting. If None, defaults are used.

    Returns
    -------
    List[Path]
        A list of file paths written in order, e.g.,
        `myfile_part_001.sas`, `myfile_part_002.sas`, ...

    Raises
    ------
    FileNotFoundError
        If `input_path` does not exist.
    IOError
        On read/write issues.

    Notes
    -----
    - This function is notebook-safe. In Databricks, `out_dir` can be a DBFS
      path (e.g., `/dbfs/FileStore/...`) if desired.
    """
    cfg = cfg or SplitConfig()
    input_path = Path(input_path)
    out_dir = Path(out_dir)

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    try:
        text = input_path.read_text(encoding="utf-8", errors="ignore")
    except Exception as e:
        logger.error("Failed to read input file: %s", e)
        raise

    chunks = split_sas_text(text, cfg)

    out_dir.mkdir(parents=True, exist_ok=True)
    prefix = cfg.filename_prefix or input_path.stem
    pad = max(3, len(str(len(chunks))))
    written: List[Path] = []

    for i, chunk in enumerate(chunks, start=1):
        p = out_dir / f"{prefix}_part_{i:0{pad}d}.sas"
        try:
            p.write_text(chunk, encoding="utf-8")
        except Exception as e:
            logger.error("Failed to write %s: %s", p, e)
            raise
        written.append(p)

    logger.info("Wrote %d chunk(s) to %s", len(written), out_dir)
    return written

# --------------------------------------------------------------------------------------
# Convenience: minimal CLI (optional in Databricks; safe to ignore)
# --------------------------------------------------------------------------------------

def _cli(argv: Optional[List[str]] = None) -> int:
    """
    Minimal command-line entry point for local runs (optional in notebooks).

    Parameters
    ----------
    argv : Optional[List[str]]
        Argument vector (without the program name). If None, `sys.argv[1:]` is used.

    Returns
    -------
    int
        Zero on success; non-zero on failure.
    """
    import argparse, sys
    ap = argparse.ArgumentParser(description="Split SAS script into <= N-line chunks.")
    ap.add_argument("input", help="Path to input .sas file")
    ap.add_argument("--out-dir", default="chunks", help="Output directory")
    ap.add_argument("--max-lines", type=int, default=200, help="Maximum lines per chunk")
    ap.add_argument("--no-prefer-blank", action="store_true", help="Do not prefer blank-line split points")
    ap.add_argument("--keep-empty-trailing-lines", action="store_true", help="Keep trailing blank lines in chunks")
    ap.add_argument("--prefix", default=None, help="Filename prefix (defaults to input stem)")
    args = ap.parse_args(argv)

    cfg = SplitConfig(
        max_lines=args.max_lines,
        prefer_blank_split=not args.no_prefer_blank,
        keep_empty_trailing_lines=args.keep_empty_trailing_lines,
        filename_prefix=args.prefix,
    )
    try:
        split_sas_file(args.input, args.out_dir, cfg)
        return 0
    except Exception as e:
        logger.error("Error: %s", e)
        return 1

# If you want to run this cell as a script locally, uncomment below.
# if __name__ == "__main__":
#     import sys
#     raise SystemExit(_cli(sys.argv[1:]))

# Example: splitting an in-memory SAS string
cfg = SplitConfig(max_lines=200)
chunks = split_sas_text(sas_text, cfg)
displayHTML(f"<pre>{chunks[0][:1000]}</pre>")  # peek first chunk

# Example: splitting a file stored in DBFS
# written_paths = split_sas_file("/dbfs/FileStore/my_project/my_code.sas", "/dbfs/FileStore/my_project/chunks", cfg)
