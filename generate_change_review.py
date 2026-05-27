#!/usr/bin/env python3
"""
Generate a Word document summarizing net differences between a historical git
commit and the current HEAD.

Requires: pip install python-docx

Usage:
    python3 generate_change_review.py [<commit>] [--output <name.docx>]

    <commit> defaults to the TARGET_COMMIT constant below.
    Edit the Configuration section to adjust thresholds, exclusions, etc.
"""

import subprocess
import os
import sys
import fnmatch
import re
import time

try:
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.enum.section import WD_ORIENT
    from docx.oxml.ns import qn
except ImportError:
    print("Error: python-docx not installed. Run: pip install python-docx")
    sys.exit(1)

# ── Configuration ────────────────────────────────────────────────────────────

TARGET_COMMIT = "CHANGE_ME"           # edit or pass as 1st CLI argument
OUTPUT_FILE = "change_review.docx"    # override with --output

# Path components to exclude (checked against every path segment).
# IMPORTANT: is_excluded() matches these patterns against EACH path segment
# individually.  Only use patterns that are unambiguous — directory names that
# are NEVER legitimate source code (e.g. "node_modules", "__pycache__") and
# file extensions that are NEVER source text (e.g. "*.so", "*.png").
# Do NOT add generic names like "dist", "build", "vendor", "test" — they will
# match identically named directories deep inside legitimate source trees.
EXCLUDE_PATTERNS = [
    "node_modules",
    "__pycache__",
    ".git",
    ".svn",
    ".hg",
    ".cache",
    "*.pyc",
    "*.pyo",
    "*.so",
    "*.o",
    "*.a",
    "*.dll",
    "*.exe",
    "*.bin",
    "*.zip",
    "*.tar",
    "*.gz",
    "*.bz2",
    "*.7z",
    "*.rar",
    "*.png",
    "*.jpg",
    "*.jpeg",
    "*.gif",
    "*.ico",
    "*.svg",
    "*.pdf",
    "*.ttf",
    "*.woff",
    "*.woff2",
    "*.eot",
    "*.mp3",
    "*.mp4",
    "*.avi",
    "*.mov",
    "*.wasm",
    "*.class",
    "*.jar",
]

# Path prefixes that are always included (regardless of MIN_CHANGED_LINES).
# Example: ["src/module/subdir", "docs/important"]
FORCE_INCLUDE_PREFIXES = []

# Document file exclusion — when SKIP_DOC_FILES is True, documentation-type files
# (markdown, .gitignore, license notices, etc.) are skipped and listed in an appendix.
SKIP_DOC_FILES = True
DOC_EXCLUDE_PATTERNS = [
    "*.md",
    "*.markdown",
    "*.MD",
    ".gitignore",
    "Third_Party_Open_Source_Software_Notice",
]

# Thresholds
MIN_CHANGED_LINES = 20       # ignore files with fewer total changes
FULL_OUTPUT_LINES = 200      # changed_lines > this  →  skip ratio-check, go straight to full output

# Word formatting
FONT_MONO = "Consolas"
FONT_SIZE_CODE = Pt(9)
PAGE_LANDSCAPE = False       # portrait (vertical) orientation

# ── Helpers ──────────────────────────────────────────────────────────────────

def run(cmd_args, cwd=None):
    """Run a command (list of args), return stdout or raise on failure."""
    if cwd is None:
        cwd = os.getcwd()
    # Force proper UTF-8 path output (git octal-escapes non-ASCII paths by default)
    if cmd_args[0] == "git":
        cmd_args = ["git", "-c", "core.quotePath=false"] + cmd_args[1:]
    result = subprocess.run(cmd_args, capture_output=True, text=True, cwd=cwd)
    if result.returncode != 0:
        raise RuntimeError(f"Command failed: {' '.join(cmd_args)}\n{result.stderr}")
    return result.stdout


def is_excluded(filepath):
    """True if any path segment matches an exclude pattern."""
    # Force-included paths are never excluded
    if is_force_include(filepath):
        return False
    filepath = filepath.replace("\\", "/")
    segments = filepath.split("/")
    for seg in segments:
        for pat in EXCLUDE_PATTERNS:
            if fnmatch.fnmatch(seg, pat):
                return True
    # Also match full path against patterns containing "/"
    for pat in EXCLUDE_PATTERNS:
        if "/" in pat and fnmatch.fnmatch(filepath, pat):
            return True
    return False


def is_force_include(filepath):
    """True if filepath starts with any force-include prefix."""
    filepath = filepath.replace("\\", "/")
    for prefix in FORCE_INCLUDE_PREFIXES:
        if filepath.startswith(prefix):
            return True
    return False


def is_doc_file(filepath):
    """True if filepath matches a document-type exclusion pattern (when SKIP_DOC_FILES is enabled)."""
    if not SKIP_DOC_FILES:
        return False
    if is_force_include(filepath):
        return False
    filepath = filepath.replace("\\", "/")
    filename = filepath.split("/")[-1] if "/" in filepath else filepath
    for pat in DOC_EXCLUDE_PATTERNS:
        if fnmatch.fnmatch(filename, pat):
            return True
    return False


def parse_name_status(output):
    """Parse 'git diff --name-status' output.
    Returns dict: new_path → {"status": "A"|"M"|"R"|"C"|..., "old_path": str|None}
    """
    statuses = {}
    for line in output.strip().split("\n"):
        if not line:
            continue
        parts = line.split("\t")
        code = parts[0]
        if code.startswith("R"):
            # R100, R050, etc.
            old_p, new_p = parts[1], parts[2]
            statuses[new_p] = {"status": "R", "old_path": old_p}
        elif code.startswith("C"):
            old_p, new_p = parts[1], parts[2]
            statuses[new_p] = {"status": "C", "old_path": old_p}
        else:
            # A, M, D, T, etc.
            filepath = parts[1]
            statuses[filepath] = {"status": code, "old_path": None}
    return statuses


def parse_numstat(output):
    """Parse 'git diff --numstat' output, handling {old => new} rename syntax.
    Returns dict: filepath → {"added": int, "deleted": int, "binary": bool, "old_path": str|None}
    """
    stats = {}
    for line in output.strip().split("\n"):
        if not line:
            continue
        parts = line.split("\t")
        if len(parts) < 3:
            continue
        added, deleted = parts[0], parts[1]
        raw_path = "\t".join(parts[2:])

        # Resolve {old => new} rename notation.
        # Git formats (non-greedy to handle prefix/{...} and { => new}):
        #   {old => new}        → simple rename
        #   {old => new}/suffix → dir rename with common suffix
        #   prefix/{old => new} → rename inside a subdirectory
        #   { => new}/suffix    → new file (empty old part)
        old_path = None
        m = re.match(r'^(.*)\{(.*?) => (.*?)\}(.*)$', raw_path)
        if m:
            prefix = m.group(1)
            old_part = m.group(2).strip()
            new_part = m.group(3).strip()
            suffix = m.group(4)
            filepath = (prefix + new_part + suffix).replace('//', '/')
            if old_part:
                old_path = (prefix + old_part + suffix).replace('//', '/')
            # else: empty old_part → file is effectively new, old_path stays None
        else:
            filepath = raw_path

        if added == "-" and deleted == "-":
            stats[filepath] = {"added": 0, "deleted": 0, "binary": True, "old_path": old_path}
        else:
            stats[filepath] = {
                "added": int(added) if added.lstrip("-").isdigit() else 0,
                "deleted": int(deleted) if deleted.lstrip("-").isdigit() else 0,
                "binary": False,
                "old_path": old_path,
            }
    return stats


def read_head_file(filepath):
    """Return (content, error_msg) for a file at HEAD."""
    try:
        out = run(["git", "show", f"HEAD:{filepath}"])
        return out, None
    except Exception as e:
        return None, str(e)


# ── Word helpers ─────────────────────────────────────────────────────────────

def setup_page(doc):
    """Configure page size and margins."""
    for section in doc.sections:
        if PAGE_LANDSCAPE:
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width = Cm(29.7)
            section.page_height = Cm(21.0)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)


def set_mono(run):
    """Apply monospace font + size to a run."""
    run.font.name = FONT_MONO
    run.font.size = FONT_SIZE_CODE
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("w:rFonts"), {})
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), FONT_MONO)
    rFonts.set(qn("w:cs"), FONT_MONO)


def add_heading_text(doc, text):
    """Add a Heading-1 style paragraph."""
    h = doc.add_heading(text, level=1)
    for run in h.runs:
        run.font.name = "Calibri"


def add_code_block(doc, code_text):
    """Append code_text in a single paragraph with line-break-separated monospace runs."""
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    para.paragraph_format.line_spacing = 1.0

    lines = code_text.split("\n")
    for i, line in enumerate(lines):
        # Strip control characters that are invalid in XML (keep tab, newline, carriage return)
        clean_line = "".join(ch for ch in (line if line else " ") if ch == "\t" or ch == "\r" or (ord(ch) >= 32 and ord(ch) != 0x7F) or ch in "\n")
        run = para.add_run(clean_line if clean_line else " ")
        set_mono(run)
        if i < len(lines) - 1:
            run.add_break(WD_BREAK.LINE)


def add_normal_para(doc, text, size=Pt(10), italic=False):
    """Add a normal paragraph with Calibri font."""
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = size
    r.italic = italic


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    t_start = time.time()

    print(f"Target commit : {TARGET_COMMIT}")
    print(f"Range         : {TARGET_COMMIT}^ .. HEAD")
    print()

    # 1. Collect raw metadata from git
    print("[1/3] Fetching file statuses (git diff --name-status) ...", flush=True)
    t0 = time.time()
    statuses_raw = run(["git", "diff", "--name-status", f"{TARGET_COMMIT}^..HEAD"])
    statuses = parse_name_status(statuses_raw)
    print(f"      → {len(statuses)} files  ({time.time() - t0:.1f}s)", flush=True)

    print("[2/3] Fetching line counts (git diff --numstat) ...", flush=True)
    t0 = time.time()
    numstats_raw = run(["git", "diff", "--numstat", f"{TARGET_COMMIT}^..HEAD"])
    numstats = parse_numstat(numstats_raw)
    print(f"      → {len(numstats)} entries  ({time.time() - t0:.1f}s)", flush=True)

    # 2. Classify every file
    print("[3/3] Classifying files ...", flush=True)
    t0 = time.time()
    files_to_process = []      # (filepath, info_dict)
    skipped_binary = []        # filepath
    skipped_docs = []          # filepath
    encoding_issues = []       # (filepath, error_msg)

    entries = sorted(numstats.items())
    total_entries = len(entries)
    force_count = 0
    last_report = 0

    for idx, (filepath, ns) in enumerate(entries):
        # ── Progress pulse: print every 500 files or on significant events ──
        if idx - last_report >= 500 or idx == total_entries - 1:
            elapsed = time.time() - t0
            sys.stdout.write(
                f"\r      scanning {idx + 1}/{total_entries} entries  "
                f"(kept {len(files_to_process)}, "
                f"binary {len(skipped_binary)}, "
                f"docs {len(skipped_docs)}, "
                f"force {force_count})  [{elapsed:.0f}s]"
            )
            sys.stdout.flush()
            last_report = idx

        # ── Exclude check ──
        if is_excluded(filepath):
            continue

        # ── Doc file check ──
        if is_doc_file(filepath):
            skipped_docs.append(filepath)
            continue

        # ── Deleted files: skip ──
        st = statuses.get(filepath, {})
        if st.get("status") == "D":
            continue

        # ── Binary ──
        if ns["binary"]:
            skipped_binary.append(filepath)
            continue

        changed = ns["added"] + ns["deleted"]
        force = is_force_include(filepath)
        if not force and changed < MIN_CHANGED_LINES:
            continue

        if force:
            force_count += 1

        st = statuses.get(filepath, {"status": "M", "old_path": None})

        # Decide output mode — defer content reads to Word-gen phase to avoid
        # reading every file twice (once here for line-count, once for output).
        if st["status"] in ("A", "R", "C"):
            # New / renamed / copied → always full, added count ≈ total lines
            output_mode = "full"
            total = ns["added"]  # exact for new files, close for renames
        elif changed > FULL_OUTPUT_LINES:
            output_mode = "full"
            total = 0  # will be counted from content
        else:
            # Need ratio check → read content once during Word gen
            output_mode = "check_ratio"
            total = 0

        files_to_process.append((filepath, {
            **ns,
            "status": st["status"],
            "changed_lines": changed,
            "total_lines": total,
            "output_mode": output_mode,
        }))

    print()  # newline after progress line
    print(f"      → {len(files_to_process)} files selected  "
          f"({time.time() - t0:.1f}s)", flush=True)

    # 3. Build Word document
    print(f"\nBuilding '{OUTPUT_FILE}' ...", flush=True)
    t0 = time.time()
    doc = Document()
    setup_page(doc)

    # ── Title ──
    # Detect repository name from the current working directory
    repo_name = os.path.basename(os.getcwd().rstrip("/"))
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = tp.add_run(f"{repo_name} 项目代码整合文档")
    r.bold = True
    r.font.name = "Calibri"
    r.font.size = Pt(16)

    # ── Per-file pages ──
    total_selected = len(files_to_process)
    for idx, (fp, info) in enumerate(files_to_process):
        mode = info["output_mode"]
        cached_content = None

        # Always output full file content (no diff mode).
        if mode == "check_ratio":
            cached_content, err = read_head_file(fp)
            if err:
                encoding_issues.append((fp, err))
                continue
            info["total_lines"] = cached_content.count("\n")
            mode = info["output_mode"] = "full"

        sys.stdout.write(
            f"\r  [{idx + 1}/{total_selected}] {fp}  "
            f"({info['status']}, {info['changed_lines']} lines, → {mode})"
        )
        sys.stdout.flush()

        if idx > 0:
            doc.add_page_break()

        status_label = {
            "A": "[NEW]", "R": "[RENAMED]", "C": "[COPIED]",
            "M": "[MODIFIED]", "T": "[TYPE CHANGED]",
        }.get(info["status"], "")

        add_heading_text(doc, f"{fp}  {status_label}  (+{info['added']}/-{info['deleted']})")

        meta = doc.add_paragraph()
        r = meta.add_run(
            f"Status: {info['status']}  |  "
            f"Changed: {info['changed_lines']} lines  |  "
            f"Total (HEAD): {info['total_lines']} lines  |  "
            f"Mode: {info['output_mode']}"
        )
        r.font.name = "Calibri"
        r.font.size = Pt(8)
        r.italic = True

        if cached_content is not None:
            content = cached_content
        else:
            content, err = read_head_file(fp)
            if err:
                encoding_issues.append((fp, err))
                add_normal_para(doc, f"[SKIPPED] Could not read file: {err}", size=Pt(9))
                continue
            # Back-fill total_lines for files that skipped counting
            if info["total_lines"] == 0:
                info["total_lines"] = content.count("\n")
        add_code_block(doc, content)

    print()  # newline after progress line
    print(f"      → document body built  ({time.time() - t0:.1f}s)", flush=True)

    # ── Appendices ──

    # Binary files
    if skipped_binary:
        doc.add_page_break()
        add_heading_text(doc, "Appendix: Skipped Binary Files")
        add_normal_para(doc, "The following binary files were excluded from the diff review:", size=Pt(10))
        for bf in sorted(skipped_binary):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            mr = p.add_run(f"  {bf}")
            set_mono(mr)

    # Doc files
    if skipped_docs:
        doc.add_page_break()
        add_heading_text(doc, "Appendix: Skipped Document Files")
        add_normal_para(doc, "The following document-type files were excluded from the review "
                        "(SKIP_DOC_FILES is enabled):", size=Pt(10))
        for dfp in sorted(skipped_docs):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            mr = p.add_run(f"  {dfp}")
            set_mono(mr)

    # Encoding issues
    if encoding_issues:
        doc.add_page_break()
        add_heading_text(doc, "Appendix: Encoding / Read Issues")
        add_normal_para(doc, "The following files could not be read as UTF-8 text:", size=Pt(10))
        for fp, err in sorted(encoding_issues):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            mr = p.add_run(f"  {fp}\n    → {err}")
            set_mono(mr)

    # Save
    print("Saving document ...", flush=True)
    doc.save(OUTPUT_FILE)

    total_elapsed = time.time() - t_start
    print(f"\nDone → {OUTPUT_FILE}  (total {total_elapsed:.1f}s)")
    print(f"  Files written to doc : {len(files_to_process) - len(encoding_issues)}")
    print(f"  Binary skipped       : {len(skipped_binary)}")
    print(f"  Doc files skipped    : {len(skipped_docs)}")
    print(f"  Encoding issues      : {len(encoding_issues)}")


def parse_args():
    """Minimal CLI: python3 generate_change_review.py [<commit>] [--output <file>]"""
    global TARGET_COMMIT, OUTPUT_FILE
    args = sys.argv[1:]
    i = 0
    while i < len(args):
        if args[i] in ("-h", "--help"):
            print("Usage: python3 generate_change_review.py [<commit>] [--output <name.docx>]")
            print()
            print("  <commit>       Target commit hash (diff from its parent to HEAD).")
            print("  --output FILE  Output filename (default: change_review.docx).")
            print()
            print("Configuration constants (edit in script):")
            print("  EXCLUDE_PATTERNS, FORCE_INCLUDE_PREFIXES, SKIP_DOC_FILES,")
            print("  DOC_EXCLUDE_PATTERNS, MIN_CHANGED_LINES, FULL_OUTPUT_LINES, PAGE_LANDSCAPE")
            sys.exit(0)
        elif args[i] == "--output" and i + 1 < len(args):
            OUTPUT_FILE = args[i + 1]
            i += 2
        elif not args[i].startswith("-"):
            TARGET_COMMIT = args[i]
            i += 1
        else:
            print(f"Unknown flag: {args[i]}")
            print("Usage: python3 generate_change_review.py [<commit>] [--output <name.docx>]")
            sys.exit(1)


if __name__ == "__main__":
    parse_args()
    if TARGET_COMMIT == "CHANGE_ME":
        print("Error: TARGET_COMMIT not set.  Edit the script or pass a commit hash on the command line.")
        print("Usage: python3 generate_change_review.py <commit> [--output <name.docx>]")
        sys.exit(1)
    main()
