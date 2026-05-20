# git-diff-review

Portable workflow for generating a Word document that summarizes net git
differences between two commits.  Ships as a Claude Code skill so the same
process can be reused across repositories.

## Quick start

```bash
# From the target repository root:
pip install python-docx
python3 /path/to/git-diff-review/generate_change_review.py <commit-hash>

# Or copy the script in and edit TARGET_COMMIT/TARGET_COMMIT in the header,
# then run without args:
python3 generate_change_review.py
```

The script must be invoked from the target repository root — all `git diff`
commands use `os.getcwd()`.

## File layout

```
generate_change_review.py   – The main script (Python 3 + python-docx)
.claude/skills/
  git-diff-review.md        – Claude Code skill definition
CLAUDE.md                   – (this file) Project-level context
README.md                   – Human-readable documentation
```

## Skill behaviour

When the `git-diff-review` skill is invoked, Claude will:

1. Confirm the target commit and any force-include / exclude paths.
2. Run `generate_change_review.py` (after adjusting `TARGET_COMMIT` in the
   temporary copy).
3. Verify the output — headings count, mode distribution, binary appendix.
4. Report results.

## Key design decisions

- **No intermediate-history analysis** — pure `git diff A..B`, not commit-by-commit.
- **Force-include prefix** — specific subtrees (e.g. `src/vendor/foo`) can be
  marked to always appear regardless of the 20-line threshold.
- **Deferred content reads** — file contents are read during Word generation,
  not during classification, to avoid double I/O.
- **Single-paragraph-per-file** — code blocks use `<w:br/>` line breaks inside
  one Word paragraph, cutting python-docx XML overhead ~100×.
