---
name: git-diff-review
description: >
  Given two git commits, generate a Word document summarizing the net
  differences.  Binary files are skipped (listed in appendix), deleted files
  are skipped, renamed files are treated as new, and small changes (< 20
  lines, configurable) are omitted.  Large changes (> 200 lines or > 30 %
  of the file) get full current-version content; otherwise only the unified
  diff is included.  Every file starts on a new page, code is rendered in
  Consolas 9 pt, and a force-include prefix list keeps important paths
  regardless of the change-size threshold.
---

## 概述

此 skill 用于将两个 git commit（或一个 commit 与 HEAD）之间的最终净差异整理
成 Word 文档。它绕过了对中间历史的分析，只关心两端点的最终状态差异。

## 何时触发

当用户要求：
- "对比 commit X 和 HEAD 的差异，生成 Word 文档"
- "审查某个 commit 以来的所有变更"
- "生成 change review 文档"
- "把 git diff 输出整理成 Word"

## 工作流

### 1. 确认参数

在运行脚本之前，确认以下参数：

- **TARGET_COMMIT**：要比较的历史 commit hash。
- **输出文件名**：默认为 `change_review.docx`，放在仓库根目录。
- **排除目录/文件模式**：脚本头部的 `EXCLUDE_PATTERNS` 列表。
- **强制包含前缀**：`FORCE_INCLUDE_PREFIXES` 列表，其中的路径即使变更低于
  最小行数阈值也会被包含（如 `src/gausskernel/storage/nvmdb`）。

### 2. 运行脚本

```bash
pip install python-docx
python3 generate_change_review.py
```

脚本将从 `.py` 文件所在目录运行 `git diff`。请在仓库根目录执行。

### 3. 验证输出

脚本完成后，检查终端输出中的关键统计信息：
- `Files written to doc` — 写入文档的文件数
- `Binary skipped` — 跳过的二进制文件数
- `Encoding issues` — 编码问题（应为 0）

## 脚本配置常量

这些常量在脚本头部，可按需调整：

```python
TARGET_COMMIT = "<commit-hash>"          # 历史起点
OUTPUT_FILE = "change_review.docx"       # 输出文件名

EXCLUDE_PATTERNS = [...]                 # 排除的路径/扩展名

FORCE_INCLUDE_PREFIXES = [               # 强制包含的路径前缀
    "src/gausskernel/storage/nvmdb",
]

MIN_CHANGED_LINES = 20                   # 忽略变更少于此值的文件
FULL_OUTPUT_LINES = 200                  # 变更超过此行数 -> 输出全文
FULL_OUTPUT_RATIO = 0.30                 # 变更超过此比例 -> 输出全文
DIFF_CONTEXT_LINES = 3                   # unified diff 上下文行数

FONT_MONO = "Consolas"                   # 代码等宽字体
FONT_SIZE_CODE = Pt(9)                   # 代码字号
PAGE_LANDSCAPE = False                   # True=横向, False=纵向
```

## 已知边界情况 & 注意事项

### 中文文件名

Git 默认 `core.quotePath=true`，会将非 ASCII 路径输出为八进制转义序列。
脚本在 `run()` 函数中自动注入 `-c core.quotePath=false` 以正确读取 UTF-8
路径。

### git diff --numstat 中的重命名格式

`--numstat` 对重命名文件使用 `{old => new}` 简写，可能包含空的老路径部分
（如 `{ => new_dir}/file`）或嵌套在子目录中（如 `prefix/{old => new}/suffix`）。
脚本使用正则 `/^(.*)\{(.*?) => (.*?)\}(.*)$/` 处理这些情况。

### 性能考虑

- 脚本将文件的完整内容一次性读取到内存中，然后逐行写入 Word 文档。
  对于超过 10,000 行的文件，可能会占用大量内存。
- 分类阶段不进行文件读取 — 所有内容读取推迟到 Word 生成阶段，避免双重读取。
- 每个文件的代码内容使用单个 Word 段落 + 换行符（`<w:br/>`）而非每行一个
  段落。这显著降低了 python-docx 的 XML 操作开销。

### 文件过滤的优先级

1. 排除模式匹配（`is_excluded()`） → 跳过（但强制包含的前缀优先）
2. 已删除文件（`--name-status` 中的 `D`） → 跳过
3. 二进制文件（`--numstat` 中的 `- -`） → 跳过，计入附录
4. 改变行数 < `MIN_CHANGED_LINES` → 跳过（但强制包含的前缀优先）
5. 其余文件 → 根据阈值判断输出模式
