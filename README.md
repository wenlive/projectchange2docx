# git-diff-review

生成 Git 仓库两个 commit 之间净差异的 Word 审查文档。

## 安装

```bash
pip install python-docx
```

## 使用

**重要**：脚本必须在目标仓库根目录执行（`git diff` 使用当前工作目录）。

### 方式一：CLI 传参（推荐，无需编辑脚本）

```bash
cd /path/to/target-repo
python3 /path/to/git-diff-review/generate_change_review.py <commit-hash>
```

可选指定输出文件名：

```bash
python3 generate_change_review.py abc123 --output my_review.docx
```

### 方式二：编辑脚本常量

1. 将 `generate_change_review.py` 复制到目标仓库根目录。
2. 编辑 `TARGET_COMMIT` 和其他常量。
3. 运行 `python3 generate_change_review.py`。

### 可调常量

| 常量 | 默认 | 说明 |
|------|------|------|
| `TARGET_COMMIT` | `"CHANGE_ME"` | 历史起点 commit hash |
| `MIN_CHANGED_LINES` | 20 | 变更低于此行数的文件忽略 |
| `FULL_OUTPUT_LINES` | 200 | 变更超过此行数 → 跳过比例检查直接全文输出 |
| `SKIP_DOC_FILES` | `True` | 是否跳过文档类型文件（开关） |
| `DOC_EXCLUDE_PATTERNS` | `[*.md, .gitignore, ...]` | 文档排除模式（仅文件名匹配） |
| `EXCLUDE_PATTERNS` | `[node_modules, __pycache__, ...]` | 排除的路径段/扩展名 |
| `FORCE_INCLUDE_PREFIXES` | `[]` | 强制包含的路径前缀 |
| `PAGE_LANDSCAPE` | `False` | 纵向/横向 |
| `FILEPATH_EXCLUDE_KEYWORDS` | `[opengauss, gaussdb, ...]` | 路径命中即排除文件（不区分大小写） |
| `PATH_REWRITE_RULES` | `[(gausskernel, kernel), ...]` | 路径组件重写规则 |
| `CONTENT_REPLACEMENTS` | `[(openGauss, HelmDB), ...]` | 代码内容中的品牌词替换 |

**注意**：`EXCLUDE_PATTERNS` 使用路径段匹配（逐段比对），请只添加无歧义的模式。不要添加 `dist`、`build`、`vendor` 等通用目录名——它们会匹配到深层源码中同名目录，导致文件被静默丢弃。`DOC_EXCLUDE_PATTERNS` 仅匹配文件名（最后一段），可以放心使用 `*.md` 等扩展名模式。

## 品牌过滤

脚本支持从输出中移除特定品牌的痕迹，适用于项目 fork/rebase 场景：

- **`FILEPATH_EXCLUDE_KEYWORDS`**：路径（不区分大小写）包含任一关键词的文件直接不导出，列入 Brand-Excluded 附录
- **`PATH_REWRITE_RULES`**：Word 文档中显示的所有路径会应用这些重写规则（原始路径仍用于 git 读取）
- **`CONTENT_REPLACEMENTS`**：文件内容在写入 Word 前按顺序应用这些替换（区分大小写）

路径重写和内容替换也适用于所有附录（Binary、Doc、Brand、Encoding）中的文件路径。

## 回归验证

`verify_change_review.py` 可对生成的 Word 文档进行交叉比对：

```bash
python3 verify_change_review.py <commit> <docx> --sample 10
```

验证内容：
- 文件分类计数是否与 git diff 规则一致
- 包含/排除附录是否无遗漏、无多余
- 随机采样文件内容是否匹配（含归一化处理）
- 全文档品牌词扫描

## 输出说明

- **全文模式**：所有纳入文件统一输出完整内容（当前 HEAD 版本），不再使用 unified diff
- **文档筛选**：默认跳过 `.md`、`.gitignore`、`Third_Party_Open_Source_Software_Notice` 等文档类文件，可关闭
- **文档附录**：被跳过的文档文件清单
- **二进制附录**：跳过的二进制文件清单
- **编码异常附录**：无法以 UTF-8 读取的文件清单
- Word 标题自动检测为 `{仓库名} 项目代码整合文档`
- 每文件从新的一页开始，代码使用 Consolas 9pt 等宽字体
- 文件按仓库层级排序（`sorted()` 路径序）
- 统计信息仅在终端输出，不出现在 Word 文档中
