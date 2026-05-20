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
| `FULL_OUTPUT_LINES` | 200 | 变更超过此行数 → 全文输出 |
| `FULL_OUTPUT_RATIO` | 0.30 | 变更超过此比例 → 全文输出 |
| `EXCLUDE_PATTERNS` | `[node_modules, dist, ...]` | 排除的路径/扩展名 |
| `FORCE_INCLUDE_PREFIXES` | `[]` | 强制包含的路径前缀 |
| `PAGE_LANDSCAPE` | `False` | 纵向/横向 |

## 输出说明

- **全文模式**：新增文件、重命名文件、或修改超过 200 行/30% 的文件
- **Diff 模式**：修改量较小的文件，仅展示 unified diff（`@@ -x,y +a,b @@`）
- **二进制附录**：跳过的二进制文件清单
- **编码异常附录**：无法以 UTF-8 读取的文件清单
- 每文件从新的一页开始，代码使用 Consolas 9pt 等宽字体
- 文件按仓库层级排序（`sorted()` 路径序）
