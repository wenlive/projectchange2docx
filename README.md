# git-diff-review

生成 Git 仓库两个 commit 之间净差异的 Word 审查文档。

## 安装

```bash
pip install python-docx
```

## 使用

1. 将 `generate_change_review.py` 复制到目标仓库根目录。
2. 编辑脚本头部的配置常量（至少修改 `TARGET_COMMIT`）：

```python
TARGET_COMMIT = "<your-commit-hash>"
```

3. 可选调整：
   - `EXCLUDE_PATTERNS` — 需要排除的目录/文件扩展名
   - `FORCE_INCLUDE_PREFIXES` — 无论如何都要包含的路径前缀
   - `MIN_CHANGED_LINES` / `FULL_OUTPUT_LINES` / `FULL_OUTPUT_RATIO` — 阈值
   - `PAGE_LANDSCAPE` — 纵向/横向

4. 运行：

```bash
python3 generate_change_review.py
```

5. 打开生成的 `change_review.docx`。

## 输出说明

- **全文模式**（649/709）：新增文件、重命名文件、或修改超过 200 行/30% 的文件
- **Diff 模式**（60/709）：修改量较小的文件，仅展示 unified diff
- **二进制附录**：跳过的二进制文件清单
- 每文件从新的一页开始，代码使用 Consolas 9pt 等宽字体
- 文件按仓库层级排序（`sorted()` 路径序）
