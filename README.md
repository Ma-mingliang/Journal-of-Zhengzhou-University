# 郑州大学工学报格式转换工具

基于 python-docx 将任意 Word 文档重新格式化为郑州大学工学报投稿格式。

## 功能特性

- 自动识别源文档结构（标题、摘要、正文、参考文献）
- 中英文字体规范（宋体/黑体 + Times New Roman）
- 单栏摘要 + 双栏正文 + 双栏参考文献 + 单栏英文摘要的分节结构
- Heading 1/2 层级标题，段前段后 0 磅
- 封面页自动拼接（文件名含"封面"的 docx）
- 双栏布局中全宽图片插入（基于宽高比自动判断）
- 封面图片 rId 修复（WMF→PNG 自动转换）
- OMML 公式插入（通过 pandoc）

## 依赖

```bash
pip install python-docx lxml Pillow
# 可选：从 PDF 提取图片
pip install PyMuPDF
# 可选：OMML 公式
conda install -c conda-forge pandoc
```

## 使用方法

### 作为 Claude Code 技能使用

将 `SKILL.md` 复制到 `~/.claude/skills/zzu-journal-format/` 目录下，Claude Code 会自动加载该技能。

### 作为 Python 脚本使用

```bash
python templates/generate_template.py
```

或参考 `examples/` 目录下的完整示例。

## 文件结构

```
├── README.md                          # 本文件
├── SKILL.md                           # Claude Code 技能定义（格式规范 + 代码模板）
├── templates/
│   └── generate_template.py           # 通用脚本模板
└── examples/
    └── generate_translation_v3.py     # 完整示例（TD-MPC2 论文格式化）
```

## 格式规范

详见 [SKILL.md](SKILL.md)。

## 版本

- 当前版本：1.3.0
- 兼容平台：Claude Code (CLI/Desktop/Web)
- Python 环境：Anaconda
