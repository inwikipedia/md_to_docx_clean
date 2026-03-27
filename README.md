
"""
Markdown 转 Word 工具（增强版）
--------------------------------
功能：
1. 自动清洗 ChatGPT 生成的 Markdown 数学公式
2. 调用 Pandoc 转换为 docx
3. 尽量输出 Word 原生公式

默认输入文件：source.md
默认输出文件：source.docx

依赖：
- Python 3.9+
- Pandoc 已安装，并可在命令行中直接调用

用法：
python md_to_docx_clean.py
或
python md_to_docx_clean.py input.md output.docx
"""
