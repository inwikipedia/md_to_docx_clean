#!/usr/bin/env python
# -*- coding: utf-8 -*-

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

import os
import re
import sys
import shutil
import subprocess
from pathlib import Path


def check_pandoc_installed():
    """检查 Pandoc 是否已安装"""
    pandoc_path = shutil.which("pandoc")
    if not pandoc_path:
        return False, "未找到 pandoc，请先安装并加入 PATH。"

    try:
        result = subprocess.run(
            ["pandoc", "--version"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore"
        )
        if result.returncode == 0:
            version_line = result.stdout.splitlines()[0] if result.stdout else "Pandoc"
            return True, version_line
        return False, result.stderr.strip() or "Pandoc 调用失败"
    except Exception as e:
        return False, str(e)


def read_text_file(file_path: Path) -> str:
    """兼容多编码读取文本"""
    encodings = ["utf-8", "utf-8-sig", "gbk", "gb18030"]
    last_err = None

    for enc in encodings:
        try:
            return file_path.read_text(encoding=enc)
        except Exception as e:
            last_err = e

    raise RuntimeError(f"读取文件失败: {file_path}\n原因: {last_err}")


def write_text_file(file_path: Path, content: str):
    """写入 UTF-8 文件"""
    file_path.write_text(content, encoding="utf-8", newline="\n")


def normalize_line_endings(text: str) -> str:
    """统一换行"""
    return text.replace("\r\n", "\n").replace("\r", "\n")


def replace_display_math_brackets(text: str) -> str:
    """
    将这种形式的块公式：
    [
    x = y
    ]
    转成：
    $$
    x = y
    $$

    仅处理“独占一行”的 [ 和 ]，避免误伤普通文本中的方括号。
    """
    pattern = re.compile(
        r'(?m)^[ \t]*\[[ \t]*\n(.*?)(?<=\n)[ \t]*\][ \t]*$',
        re.DOTALL
    )

    def repl(match):
        body = match.group(1).strip("\n")
        return f"$$\n{body}\n$$"

    prev = None
    while prev != text:
        prev = text
        text = pattern.sub(repl, text)

    return text


def replace_latex_display_math(text: str) -> str:
    """
    将 \[ ... \] 统一转为 $$ ... $$
    """
    pattern = re.compile(r'\\\[(.*?)\\\]', re.DOTALL)

    def repl(match):
        body = match.group(1).strip()
        return f"$$\n{body}\n$$"

    return pattern.sub(repl, text)


def replace_latex_inline_math(text: str) -> str:
    """
    将 \( ... \) 转为 $ ... $
    """
    pattern = re.compile(r'\\\((.*?)\\\)', re.DOTALL)

    def repl(match):
        body = match.group(1).strip()
        return f"${body}$"

    return pattern.sub(repl, text)


def repair_star_subscripts(text: str) -> str:
    r"""
    修复被写坏的下标：
    1) \bar{x}*{i,w}   -> \bar{x}_{i,w}
    2) \hat{x}*i       -> \hat{x}_i
    3) r*{ij}          -> r_{ij}
    4) \sum*{m=1}^{M}  -> \sum_{m=1}^{M}

    只在明显是公式 token 的位置修复，尽量减少误伤普通文本中的星号。
    """

    # 情况1：something*{...} -> something_{...}
    # 例如 \bar{x}*{i,w} / r*{ij} / \sum*{m=1}^{M}
    text = re.sub(
        r'(?<!\w)(\\?[A-Za-z]+(?:\{[^{}]+\})?|\\[A-Za-z]+|[A-Za-z])\*\{([^{}\n]+)\}',
        r'\1_{\2}',
        text
    )

    # 情况2：something*i -> something_i
    # 例如 \hat{x}*i / w*m / r*ij（只取一个 token）
    text = re.sub(
        r'(?<!\w)(\\?[A-Za-z]+(?:\{[^{}]+\})?|\\[A-Za-z]+|[A-Za-z])\*([A-Za-z0-9]+)',
        r'\1_\2',
        text
    )

    return text


def ensure_blank_lines_around_display_math(text: str) -> str:
    """
    保证 $$ 块公式前后有空行，避免 Pandoc 在列表或段落里误判。
    """
    lines = text.split("\n")
    out = []

    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if stripped == "$$":
            if out and out[-1].strip() != "":
                out.append("")

            out.append(line)
            i += 1

            while i < len(lines):
                out.append(lines[i])
                if lines[i].strip() == "$$":
                    break
                i += 1

            if i + 1 < len(lines) and lines[i + 1].strip() != "":
                out.append("")
            i += 1
            continue

        out.append(line)
        i += 1

    return "\n".join(out)


def cleanup_excess_blank_lines(text: str) -> str:
    """压缩过多空行"""
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip() + "\n"


def clean_markdown_math(text: str) -> str:
    """总清洗流程"""
    text = normalize_line_endings(text)
    text = replace_display_math_brackets(text)
    text = replace_latex_display_math(text)
    text = replace_latex_inline_math(text)
    text = repair_star_subscripts(text)
    text = ensure_blank_lines_around_display_math(text)
    text = cleanup_excess_blank_lines(text)
    return text


def convert_markdown_to_docx(input_file: Path, output_file: Path):
    """
    使用 Pandoc 转换 Markdown 到 docx
    显式启用 tex_math_dollars，提升数学公式识别稳定性
    """
    cmd = [
        "pandoc",
        str(input_file),
        "-f", "markdown+tex_math_dollars",
        "-t", "docx",
        "-o", str(output_file),
        "--standalone",
        "--toc"
    ]

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore"
        )
        if result.returncode == 0:
            return True, f"转换成功：{output_file}"
        return False, result.stderr.strip() or "Pandoc 转换失败"
    except Exception as e:
        return False, f"执行 Pandoc 出错：{e}"


def main():
    print("=" * 70)
    print("Markdown -> Word 转换工具（公式清洗增强版）")
    print("=" * 70)

    ok, info = check_pandoc_installed()
    if not ok:
        print(f"\n❌ Pandoc 检查失败：{info}")
        print("\n请先安装 Pandoc：https://pandoc.org/installing.html")
        sys.exit(1)

    print(f"\n✓ 检测到 {info}")

    # 参数
    if len(sys.argv) >= 2:
        input_path = Path(sys.argv[1]).resolve()
    else:
        input_path = (Path.cwd() / "source.md").resolve()

    if len(sys.argv) >= 3:
        output_path = Path(sys.argv[2]).resolve()
    else:
        output_path = input_path.with_suffix(".docx")

    cleaned_md_path = input_path.with_name(input_path.stem + "_cleaned.md")

    print(f"\n输入文件：{input_path}")
    print(f"清洗文件：{cleaned_md_path}")
    print(f"输出文件：{output_path}")

    if not input_path.exists():
        print(f"\n❌ 输入文件不存在：{input_path}")
        sys.exit(1)

    try:
        print("\n[1/3] 读取 Markdown ...")
        original_text = read_text_file(input_path)

        print("[2/3] 清洗公式格式 ...")
        cleaned_text = clean_markdown_math(original_text)
        write_text_file(cleaned_md_path, cleaned_text)

        print("[3/3] 调用 Pandoc 转 Word ...")
        success, msg = convert_markdown_to_docx(cleaned_md_path, output_path)

        if success:
            print(f"\n✅ {msg}")
            print(f"✅ 已生成清洗后的 Markdown：{cleaned_md_path}")
            print("✅ 建议优先检查 cleaned.md 中个别复杂公式是否仍需人工微调。")
        else:
            print(f"\n❌ {msg}")
            sys.exit(1)

    except Exception as e:
        print(f"\n❌ 处理失败：{e}")
        sys.exit(1)

    print("\n" + "=" * 70)
    print("完成")
    print("=" * 70)


if __name__ == "__main__":
    main()
