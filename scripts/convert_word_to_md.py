from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile


WORDCONV = Path(r"C:\Program Files\Microsoft Office\Office14\wordconv.exe")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def xml_text(element: ET.Element | None) -> str:
    if element is None:
        return ""
    texts: list[str] = []
    for node in element.iter():
        if node.tag == f"{{{W_NS}}}tab":
            texts.append("\t")
        elif node.tag == f"{{{W_NS}}}br":
            texts.append("\n")
        elif node.text:
            texts.append(node.text)
    return "".join(texts)


def normalize_text(text: str) -> str:
    text = text.replace("\r", "")
    text = text.replace("\xa0", " ")
    text = text.replace("\t", "    ")
    lines = [re.sub(r"[ \t]+$", "", line) for line in text.split("\n")]
    text = "\n".join(lines)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def load_style_map(zf: ZipFile) -> dict[str, str]:
    style_map: dict[str, str] = {}
    try:
        root = ET.fromstring(zf.read("word/styles.xml"))
    except KeyError:
        return style_map

    for style in root.findall("w:style", NS):
        style_id = style.attrib.get(f"{{{W_NS}}}styleId")
        if not style_id:
            continue
        name = style.find("w:name", NS)
        if name is not None:
            style_map[style_id] = name.attrib.get(f"{{{W_NS}}}val", style_id)
    return style_map


def heading_level(style_name: str) -> int | None:
    lowered = style_name.lower()
    if "heading 1" in lowered or lowered == "title":
        return 1
    if "heading 2" in lowered:
        return 2
    if "heading 3" in lowered:
        return 3
    if "heading 4" in lowered:
        return 4
    if any(token in style_name for token in ("標題 1", "标题 1", "標題1", "标题1", "主標題")):
        return 1
    if any(token in style_name for token in ("標題 2", "标题 2", "標題2", "标题2")):
        return 2
    if any(token in style_name for token in ("標題 3", "标题 3", "標題3", "标题3")):
        return 3
    return None


def paragraph_to_markdown(paragraph: ET.Element, style_map: dict[str, str]) -> str:
    p_pr = paragraph.find("w:pPr", NS)
    style_name = ""
    is_list = False
    if p_pr is not None:
        p_style = p_pr.find("w:pStyle", NS)
        if p_style is not None:
            style_id = p_style.attrib.get(f"{{{W_NS}}}val", "")
            style_name = style_map.get(style_id, style_id)
        is_list = p_pr.find("w:numPr", NS) is not None

    text = normalize_text(xml_text(paragraph))
    if not text:
        return ""

    level = heading_level(style_name)
    if level is not None:
        return f"{'#' * level} {text}"
    if is_list:
        return f"- {text}"
    return text


def table_to_markdown(table: ET.Element) -> list[str]:
    rows: list[list[str]] = []
    for tr in table.findall("w:tr", NS):
        cells: list[str] = []
        for tc in tr.findall("w:tc", NS):
            parts: list[str] = []
            for child in tc:
                if child.tag == f"{{{W_NS}}}p":
                    paragraph = normalize_text(xml_text(child))
                    if paragraph:
                        parts.append(paragraph.replace("\n", "<br>"))
            cells.append(" ".join(parts).strip())
        if any(cell for cell in cells):
            rows.append(cells)

    if not rows:
        return []

    width = max(len(row) for row in rows)
    rows = [row + [""] * (width - len(row)) for row in rows]
    lines = ["| " + " | ".join(row) + " |" for row in rows]
    lines.insert(1, "| " + " | ".join(["---"] * width) + " |")
    return lines


def docx_to_markdown(docx_path: Path, title: str) -> str:
    with ZipFile(docx_path) as zf:
        style_map = load_style_map(zf)
        document = ET.fromstring(zf.read("word/document.xml"))

    body = document.find("w:body", NS)
    if body is None:
        return f"# {title}\n"

    lines = [f"# {title}"]
    for child in body:
        if child.tag == f"{{{W_NS}}}p":
            paragraph = paragraph_to_markdown(child, style_map)
            if paragraph:
                if lines[-1] != "":
                    lines.append("")
                lines.append(paragraph)
        elif child.tag == f"{{{W_NS}}}tbl":
            table_lines = table_to_markdown(child)
            if table_lines:
                if lines[-1] != "":
                    lines.append("")
                lines.extend(table_lines)

    text = "\n".join(lines).strip() + "\n"
    return re.sub(r"\n{3,}", "\n\n", text)


def convert_doc_to_docx(doc_path: Path, temp_root: Path) -> Path:
    if not WORDCONV.exists():
        raise FileNotFoundError(f"Missing Word converter: {WORDCONV}")

    temp_dir = Path(tempfile.mkdtemp(prefix="wordconv_", dir=str(temp_root)))
    out_path = temp_dir / f"{doc_path.stem}.docx"
    subprocess.run(
        [str(WORDCONV), "-oice", "-nme", str(doc_path.resolve()), str(out_path.resolve())],
        check=True,
    )
    if not out_path.exists():
        raise RuntimeError(f"Conversion failed for {doc_path}")
    return out_path


def choose_sources(base_dir: Path) -> list[Path]:
    candidates = sorted(path for path in base_dir.iterdir() if path.suffix.lower() in {".doc", ".docx"})
    by_stem: dict[str, Path] = {}
    for path in candidates:
        existing = by_stem.get(path.stem)
        if existing is None or (existing.suffix.lower() == ".docx" and path.suffix.lower() == ".doc"):
            by_stem[path.stem] = path
    return sorted(by_stem.values())


def convert_directory(base_dir: Path) -> None:
    if not base_dir.exists():
        raise FileNotFoundError(f"Directory not found: {base_dir}")
    if not base_dir.is_dir():
        raise NotADirectoryError(f"Not a directory: {base_dir}")

    for path in choose_sources(base_dir):
        if path.suffix.lower() == ".docx":
            source_docx = path
            temp_dir = None
        else:
            source_docx = convert_doc_to_docx(path, temp_root=base_dir)
            temp_dir = source_docx.parent

        try:
            markdown = docx_to_markdown(source_docx, path.stem)
            md_path = path.with_suffix(".md")
            md_path.write_text(markdown, encoding="utf-8")
            print(f"Wrote {md_path}")
        finally:
            if temp_dir is not None and temp_dir.exists():
                shutil.rmtree(temp_dir)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert .doc/.docx files in a directory to Markdown.")
    parser.add_argument("directory", nargs="?", default="speech/36-topics", help="Directory containing Word files")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    convert_directory(Path(args.directory))


if __name__ == "__main__":
    main()
