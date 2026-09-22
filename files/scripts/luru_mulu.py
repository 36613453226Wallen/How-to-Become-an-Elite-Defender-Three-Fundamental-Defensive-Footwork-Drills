#!/usr/bin/env python3
"""录入目录：把带标题截图的视频页 + 链接写入单视频收录，并累加到综合拓展目录。"""

from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path

import qrcode
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from qrcode.constants import ERROR_CORRECT_M

ROOT = Path(__file__).resolve().parents[1]
SINGLE_DIR = ROOT / "单视频收录"
DIGEST_DIR = ROOT / "综合拓展"
DIGEST_DOC = DIGEST_DIR / "视频综合目录.docx"
CATALOG = DIGEST_DIR / "条目清单.json"

FONT = "微软雅黑"
NAVY = RGBColor(0x1F, 0x2A, 0x44)
ACCENT = RGBColor(0xC4, 0x5C, 0x26)
MUTED = RGBColor(0x5C, 0x64, 0x73)
LINK_BLUE = RGBColor(0x15, 0x65, 0xC0)
LINE = RGBColor(0xD0, 0xD5, 0xDD)
FILL_HEAD = "1F2A44"
FILL_SOFT = "F4F1EA"
FILL_BOX = "FAFBFC"


def set_run_font(run, size=11, bold=False, color=None, font=FONT):
    run.font.name = font
    run.bold = bold
    run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:ascii"), font)
    rfonts.set(qn("w:hAnsi"), font)
    rfonts.set(qn("w:eastAsia"), font)
    rfonts.set(qn("w:cs"), font)


def set_cell_shading(cell, hex_color: str):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tc_pr.append(shd)


def set_cell_borders(cell, color="D0D5DD", sz="8"):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        tc_borders.append(el)
    tc_pr.append(tc_borders)


def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = OxmlElement("w:tcMar")
    for name, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{name}")
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        tc_mar.append(node)
    tc_pr.append(tc_mar)


def add_hyperlink(paragraph, text, url):
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "1565C0")
    r_pr.append(color)
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    r_pr.append(u)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "22")
    r_pr.append(sz)
    sz_cs = OxmlElement("w:szCs")
    sz_cs.set(qn("w:val"), "22")
    r_pr.append(sz_cs)
    rfonts = OxmlElement("w:rFonts")
    rfonts.set(qn("w:ascii"), FONT)
    rfonts.set(qn("w:hAnsi"), FONT)
    rfonts.set(qn("w:eastAsia"), FONT)
    rfonts.set(qn("w:cs"), FONT)
    r_pr.append(rfonts)
    new_run.append(r_pr)
    text_el = OxmlElement("w:t")
    text_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_el.text = text
    new_run.append(text_el)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def setup_document(title: str) -> Document:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)
    section.top_margin = Cm(1.6)
    section.bottom_margin = Cm(1.6)

    for style_name in ("Normal", "Title", "Heading 1", "Heading 2"):
        style = doc.styles[style_name]
        style.font.name = FONT
        style.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)

    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = hp.add_run(title)
    set_run_font(run, 9, color=MUTED)

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = fp.add_run("录入目录  ·  单视频收录 / 综合拓展")
    set_run_font(run, 8, color=MUTED)
    add_page_number(fp)
    return doc


def add_page_number(paragraph):
    run1 = paragraph.add_run("    ")
    set_run_font(run1, 8, color=MUTED)
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = " PAGE "
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_sep)
    run._r.append(fld_end)
    set_run_font(run, 8, color=MUTED)


def add_heading_bar(doc: Document, text: str):
    table = doc.add_table(rows=1, cols=1)
    table.autofit = True
    cell = table.cell(0, 0)
    set_cell_shading(cell, FILL_HEAD)
    set_cell_margins(cell, 60, 60, 140, 140)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(text)
    set_run_font(run, 12, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    spacer(doc, 8)


def spacer(doc: Document, pt=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(pt)
    p.paragraph_format.line_spacing = Pt(pt)


def add_label_value(doc: Document, label: str, value: str | None = None, url: str | None = None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(label)
    set_run_font(run, 10, bold=True, color=ACCENT)
    if url:
        p.add_run("\n")
        add_hyperlink(p, value or url, url)
    elif value:
        run2 = p.add_run(value)
        set_run_font(run2, 12, bold=False, color=NAVY)


def add_kv_table(doc: Document, rows: list[tuple[str, str]]):
    table = doc.add_table(rows=len(rows), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for i, (k, v) in enumerate(rows):
        left, right = table.cell(i, 0), table.cell(i, 1)
        set_cell_shading(left, FILL_SOFT)
        set_cell_shading(right, "FFFFFF")
        set_cell_borders(left)
        set_cell_borders(right)
        set_cell_margins(left)
        set_cell_margins(right)
        lp = left.paragraphs[0]
        rp = right.paragraphs[0]
        lr = lp.add_run(k)
        rr = rp.add_run(v)
        set_run_font(lr, 10, bold=True, color=NAVY)
        set_run_font(rr, 10, color=NAVY)
        left.width = Cm(4.2)
        right.width = Cm(13.0)
    spacer(doc, 8)


def generate_qr(url: str, dest: Path) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    qr = qrcode.QRCode(
        version=None,
        error_correction=ERROR_CORRECT_M,
        box_size=12,
        border=2,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="#1F2A44", back_color="white")
    img.save(dest)
    return dest


def load_catalog() -> list[dict]:
    if CATALOG.exists():
        return json.loads(CATALOG.read_text(encoding="utf-8"))
    return []


def save_catalog(items: list[dict]):
    CATALOG.parent.mkdir(parents=True, exist_ok=True)
    CATALOG.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")


def write_single_doc(entry: dict, screenshot: Path, qr_path: Path, dest: Path):
    dest.parent.mkdir(parents=True, exist_ok=True)
    doc = setup_document("单视频收录  ·  仅此一条视频")
    add_heading_bar(doc, "单视频收录")

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(entry["title"])
    set_run_font(run, 16, bold=True, color=NAVY)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(f"BV号 {entry['bvid']}    ·    本文件只收录这一条视频及其页面截图信息。")
    set_run_font(run, 9, color=MUTED)

    if entry.get("title_note"):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(10)
        run = p.add_run(entry["title_note"])
        set_run_font(run, 9, color=MUTED)

    add_label_value(doc, "视频链接（用户提供）", entry["url"], url=entry["url"])
    if entry.get("canonical_url") and entry["canonical_url"] != entry["url"]:
        add_label_value(doc, "规范链接", entry["canonical_url"], url=entry["canonical_url"])

    add_heading_bar(doc, "二维码")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run("手机扫码打开本视频：")
    set_run_font(run, 10, color=MUTED)
    doc.add_picture(str(qr_path), width=Cm(4.2))
    last = doc.paragraphs[-1]
    last.alignment = WD_ALIGN_PARAGRAPH.LEFT
    spacer(doc, 10)

    add_heading_bar(doc, "页面截图")
    p = doc.add_paragraph()
    run = p.add_run("以下为用户提供的视频页标题截图，已原样嵌入。")
    set_run_font(run, 10, color=MUTED)
    doc.add_picture(str(screenshot), width=Cm(17.2))
    spacer(doc, 10)

    add_heading_bar(doc, "截图信息摘录")
    add_kv_table(doc, entry["screenshot_fields"])

    if entry.get("description"):
        p = doc.add_paragraph()
        run = p.add_run("页面简介（据截图转写）")
        set_run_font(run, 10, bold=True, color=ACCENT)
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(entry["description"])
        set_run_font(run, 11, color=NAVY)

    if entry.get("related_titles"):
        p = doc.add_paragraph()
        run = p.add_run("截图中可见的相关推荐")
        set_run_font(run, 10, bold=True, color=ACCENT)
        for item in entry["related_titles"]:
            p = doc.add_paragraph(style="List Bullet")
            run = p.add_run(item)
            set_run_font(run, 10, color=NAVY)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    run = p.add_run("说明：本文件用于单视频对照（页面、链接、二维码）。动作拆解不在此份；综合评论与细节关注见《综合拓展 / 视频综合目录》。")
    set_run_font(run, 9, color=MUTED)

    doc.save(dest)


def new_digest_doc() -> Document:
    doc = setup_document("综合拓展  ·  视频综合目录")
    add_heading_bar(doc, "视频综合目录")
    p = doc.add_paragraph()
    run = p.add_run("拓展用累计目录。每条只保留标题、二维码，并留出「评论空间」「细节关注」两栏。")
    set_run_font(run, 11, color=NAVY)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(12)
    run = p.add_run("单视频动作细节属于另一追求方向，不在此列。收到「录入目录」后在文末追加，不覆盖已有条目。")
    set_run_font(run, 10, color=MUTED)
    return doc


def append_digest_entry(doc: Document, entry: dict, qr_path: Path, index: int):
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    set_cell_shading(cell, FILL_SOFT)
    set_cell_borders(cell, "C4A574", "12")
    set_cell_margins(cell, 100, 100, 140, 140)

    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(f"条目 {index:02d}")
    set_run_font(run, 9, bold=True, color=ACCENT)

    p = cell.add_paragraph()
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(entry["title"])
    set_run_font(run, 13, bold=True, color=NAVY)

    p = cell.add_paragraph()
    run = p.add_run("二维码（扫码打开视频）")
    set_run_font(run, 9, color=MUTED)
    p = cell.add_paragraph()
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run()
    run.add_picture(str(qr_path), width=Cm(3.6))

    inner = cell.add_table(rows=2, cols=2)
    inner.autofit = True
    headers = ("评论空间", "细节关注")
    hints = (
        "写下观后感、疑问、可迁移到自己训练的点。",
        "记下需要反复回看的时机、站位、脚步或口令。",
    )
    for col, (head, hint) in enumerate(zip(headers, hints)):
        head_cell = inner.cell(0, col)
        body_cell = inner.cell(1, col)
        set_cell_shading(head_cell, FILL_HEAD)
        set_cell_shading(body_cell, FILL_BOX)
        set_cell_borders(head_cell, "1F2A44")
        set_cell_borders(body_cell)
        set_cell_margins(head_cell, 60, 60, 100, 100)
        set_cell_margins(body_cell, 80, 80, 100, 100)
        hp = head_cell.paragraphs[0]
        hr = hp.add_run(head)
        set_run_font(hr, 11, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
        bp = body_cell.paragraphs[0]
        br = bp.add_run(hint)
        set_run_font(br, 9, color=MUTED)
        for _ in range(8):
            blank = body_cell.add_paragraph()
            blank.paragraph_format.space_after = Pt(8)
            run = blank.add_run(" ")
            set_run_font(run, 11)

    spacer(doc, 14)


def write_digest(entries: list[dict], qr_by_bvid: dict[str, Path]):
    DIGEST_DOC.parent.mkdir(parents=True, exist_ok=True)
    doc = new_digest_doc()
    for i, entry in enumerate(entries, start=1):
        append_digest_entry(doc, entry, qr_by_bvid[entry["bvid"]], i)
    doc.save(DIGEST_DOC)


def ingest(entry: dict, screenshot: Path):
    SINGLE_DIR.mkdir(parents=True, exist_ok=True)
    DIGEST_DIR.mkdir(parents=True, exist_ok=True)

    folder_name = f"{entry['bvid']}_{entry['short_title']}"
    video_dir = SINGLE_DIR / folder_name
    video_dir.mkdir(parents=True, exist_ok=True)

    shot_dest = video_dir / "页面截图.png"
    if screenshot.resolve() != shot_dest.resolve():
        shutil.copy2(screenshot, shot_dest)

    qr_path = video_dir / "二维码.png"
    generate_qr(entry["canonical_url"], qr_path)
    shutil.copy2(qr_path, DIGEST_DIR / f"二维码_{entry['bvid']}.png")

    single_doc = video_dir / f"{entry['bvid']}_收录.docx"
    write_single_doc(entry, shot_dest, qr_path, single_doc)

    items = load_catalog()
    existing = {item["bvid"]: i for i, item in enumerate(items)}
    if entry["bvid"] in existing:
        items[existing[entry["bvid"]]] = entry
    else:
        items.append(entry)
    save_catalog(items)

    qr_by_bvid = {}
    for item in items:
        q = SINGLE_DIR / f"{item['bvid']}_{item['short_title']}" / "二维码.png"
        if not q.exists():
            q = DIGEST_DIR / f"二维码_{item['bvid']}.png"
        qr_by_bvid[item["bvid"]] = q
    write_digest(items, qr_by_bvid)
    return {
        "single_doc": str(single_doc),
        "digest_doc": str(DIGEST_DOC),
        "qr": str(qr_path),
        "screenshot": str(shot_dest),
    }


def default_entry() -> dict:
    url = (
        "https://www.bilibili.com/video/BV1L54y1b7Wu/"
        "?spm_id_from=333.1391.0.0&vd_source=c4192c8d5e5aca3168ace4efb3460590"
    )
    return {
        "bvid": "BV1L54y1b7Wu",
        "short_title": "如何成为精英防守者",
        "title": "如何成为精英防守者？3个基础防守脚步训练 让你锁住持球人！【宝石碎片 GemPieces Vol.56】",
        "url": url,
        "canonical_url": "https://www.bilibili.com/video/BV1L54y1b7Wu/",
        "screenshot_fields": [
            ("UP主", "隐藏宝石HiddenGems"),
            ("关注数（截图）", "58.3万"),
            ("播放量（截图）", "30.7万"),
            ("弹幕（截图）", "3588"),
            ("发布时间（截图）", "2021-04-17 16:00:05"),
            ("版权提示（截图）", "未经作者授权，禁止转载"),
            ("点赞 / 投币 / 收藏 / 转发（截图）", "1.5万  /  75205  /  1.7万  /  239707"),
            ("BV号", "BV1L54y1b7Wu"),
        ],
        "description": (
            "Hey收集者，之前收到了许多关于防守的疑问。"
            "这期视频包含了成为一名精英防守者必备的基础防守脚步与概念。"
            "当然也准备了相对应的3个个人防守脚步训练，来供大家提升自己的防守能力。"
        ),
        "related_titles": [
            "防守大师思维决定高度 @徐远征",
            "单防必备技巧",
            "如何成为精英防守者？（同系列封面）",
            "防守技术分析 / 精英防守者",
            "防守基本功 / 脚步手位姿势",
            "“Cedd” 北美顶级防守专家",
            "【干货】防守正确摆放",
        ],
        "title_note": "截图标题在「【宝石」处被页面截断；完整后缀据同系列相关页补为【宝石碎片 GemPieces Vol.56】。",
    }


def main():
    parser = argparse.ArgumentParser(description="录入目录")
    parser.add_argument("--screenshot", required=True)
    args = parser.parse_args()
    result = ingest(default_entry(), Path(args.screenshot))
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
