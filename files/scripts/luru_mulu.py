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
ASSETS = Path("/home/ubuntu/.cursor/projects/workspace/assets")

FONT = "微软雅黑"
NAVY = RGBColor(0x1F, 0x2A, 0x44)
ACCENT = RGBColor(0xC4, 0x5C, 0x26)
MUTED = RGBColor(0x5C, 0x64, 0x73)
FILL_HEAD = "1F2A44"
FILL_SOFT = "F4F1EA"
FILL_BOX = "FAFBFC"
MAIN_QR_CM = 4.2
DIGEST_QR_CM = 3.6
REL_QR_CM = 2.4
REL_IMG_CM = 2.4
DIGEST_REMINDER = (
    "综合目录中不放入a滑步b交叉步c双滑步的链接和二维码，但是需要这些中文字描述，"
    "还有图二到图五的动作细节对应。图片高度（保持比例）和二维码一致，以便放在同一水平位置。"
    "（比如图二三、图三四五分别一组）。"
)


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


def add_text(doc: Document, text: str, size=11, bold=False, color=NAVY, after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(after)
    run = p.add_run(text)
    set_run_font(run, size, bold=bold, color=color)
    return p


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


def add_aligned_media_row(host, items: list[dict], height_cm: float = REL_IMG_CM):
    """Place QR/images in one row at the same height, captions underneath."""
    if not items:
        return None
    table = host.add_table(rows=2, cols=len(items))
    table.autofit = True
    for i, item in enumerate(items):
        img_cell = table.cell(0, i)
        cap_cell = table.cell(1, i)
        set_cell_shading(img_cell, "FFFFFF")
        set_cell_shading(cap_cell, "FFFFFF")
        set_cell_borders(img_cell, "E8E8E8", "4")
        set_cell_borders(cap_cell, "E8E8E8", "4")
        set_cell_margins(img_cell, 40, 40, 50, 50)
        set_cell_margins(cap_cell, 20, 20, 50, 50)
        p = img_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run()
        run.add_picture(item["path"], height=Cm(height_cm))
        cp = cap_cell.paragraphs[0]
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cp.add_run(item.get("caption", ""))
        set_run_font(run, 8, color=MUTED)
    return table


def load_catalog() -> list[dict]:
    if CATALOG.exists():
        return json.loads(CATALOG.read_text(encoding="utf-8"))
    return []


def save_catalog(items: list[dict]):
    CATALOG.parent.mkdir(parents=True, exist_ok=True)
    CATALOG.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")


def folder_name(entry: dict) -> str:
    return f"{entry['index']:02d}_{entry['bvid']}_{entry['short_title']}"


def write_single_doc(entry: dict, dest: Path):
    dest.parent.mkdir(parents=True, exist_ok=True)
    idx = entry["index"]
    doc = setup_document(f"单视频收录  ·  {idx:02d}  ·  仅此一条视频")
    add_heading_bar(doc, f"{idx:02d}  单视频收录")

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(entry["title"])
    set_run_font(run, 16, bold=True, color=NAVY)

    add_text(
        doc,
        f"序号 {idx}    ·    BV号 {entry['bvid']}    ·    本文件只收录这一条视频及其页面截图信息。",
        size=9,
        color=MUTED,
        after=6,
    )

    if entry.get("title_note"):
        add_text(doc, entry["title_note"], size=9, color=MUTED, after=10)

    add_label_value(doc, "视频链接（用户提供）", entry["url"], url=entry["url"])
    if entry.get("canonical_url") and entry["canonical_url"] != entry["url"]:
        add_label_value(doc, "规范链接", entry["canonical_url"], url=entry["canonical_url"])

    add_heading_bar(doc, "二维码")
    add_text(doc, "手机扫码打开本视频：", size=10, color=MUTED, after=4)
    doc.add_picture(entry["qr_path"], width=Cm(MAIN_QR_CM))
    spacer(doc, 10)

    add_heading_bar(doc, "页面截图")
    shot_note = "以下为用户提供的视频页界面图（图一），已原样嵌入。" if entry.get("interface_shot") else "以下为用户提供的视频页标题截图，已原样嵌入。"
    add_text(doc, shot_note, size=10, color=MUTED, after=6)
    doc.add_picture(entry["screenshot_path"], width=Cm(17.2))
    spacer(doc, 10)

    add_heading_bar(doc, "截图信息摘录")
    add_kv_table(doc, [tuple(row) for row in entry["screenshot_fields"]])

    if entry.get("description"):
        add_text(doc, "页面简介（据截图转写）", size=10, bold=True, color=ACCENT, after=4)
        add_text(doc, entry["description"], size=11, color=NAVY, after=8)

    if entry.get("related_titles"):
        add_text(doc, "截图中可见的相关推荐", size=10, bold=True, color=ACCENT, after=4)
        for item in entry["related_titles"]:
            p = doc.add_paragraph(style="List Bullet")
            run = p.add_run(item)
            set_run_font(run, 10, color=NAVY)

    extras = entry.get("extras") or {}
    if extras.get("related_title") or extras.get("comment_image"):
        add_heading_bar(doc, "文后附图  ·  评论区与延伸视频")
        if extras.get("related_title"):
            add_text(doc, extras["related_title"], size=14, bold=True, color=NAVY, after=4)
            add_label_value(doc, "延伸视频链接", extras["related_url"], url=extras["related_url"])
            add_text(doc, "扫码打开延伸视频：", size=10, color=MUTED, after=4)
            doc.add_picture(extras["related_qr"], width=Cm(REL_QR_CM))
            spacer(doc, 8)
        if extras.get("comment_image"):
            add_text(doc, "评论区截图（图一，按页宽排版）", size=10, bold=True, color=ACCENT, after=4)
            add_text(
                doc,
                extras.get(
                    "comment_note",
                    "红框两处：脚后跟先着地再到前脚掌；具体细节可对照隔壁 KBT。",
                ),
                size=10,
                color=MUTED,
                after=6,
            )
            doc.add_picture(extras["comment_image"], width=Cm(16.4))
            spacer(doc, 8)
            if extras.get("comment_points"):
                add_text(doc, "评论要点摘录", size=10, bold=True, color=ACCENT, after=4)
                for item in extras["comment_points"]:
                    p = doc.add_paragraph(style="List Bullet")
                    run = p.add_run(item)
                    set_run_font(run, 10, color=NAVY)

    write_single_relations(doc, entry.get("relations") or [])

    if entry.get("end_reminder"):
        add_heading_bar(doc, "提示词")
        add_text(doc, entry["end_reminder"], size=10, color=MUTED, after=8)

    add_text(
        doc,
        "说明：本文件用于单视频对照（页面、链接、二维码）。动作拆解不在此份；综合评论与细节关注见《综合拓展 / 视频综合目录》。",
        size=9,
        color=MUTED,
        after=4,
    )
    doc.save(dest)


def write_single_relations(doc: Document, relations: list[dict]):
    if not relations:
        return
    add_heading_bar(doc, "视频间联系")
    add_text(doc, "精准空降链接配略小二维码；动作图与该二维码同高，排在同一水平。", size=9, color=MUTED, after=8)
    for rel in relations:
        add_text(doc, rel["name"], size=13, bold=True, color=NAVY, after=4)
        add_text(doc, rel["description"], size=11, after=6)
        if rel.get("link_title"):
            add_text(doc, rel["link_title"], size=10, bold=True, color=NAVY, after=2)
        if rel.get("url"):
            add_label_value(doc, "精准空降链接", rel["url"], url=rel["url"])
        row = []
        if rel.get("qr_path"):
            row.append({"path": rel["qr_path"], "caption": "精准二维码"})
        for img in rel.get("images") or []:
            row.append({"path": img["path"], "caption": img["label"]})
        if row:
            add_aligned_media_row(doc, row, height_cm=REL_IMG_CM)
            spacer(doc, 10)
        for extra in rel.get("extra_images") or []:
            add_text(doc, extra["label"], size=10, bold=True, color=ACCENT, after=4)
            doc.add_picture(extra["path"], width=Cm(8.0))
            spacer(doc, 8)


def new_digest_doc() -> Document:
    doc = setup_document("综合拓展  ·  视频综合目录")
    add_heading_bar(doc, "视频综合目录")
    add_text(
        doc,
        "拓展用累计目录。每条按 1. 2. 3. 编号，保留标题、本视频二维码，并留出「评论空间」「细节关注」两栏。视频间联系只写中文和动作图。",
        size=11,
        color=NAVY,
        after=4,
    )
    add_text(
        doc,
        "单视频动作细节属于另一追求方向，不在此列。收到「录入目录」后在文末追加，不覆盖已有条目。",
        size=10,
        color=MUTED,
        after=12,
    )
    return doc


def append_digest_entry(doc: Document, entry: dict):
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    set_cell_shading(cell, FILL_SOFT)
    set_cell_borders(cell, "C4A574", "12")
    set_cell_margins(cell, 100, 100, 140, 140)

    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(f"{entry['index']}.")
    set_run_font(run, 14, bold=True, color=ACCENT)

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
    run.add_picture(entry["qr_path"], width=Cm(DIGEST_QR_CM))

    extras = entry.get("extras") or {}

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

    if extras.get("comment_image") or extras.get("related_title"):
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        run = p.add_run("文后附图")
        set_run_font(run, 9, bold=True, color=ACCENT)
        if extras.get("related_title"):
            p = cell.add_paragraph()
            run = p.add_run(extras["related_title"])
            set_run_font(run, 10, bold=True, color=NAVY)
            p = cell.add_paragraph()
            run = p.add_run("延伸视频二维码")
            set_run_font(run, 9, color=MUTED)
            p = cell.add_paragraph()
            p.paragraph_format.space_after = Pt(6)
            run = p.add_run()
            run.add_picture(extras["related_qr"], width=Cm(REL_QR_CM))
        if extras.get("comment_image"):
            p = cell.add_paragraph()
            run = p.add_run("评论区截图（图一）")
            set_run_font(run, 9, color=MUTED)
            p = cell.add_paragraph()
            p.paragraph_format.space_after = Pt(6)
            run = p.add_run()
            run.add_picture(extras["comment_image"], width=Cm(15.2))

    write_digest_relations(cell, entry.get("relations") or [])
    spacer(doc, 14)


def write_digest_relations(cell, relations: list[dict]):
    if not relations:
        return
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    run = p.add_run("视频间联系")
    set_run_font(run, 11, bold=True, color=ACCENT)
    p = cell.add_paragraph()
    run = p.add_run("只写中文描述和动作图，不放 a滑步 / b交叉步 / c连续滑步 的链接与二维码。图与本条视频二维码同高。")
    set_run_font(run, 9, color=MUTED)

    for rel in relations:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        run = p.add_run(rel["name"])
        set_run_font(run, 11, bold=True, color=NAVY)
        p = cell.add_paragraph()
        run = p.add_run(rel["description"])
        set_run_font(run, 10, color=NAVY)

    groups = []
    seen = set()
    for rel in relations:
        for img in rel.get("digest_images") or rel.get("images") or []:
            key = img.get("path")
            if key and key not in seen:
                seen.add(key)
                groups.append(img)
    # 图二三一组，图四五一组（图三四五按图四+图五）
    pair_a = [img for img in groups if img.get("label") in {"图二", "图三"}]
    pair_b = [img for img in groups if img.get("label") in {"图四", "图五", "图五侧面"}]
    if pair_a:
        p = cell.add_paragraph()
        run = p.add_run("图二三一组")
        set_run_font(run, 9, bold=True, color=ACCENT)
        add_aligned_media_row(cell, [{"path": i["path"], "caption": i["label"]} for i in pair_a], height_cm=DIGEST_QR_CM)
    if pair_b:
        p = cell.add_paragraph()
        run = p.add_run("图四五一组")
        set_run_font(run, 9, bold=True, color=ACCENT)
        add_aligned_media_row(cell, [{"path": i["path"], "caption": i["label"]} for i in pair_b], height_cm=DIGEST_QR_CM)


def write_digest(entries: list[dict]):
    DIGEST_DOC.parent.mkdir(parents=True, exist_ok=True)
    doc = new_digest_doc()
    for entry in entries:
        append_digest_entry(doc, entry)
    doc.save(DIGEST_DOC)


def materialize_entry(entry: dict) -> dict:
    """Copy assets, generate QR, write numbered files, return paths on the entry."""
    idx = entry["index"]
    video_dir = SINGLE_DIR / folder_name(entry)
    if video_dir.exists():
        shutil.rmtree(video_dir)
    video_dir.mkdir(parents=True, exist_ok=True)

    shot_src = Path(entry["screenshot_src"])
    shot_dest = video_dir / f"02_页面截图.png"
    shutil.copy2(shot_src, shot_dest)

    qr_path = video_dir / "03_二维码.png"
    generate_qr(entry["canonical_url"], qr_path)
    digest_qr = DIGEST_DIR / f"{idx:02d}_二维码_{entry['bvid']}.png"
    shutil.copy2(qr_path, digest_qr)

    extras = dict(entry.get("extras") or {})
    file_i = 4
    if extras.get("comment_src"):
        comment_dest = video_dir / f"{file_i:02d}_评论区.png"
        shutil.copy2(extras["comment_src"], comment_dest)
        extras["comment_image"] = str(comment_dest)
        file_i += 1
    if extras.get("related_url"):
        related_qr = video_dir / f"{file_i:02d}_延伸_{extras.get('related_bvid', 'video')}_二维码.png"
        generate_qr(extras["related_canonical"], related_qr)
        extras["related_qr"] = str(related_qr)
        file_i += 1

    relations = []
    for rel in entry.get("relations") or []:
        rel = dict(rel)
        if rel.get("url"):
            q = video_dir / f"{file_i:02d}_{rel['code']}_精准二维码.png"
            generate_qr(rel["url"], q)
            rel["qr_path"] = str(q)
            file_i += 1
        copied = []
        for img in rel.get("images") or []:
            dest = video_dir / f"{file_i:02d}_{img['label']}.png"
            shutil.copy2(img["src"], dest)
            copied.append({"label": img["label"], "path": str(dest)})
            file_i += 1
        rel["images"] = copied
        extra_copied = []
        for img in rel.get("extra_images") or []:
            dest = video_dir / f"{file_i:02d}_{img['label']}.png"
            shutil.copy2(img["src"], dest)
            extra_copied.append({"label": img["label"], "path": str(dest)})
            file_i += 1
        rel["extra_images"] = extra_copied
        relations.append(rel)

    entry = dict(entry)
    entry["qr_path"] = str(qr_path)
    entry["screenshot_path"] = str(shot_dest)
    entry["extras"] = extras
    entry["relations"] = relations
    entry["folder"] = str(video_dir)

    short = entry.get("file_title") or entry["short_title"]
    single_doc = video_dir / f"01_{short}.docx"
    write_single_doc(entry, single_doc)
    entry["single_doc"] = str(single_doc)
    return entry


def write_kbt_angle_docs():
    folder = SINGLE_DIR / "如何才能打造铁血防守？Kbt迈脚角度验证"
    if folder.exists():
        shutil.rmtree(folder)
    folder.mkdir(parents=True, exist_ok=True)

    fig2_src = ASSETS / "0a0981d6-88f6-4558-8f37-f8b59c96e400.png"
    fig2_dest = folder / "00_图二原图.png"
    shutil.copy2(fig2_src, fig2_dest)

    brief = folder / "01_如何才能打造铁血防守？Kbt迈脚角度验证.docx"
    full = folder / "02_如何才能打造铁血防守？Kbt迈脚角度验证Grok4.3.docx"

    doc = setup_document("迈脚角度验证  ·  简短结论")
    add_heading_bar(doc, "01  如何才能打造铁血防守？Kbt迈脚角度验证")
    add_text(doc, "【KBT干货】如何才能打造铁血防守？今天KBT告诉你2个细节！", size=13, bold=True, after=6)
    add_text(doc, "来源：图二（Grok 4.3 对「滑步迈步脚角度」的核验）。本份只留短结论。", size=10, color=MUTED, after=10)

    add_heading_bar(doc, "一句话")
    add_text(
        doc,
        "迈步脚朝向移动方向、后跟先着地再滚到脚尖，这一点基本正确；明显外开是求速度时的进阶用法，不是人人随时都要开很大。",
        size=12,
        bold=True,
        after=10,
    )

    add_heading_bar(doc, "三点")
    for item in (
        "不要用脚掌侧面拖地。侧缘着地会加大摩擦，滑步发顿、发慢。",
        "迈步脚脚尖略朝移动方向（open the lead foot / angle the stepping foot），让脚掌更能正对横移方向。",
        "每一步后跟先着地，再把力滚到脚尖（heel-to-toe），横移才推得走。",
    ):
        p = doc.add_paragraph(style="List Number")
        run = p.add_run(item)
        set_run_font(run, 11, color=NAVY)

    spacer(doc, 8)
    add_heading_bar(doc, "只需补一句")
    add_text(
        doc,
        "传统教学仍多讲双脚大致平行、脚尖朝前。只有快速启防、强调横移速度时，才明显打开迈步脚。开太大可能丢掉低姿和横向稳定。",
        size=11,
        after=8,
    )
    add_text(doc, "完整拆解与外网对照见同目录「02_……Grok4.3.docx」。", size=9, color=MUTED, after=4)
    doc.save(brief)

    doc = setup_document("迈脚角度验证  ·  Grok4.3")
    add_heading_bar(doc, "02  如何才能打造铁血防守？Kbt迈脚角度验证Grok4.3")
    add_text(doc, "同名加后缀 Grok4.3。据图二转写并整理，保留验证结构。", size=10, color=MUTED, after=8)

    add_heading_bar(doc, "原问题（图二上方）")
    add_text(
        doc,
        "在 KBT 篮球训练的其中一个视频里，作者文案大意是：滑步速度代表防守质量。习惯用脚掌侧面接触地面，会加大摩擦、变慢。正确做法是迈步脚朝向侧面；并且每滑一步，先让脚后跟落地，再把力传到脚尖、滚动向前，横移才会变快。",
        size=11,
        after=6,
    )
    add_text(
        doc,
        "要核验的重点是：「滑步的迈步脚的角度」是不是基本正确？对照外网、专业、NBA 常见教学，且不要用门槛过高、普通人做不到的素质动作来证明。",
        size=11,
        after=10,
    )

    add_heading_bar(doc, "结论")
    add_text(doc, "基本正确，但要按情境微调，不能把它说成唯一标准姿势。", size=12, bold=True, after=10)

    add_heading_bar(doc, "核心论点拆解")
    add_text(doc, "1. 「用脚掌侧面接触地面会增大摩擦、降低速度」——正确", size=12, bold=True, after=4)
    add_text(
        doc,
        "这是 defensive slide / shuffle 里的常见错误提醒。翻脚、只用内外侧缘着地，会加大阻力，阻碍推地和滚动，容易拖步、卡顿。外网常见口令是 full foot contact，或 avoid dragging the side of the foot。",
        size=11,
        after=8,
    )
    add_text(doc, "2. 「迈步脚朝向侧面」——基本正确，表述要更精确", size=12, bold=True, after=4)
    add_text(
        doc,
        "实际教的是 open the lead foot / angle the stepping foot：向左侧滑时，左脚（迈步脚）脚尖略指向左侧，而不是完全平行底线、正对进攻者。这样脚掌更能正对移动方向，方便 heel-to-toe rolling，少侧向拖拽，横向推力更干净。",
        size=11,
        after=6,
    )
    add_text(
        doc,
        "这和传统「双脚大致平行、脚尖朝前」的纯 shuffle 有细微差别，属于追求横移速度时的进阶变体，部分教练会用。",
        size=11,
        after=8,
    )
    add_text(doc, "3. 「每步后跟先着地，再滚到脚尖」——正确", size=12, bold=True, after=4)
    add_text(
        doc,
        "这是地面接触的经典 cue。Heel-to-toe roll 减少制动、提高滚动效率。篮球、网球、足球的横向移动里都常见，接近 active feet / light on your feet。",
        size=11,
        after=10,
    )

    add_heading_bar(doc, "外网 / 专业 / NBA 对照")
    for item in (
        "USA Basketball / FIBA：push with the outside foot, step with the lead foot；提醒不要 turning the feet sideways too much（过度外翻），同时会教 angle the lead foot slightly，方便发力。",
        "NBA / G League 常见资源（如 Mike Brown、Adrian Griffin 一类脚步模块）：lateral quickness、defensive footwork 里会出现 turn the lead foot、open the hips、roll through the foot，尤其在强调 speed of slide 和 low and wide 时。",
        "力学上：迈步脚略朝移动方向，推力更接近水平横向，少跟地面较劲。",
    ):
        p = doc.add_paragraph(style="List Bullet")
        run = p.add_run(item)
        set_run_font(run, 11, color=NAVY)

    spacer(doc, 8)
    add_heading_bar(doc, "总结判断")
    add_text(
        doc,
        "「滑步的迈步脚角度」抓住了两件事：避免侧面拖拽，以及允许后跟到脚尖的滚动。方向正确，也能加快横移。",
        size=11,
        after=6,
    )
    add_text(
        doc,
        "必须补的一句：传统防守滑步仍以双脚大致平行、脚尖朝前为主。只有快速启防、强调更大横移速度时，才明显 open the lead foot。过度朝侧面，可能牺牲低姿势和横向稳定。",
        size=11,
        after=6,
    )
    add_text(
        doc,
        "因此，视频作者这一点方向正确、能用，属于有依据的进阶教学，不是人人随时都要把脚开到很大。",
        size=11,
        bold=True,
        after=10,
    )

    add_heading_bar(doc, "图二原图")
    add_text(doc, "下面嵌入用户提供的图二，便于对照转写。", size=10, color=MUTED, after=6)
    doc.add_picture(str(fig2_dest), width=Cm(16.0))
    doc.save(full)

    return {"folder": str(folder), "brief": str(brief), "grok": str(full)}


def catalog_entries() -> list[dict]:
    kbt_url = (
        "https://www.bilibili.com/video/BV11U4y1b7EV/"
        "?spm_id_from=333.337.search-card.all.click&vd_source=c4192c8d5e5aca3168ace4efb3460590"
    )
    kbt_title = "【KBT干货】如何才能打造铁血防守？今天KBT告诉你2个细节！"
    return [
        {
            "index": 1,
            "bvid": "BV1L54y1b7Wu",
            "short_title": "如何成为精英防守者",
            "file_title": "精英防守者",
            "title": "如何成为精英防守者？3个基础防守脚步训练 让你锁住持球人！【宝石碎片 GemPieces Vol.56】",
            "url": (
                "https://www.bilibili.com/video/BV1L54y1b7Wu/"
                "?spm_id_from=333.1391.0.0&vd_source=c4192c8d5e5aca3168ace4efb3460590"
            ),
            "canonical_url": "https://www.bilibili.com/video/BV1L54y1b7Wu/",
            "screenshot_src": str(ASSETS / "a73333e2-3583-433d-b2b6-c9642e729039.png"),
            "screenshot_fields": [
                ["UP主", "隐藏宝石HiddenGems"],
                ["关注数（截图）", "58.3万"],
                ["播放量（截图）", "30.7万"],
                ["弹幕（截图）", "3588"],
                ["发布时间（截图）", "2021-04-17 16:00:05"],
                ["版权提示（截图）", "未经作者授权，禁止转载"],
                ["点赞 / 投币 / 收藏 / 转发（截图）", "1.5万  /  75205  /  1.7万  /  239707"],
                ["BV号", "BV1L54y1b7Wu"],
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
            "title_note": "标题后缀已由用户确认为【宝石碎片 GemPieces Vol.56】。",
            "extras": {
                "comment_src": str(ASSETS / "a10088bc-956a-481a-88d4-8a8fb6c6bc6c.png"),
                "comment_note": "图一按页宽放在文后。红框两处：落地顺序，以及指向隔壁 KBT。",
                "comment_points": [
                    "是阿森啦：做步时前脚掌着地再过渡，不要脚面侧向用力。",
                    "求道楼（红框）：就是要脚后跟着地先，再到前脚掌着地。",
                    "牛牛不牛（红框点出 KBT）：具体可以看隔壁 KBT 讲得很细。",
                    "是阿森啦：壮哥更注重训练方法，KBT 更偏技术细节。",
                ],
                "related_title": kbt_title,
                "related_url": kbt_url,
                "related_canonical": "https://www.bilibili.com/video/BV11U4y1b7EV/",
                "related_bvid": "BV11U4y1b7EV",
            },
        },
        {
            "index": 2,
            "bvid": "BV11U4y1b7EV",
            "short_title": "如何才能打造铁血防守",
            "file_title": "铁血防守",
            "title": kbt_title,
            "url": kbt_url,
            "canonical_url": "https://www.bilibili.com/video/BV11U4y1b7EV/",
            "screenshot_src": str(ASSETS / "2de1e477-e587-4cac-8c8f-5a429fdd0780.png"),
            "screenshot_fields": [
                ["标题（截图）", kbt_title],
                ["画面水印", "KBT篮球训练营（标题栏未截出 UP 名，故不单列 UP 主）"],
                ["播放量（截图）", "19万"],
                ["弹幕（截图）", "81"],
                ["发布时间（截图）", "2021-04-17 15:06:02"],
                ["版权提示（截图）", "未经作者授权，禁止转载"],
                ["画面字幕（截图）", "连续的大幅度转髋跳练习"],
                ["点赞 / 投币 / 收藏 / 转发（截图）", "72228  /  11777  /  51332  /  48404"],
                ["BV号", "BV11U4y1b7EV"],
            ],
            "description": "图三为视频播放页。标题完整可见；画面正在做大幅度转髋跳。",
            "title_note": "图三标题完整，按截图录入。",
        },
        {
            "index": 3,
            "bvid": "BV1oM411y7Bm",
            "short_title": "快速离心",
            "file_title": "快速离心",
            "interface_shot": True,
            "title": (
                "被忽视的“快速离心”能力？真正解锁你的运球重心！一套30分钟下肢综合训练计划 "
                "涵盖下肢灵活度/基础肌力/快速离心【宝石碎片 GemPieces Vol.108】"
            ),
            "url": "https://www.bilibili.com/video/BV1oM411y7Bm/?spm_id_from=333.1391.0.0",
            "canonical_url": "https://www.bilibili.com/video/BV1oM411y7Bm/",
            "screenshot_src": str(ASSETS / "9cecc1f5-d77f-4eb2-8918-58d813274842.png"),
            "screenshot_fields": [
                ["UP主", "隐藏宝石HiddenGems"],
                ["关注数（截图）", "58.3万"],
                ["播放量（截图）", "16.6万"],
                ["弹幕（截图）", "301"],
                ["发布时间（截图）", "2023-01-04 18:00:00"],
                ["版权提示（截图）", "未经作者授权，禁止转载"],
                ["点赞 / 投币 / 收藏 / 转发（截图）", "8519  /  4244  /  1.2万  /  10668"],
                ["画面提示（图一）", "其中分为3个模块；绿色折线标在髋部"],
                ["BV号", "BV1oM411y7Bm"],
            ],
            "description": (
                "Hey收集者。首先感谢大家22年对频道的支持与陪伴，祝大家2023年一切顺利，一起继续前行。"
                "收到大家对上一期运球重心视频的支持后，这期带来相应的下肢综合训练计划。"
                "围绕下肢关节灵活度、基础肌力与快速离心安排动作，不需要铃也可以完成。"
            ),
            "title_note": "标题由用户完整给出，图一界面图一并录入。",
            "end_reminder": DIGEST_REMINDER,
            "relations": [
                {
                    "code": "a滑步",
                    "name": "a滑步",
                    "description": (
                        "Vol.108 应结合 Vol.113 的 a滑步。"
                        "精准空降 Vol.113 的 02:55，是快速离心的后置动作（图二：#4 快速离心降+横移）。"
                    ),
                    "link_title": (
                        "【拿出每天1%的时间 全面提升你的防守移动能力！15分钟防守专项跟练视频 "
                        "防守脚步/移动能力/体能心肺【宝石碎片 GemPieces Vol.113】】【精准空降到 02:55】"
                    ),
                    "url": (
                        "https://www.bilibili.com/video/BV1Sh411G746/"
                        "?share_source=copy_web&vd_source=731103ad1d48157617bc27e0f3b87025&t=175"
                    ),
                    "images": [
                        {
                            "label": "图二",
                            "src": str(ASSETS / "3e16db49-238b-4f5f-8e21-3b71aca8040c.png"),
                        }
                    ],
                },
                {
                    "code": "b交叉步",
                    "name": "b交叉步",
                    "description": (
                        "Vol.108 应结合 Vol.113 的 b交叉步。"
                        "精准空降 Vol.113 的 03:56，是图二「快速离心+横移」的后置动作，中间有罚篮休息（图三：#5 快速离心降+横移交叉步）。"
                    ),
                    "link_title": (
                        "【拿出每天1%的时间 全面提升你的防守移动能力！15分钟防守专项跟练视频 "
                        "防守脚步/移动能力/体能心肺【宝石碎片 GemPieces Vol.113】】【精准空降到 03:56】"
                    ),
                    "url": (
                        "https://www.bilibili.com/video/BV1Sh411G746/"
                        "?share_source=copy_web&vd_source=731103ad1d48157617bc27e0f3b87025&t=236"
                    ),
                    "images": [
                        {
                            "label": "图三",
                            "src": str(ASSETS / "cf0a099c-0bb8-44fe-a535-21ccaef0c023.png"),
                        }
                    ],
                },
                {
                    "code": "c连续滑步",
                    "name": "c连续滑步",
                    "description": (
                        "c连续滑步跟在 Vol.113 的 a滑步后面。"
                        "动作细节见图四、图五（两次转髋 hip turn 后的防守滑步）；图六是封面图。"
                    ),
                    "link_title": (
                        "【如何成为精英防守者？3个基础防守脚步训练 让你锁住持球人！"
                        "【宝石碎片 GemPieces Vol.56】】【精准空降到 01:22】"
                    ),
                    "url": (
                        "https://www.bilibili.com/video/BV1L54y1b7Wu/"
                        "?share_source=copy_web&vd_source=731103ad1d48157617bc27e0f3b87025&t=82"
                    ),
                    "images": [
                        {
                            "label": "图四",
                            "src": str(ASSETS / "a64a6358-cfe8-455b-a64f-f4b41e62ee77.png"),
                        },
                        {
                            "label": "图五",
                            "src": str(ASSETS / "dee92f59-3541-46e8-b936-0be7376346af.png"),
                        },
                        {
                            "label": "图五侧面",
                            "src": str(ASSETS / "f1d2d78c-25fd-4059-8720-7e70a53cb94b.png"),
                        },
                    ],
                    "extra_images": [
                        {
                            "label": "图六封面",
                            "src": str(ASSETS / "0cd3be50-3615-4aac-96de-c1fe17ecb228.png"),
                        }
                    ],
                },
            ],
        },
    ]


def rebuild():
    SINGLE_DIR.mkdir(parents=True, exist_ok=True)
    DIGEST_DIR.mkdir(parents=True, exist_ok=True)

    stale = SINGLE_DIR / "BV1L54y1b7Wu_如何成为精英防守者"
    if stale.exists():
        shutil.rmtree(stale)
    for leftover in DIGEST_DIR.glob("二维码_*.png"):
        leftover.unlink()

    entries = [materialize_entry(item) for item in catalog_entries()]
    serializable = []
    for item in entries:
        keep = {k: v for k, v in item.items() if k not in {"screenshot_src"}}
        serializable.append(keep)
    save_catalog(serializable)
    write_digest(entries)
    analysis = write_kbt_angle_docs()
    return {
        "entries": [
            {"index": e["index"], "title": e["title"], "folder": e["folder"], "single_doc": e["single_doc"]}
            for e in entries
        ],
        "digest_doc": str(DIGEST_DOC),
        "analysis": analysis,
    }


def main():
    parser = argparse.ArgumentParser(description="录入目录")
    parser.add_argument("--rebuild", action="store_true")
    args = parser.parse_args()
    if args.rebuild:
        print(json.dumps(rebuild(), ensure_ascii=False, indent=2))
        return
    parser.error("请使用 --rebuild 按当前条目清单生成")


if __name__ == "__main__":
    main()
