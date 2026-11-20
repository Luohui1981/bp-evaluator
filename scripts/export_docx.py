"""Export a bp-evaluator Markdown version to the local Word presentation format."""

import argparse
from datetime import datetime
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import re
import subprocess
from urllib.parse import urlsplit
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree
import yaml

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"w": W}
PAGE_WIDTH, PAGE_HEIGHT, SIDE_MARGIN = 11906, 16838, 1276
WIDTH = PAGE_WIDTH - SIDE_MARGIN * 2
BLACK, GRAY, LINE = "000000", "F1F1F1", "C8C8C8"
TITLE_FONT, BODY_FONT, NOTE_FONT = "Microsoft YaHei", "SimSun", "楷体"
OMIT = {"底稿与来源索引", "资料来源附注", "资料来源附注（脚注表）", "修订记录"}
RATING_NOTES = (
    ("重点推进", "最符合当前投资偏好，应尽快推进。"),
    ("研究推进", "与当前投资策略基本相容，仍需主动补充调研。"),
    ("保持跟踪", "暂不符合当前投资策略，以观察为主。"),
)
TABLES = {
    "字段": ((28, 72), ()), "核心成员": ((26, 74), ()),
    "产品线": ((20, 39, 41), ()), "主要竞争者": ((23, 40, 37), ()),
    "对应业务": ((23, 39, 20, 18), (2, 3)), "期间／口径": ((42, 20, 16, 22), (1, 2, 3)),
    "指标": ((30, 49, 21), (2,)), "评估理由": ((83, 17), ()),
    "类型": ((24, 76), ()), "事项": ((25, 75), ()),
}


def w(tag):
    return f"{{{W}}}{tag}"


def node(tag, **attrs):
    return etree.Element(w(tag), {w(key): str(value) for key, value in attrs.items()})


def add(parent, tag, **attrs):
    el = node(tag, **attrs)
    parent.append(el)
    return el


def replace(parent, tag, values=(), *, first=False):
    old = parent.find(w(tag))
    if old is not None:
        parent.remove(old)
    group = node(tag)
    group.extend(values)
    if first:
        parent.insert(0, group)
    else:
        parent.append(group)
    return group


def text(element):
    return "".join(element.xpath(".//w:t/text()", namespaces=NS))


def run_format(element, face, points, bold=False, latin="Arial"):
    group = replace(element, "rPr", first=element.tag == w("r"))
    add(group, "rFonts", ascii=latin, hAnsi=latin, eastAsia=face, cs=latin)
    if bold:
        add(group, "b")
        add(group, "bCs")
    add(group, "color", val=BLACK)
    add(group, "sz", val=round(points * 2))
    add(group, "szCs", val=round(points * 2))


def add_text(paragraph, value, face=BODY_FONT, points=9, bold=False, latin="Arial"):
    run = add(paragraph, "r")
    run_format(run, face, points, bold, latin)
    term = add(run, "t")
    term.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    term.text = value


def spacing(before=0, after=0, line=324):
    return node("spacing", before=before, after=after, line=line, lineRule="auto")


def pandoc(arguments, source, directory):
    return subprocess.run(
        ["pandoc", *arguments], input=source, cwd=directory, check=True, capture_output=True
    ).stdout


def load_source(args):
    source_bytes = args.source.read_bytes()
    markdown = source_bytes.decode("utf-8-sig")
    match = re.match(r"\A---\r?\n(.*?)\r?\n---(?:\r?\n|\Z)", markdown, re.S)
    meta = yaml.safe_load(match[1]) if match else {}
    if not isinstance(meta, dict) or meta.get("report_type", "初筛报告") != "初筛报告":
        raise ValueError("源文件必须是初筛报告 Markdown。")
    if meta.get("document_type", "report") != "report" or meta.get("canonical") is False:
        raise ValueError("请选择报告主稿。")
    actual = []
    for key, supplied in (("project_name", args.project), ("version", args.version), ("status", args.status)):
        existing = str(meta.get(key) or "").strip()
        if existing and supplied and supplied != existing:
            raise ValueError(f"{key} 与 Markdown 源稿冲突。")
        if not (existing or supplied):
            raise ValueError(f"源稿没有 {key}，请按已核实的版本记录提供对应参数。")
        actual.append(existing or supplied)
    project, version, status = actual
    if not re.fullmatch(r"v\d+(?:\.\d+)+", version):
        raise ValueError("无法识别源报告版本。")
    if status not in {"待复核", "已复核", "已冻结"}:
        raise ValueError("源报告状态须为待复核、已复核或已冻结。")
    if args.trigger == "review" and status == "待复核":
        raise ValueError("自动导出要求先登记用户复核通过。待复核稿仅能按用户明确要求导出。")
    legal_project = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", project).strip(" .")
    if not legal_project:
        raise ValueError("项目名为空。")
    # Existing formal drafts always use 已复核/待复核. A separately frozen master keeps that status.
    filename = f"{legal_project}_初筛报告_{version}_{status}.docx"
    return source_bytes, markdown[match.end():] if match else markdown, filename, project, version, status


def ast_text(value):
    if isinstance(value, list):
        return "".join(ast_text(x) for x in value)
    if not isinstance(value, dict):
        return ""
    if value.get("t") == "Str":
        return value["c"]
    if value.get("t") in {"Space", "SoftBreak"}:
        return " "
    return ast_text(value.get("c"))


def clean(value):
    if isinstance(value, list):
        return [clean(x) for x in value if not isinstance(x, dict) or x.get("t") != "Note"]
    if not isinstance(value, dict):
        return value
    if value.get("t") == "Image":
        url = value["c"][2][0]
        parts = urlsplit(url)
        if (parts.scheme and not Path(url).is_absolute()) or url.startswith(("//", "\\\\")):
            raise ValueError("请将报告图片保存为本地文件后再导出。")
    if value.get("t") == "Str":
        return {"t": "Str", "c": value["c"].replace("⭐", "★")}
    return {key: clean(item) for key, item in value.items()}


def filter_report(markdown, directory):
    ast = json.loads(pandoc(["-f", "markdown-smart", "-t", "json"], markdown.encode("utf-8"), directory))
    ast["meta"] = {}
    result, omit_level = [], None
    for block in ast["blocks"]:
        if block["t"] == "Header":
            level, _, inlines = block["c"]
            if omit_level is not None and level <= omit_level:
                omit_level = None
            title = ast_text(inlines).strip()
            if title in OMIT and omit_level is None:
                omit_level = level
        if omit_level is None:
            result.append(clean(block))
    if not result or result[0]["t"] != "Header" or result[0]["c"][0] != 1 or "初筛报告" not in ast_text(result[0]["c"][2]):
        raise ValueError("请选择以一级报告标题开头的初筛 Markdown 主稿。")
    ast["blocks"] = result
    return json.dumps(ast, ensure_ascii=False).encode("utf-8")


def set_styles(styles):
    ids = {el.get(w("styleId")): el for el in styles.findall(w("style"))}
    for name, size, before, after in (
        ("Title", 24, 0, 180), ("Heading1", 16, 240, 140),
        ("Heading2", 13, 180, 100), ("Heading3", 11, 145, 78),
    ):
        st = ids[name]
        fields = [node("keepNext"), node("keepLines"), node("widowControl"), spacing(before, after, 285)]
        if name == "Title":
            fields.append(node("jc", val="center"))
        else:
            fields.append(node("outlineLvl", val=int(name[-1]) - 1))
        replace(st, "pPr", fields)
        run_format(st, TITLE_FONT, size, True)
    for name in ("Normal", "BodyText", "FirstParagraph", "Compact"):
        if name not in ids:
            continue
        st = ids[name]
        replace(st, "pPr", [node("widowControl"), spacing(after=125)])
        run_format(st, BODY_FONT, 11)
    note_style = add(styles, "style", type="paragraph", styleId="BPTypeNote", customStyle=1)
    add(note_style, "name", val="项目类型说明")
    add(note_style, "basedOn", val="BodyText")
    add_style = add(note_style, "pPr")
    add_style.extend([node("keepLines"), node("widowControl"), spacing(after=60, line=275), node("ind", firstLineChars=200, firstLine=360)])
    run_format(note_style, NOTE_FONT, 9, latin="KaiTi")


def add_heading_numbers(numbering):
    abstracts = numbering.findall(w("abstractNum"))
    instances = numbering.findall(w("num"))
    abstract_id = 1 + max((int(x.get(w("abstractNumId"))) for x in abstracts), default=0)
    num_id = 1 + max((int(x.get(w("numId"))) for x in instances), default=0)
    abstract = node("abstractNum", abstractNumId=abstract_id)
    numbering.insert(len(numbering) - len(instances), abstract)
    add(abstract, "multiLevelType", val="multilevel")
    for level, (fmt, label, size) in enumerate((
        ("chineseCounting", "%1、", 16), ("chineseCounting", "（%2）、", 13), ("decimal", "%3、", 11)
    )):
        lvl = add(abstract, "lvl", ilvl=level)
        add(lvl, "start", val=1)
        add(lvl, "numFmt", val=fmt)
        if level:
            add(lvl, "lvlRestart", val=level)
        add(lvl, "pStyle", val=f"Heading{level + 1}")
        add(lvl, "suff", val="nothing")
        add(lvl, "lvlText", val=label)
        add(lvl, "lvlJc", val="left")
        add(add(lvl, "pPr"), "ind", left=0, hanging=0)
        run_format(lvl, TITLE_FONT, size, True)
    add(add(numbering, "num", numId=num_id), "abstractNumId", val=abstract_id)
    return num_id


def format_paragraphs(body, num_id):
    category = None
    end_of_category = None
    for p in body.findall(w("p")):
        style = p.find("./" + w("pPr") + "/" + w("pStyle"))
        name = style.get(w("val")) if style is not None else ""
        if name.startswith("Heading"):
            level = int(name.removeprefix("Heading"))
            if level == 1:
                style.set(w("val"), "Title")
            else:
                heading = min(3, level - 1)
                style.set(w("val"), f"Heading{heading}")
                pp = p.find(w("pPr"))
                num = node("numPr")
                add(num, "ilvl", val=heading - 1)
                add(num, "numId", val=num_id)
                pp.insert(1, num)
            if category is not None and end_of_category is None:
                end_of_category = p
            if text(p).strip() == "项目归类":
                category = p
        elif text(p).strip():
            fields = [
                node("pStyle", val="BodyText"), node("widowControl"), spacing(after=125),
                node("ind", firstLineChars=200, firstLine=440),
            ]
            replace(p, "pPr", fields, first=True)
    if category is None:
        raise ValueError("源报告缺少“项目归类”章节。")
    if end_of_category is None:
        end_of_category = body.find(w("sectPr"))
    at = body.index(end_of_category)
    # Pandoc anchors are before the respective headings. Keep the notes inside the category anchor.
    while at and body[at - 1].tag in {w("bookmarkStart"), w("bookmarkEnd")}:
        at -= 1
    for i, (label, description) in enumerate(RATING_NOTES):
        p = node("p")
        pp = add(p, "pPr")
        add(pp, "pStyle", val="BPTypeNote")
        if i < len(RATING_NOTES) - 1:
            add(pp, "keepNext")
        if i == 0:
            pp.append(spacing(before=120, after=60, line=275))
        add_text(p, label, NOTE_FONT, 9, True, "KaiTi")
        add_text(p, "　" + description, NOTE_FONT, 9, latin="KaiTi")
        body.insert(at + i, p)


def format_table(table):
    rows = table.findall(w("tr"))
    columns = [text(c) for c in rows[0].findall(w("tc"))]
    key = columns[0]
    ratios, numeric = TABLES.get(key, ((), ()))
    if len(ratios) != len(columns):
        ratios = (1,) * len(columns)
        numeric = tuple(i for i, c in enumerate(columns) if any(k in c for k in ("金额", "利润", "增幅", "比例", "总产出")))
    widths = [round(WIDTH * x / sum(ratios)) for x in ratios]
    widths[-1] = WIDTH - sum(widths[:-1])
    pr = replace(table, "tblPr", first=True)
    add(pr, "tblW", type="dxa", w=WIDTH)
    add(pr, "jc", val="left")
    borders = add(pr, "tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        if edge in {"top", "bottom", "insideH"}:
            add(borders, edge, val="single", sz=4, color=LINE, space=0)
        else:
            add(borders, edge, val="nil")
    add(pr, "tblLayout", type="fixed")
    pad = add(pr, "tblCellMar")
    for edge, val in (("top", 65), ("left", 122), ("bottom", 65), ("right", 122)):
        add(pad, edge, type="dxa", w=val)
    grid = table.find(w("tblGrid"))
    grid.clear()
    for width in widths:
        add(grid, "gridCol", w=width)
    for index, row in enumerate(rows):
        rp = replace(row, "trPr", first=True)
        add(rp, "cantSplit")
        if index == 0:
            add(rp, "tblHeader")
        col = 0
        for cell in row.findall(w("tc")):
            old = cell.find(w("tcPr"))
            span = old.find(w("gridSpan")) if old is not None else None
            ncol = int(span.get(w("val"))) if span is not None else 1
            selected = key == "类型" and col == 0 and text(cell) in dict(RATING_NOTES)
            # Keep Pandoc's gridSpan/vMerge for any merged cells.
            tc = old if old is not None else replace(cell, "tcPr", first=True)
            for field in ("tcW", "shd", "vAlign"):
                previous = tc.find(w(field))
                if previous is not None:
                    tc.remove(previous)
            tc.insert(0, node("tcW", type="dxa", w=sum(widths[col:col + ncol])))
            add(tc, "shd", val="clear", color="auto", fill=GRAY if not index or selected else "FFFFFF")
            add(tc, "vAlign", val="center" if key in {"评估理由", "字段"} else "top")
            for p in cell.findall(w("p")):
                jc = "center" if key == "评估理由" and col == 1 else "right" if col in numeric else "left"
                fields = [node("keepLines"), node("widowControl"), spacing(line=276), node("jc", val=jc)]
                if not index:
                    fields.insert(0, node("keepNext"))
                replace(p, "pPr", fields, first=True)
                for run in p.findall(w("r")):
                    props = run.find(w("rPr"))
                    is_bold = props is not None and props.find(w("b")) is not None
                    face = TITLE_FONT if not index else "SimHei" if selected else BODY_FONT
                    latin = "Segoe UI Symbol" if "★" in text(run) else "Arial"
                    run_format(run, face, 10.5, not index or selected or is_bold, latin)
            col += ncol


def page_parts(entries, body, title, classification):
    section = body.find(w("sectPr"))
    section.clear()
    relationships = etree.fromstring(entries["word/_rels/document.xml.rels"])
    contents = etree.fromstring(entries["[Content_Types].xml"])
    hp = node("hdr")
    p = add(hp, "p")
    pp = add(p, "pPr")
    add(add(pp, "pBdr"), "bottom", val="single", sz=4, color=LINE, space=5)
    add(add(pp, "tabs"), "tab", val="right", pos=WIDTH)
    add_text(p, title, TITLE_FONT)
    if classification:
        add(add(p, "r"), "tab")
        add_text(p, classification, TITLE_FONT)
    fp = node("ftr")
    p = add(fp, "p")
    add(add(p, "pPr"), "jc", val="center")
    add_text(p, "第 ")
    add_text(add(p, "fldSimple", instr=" PAGE "), "1")
    add_text(p, " 页")
    for part, xml in (("header", hp), ("footer", fp)):
        ident, name = f"rIdBP{part.title()}", f"bp-{part}.xml"
        add(section, f"{part}Reference", type="default").set(f"{{{R}}}id", ident)
        etree.SubElement(relationships, f"{{{REL}}}Relationship", Id=ident, Type=f"{R}/{part}", Target=name)
        etree.SubElement(contents, f"{{{CT}}}Override", PartName=f"/word/{name}",
                         ContentType=f"application/vnd.openxmlformats-officedocument.wordprocessingml.{part}+xml")
        entries["word/" + name] = dump(xml)
    add(section, "pgSz", w=PAGE_WIDTH, h=PAGE_HEIGHT, orient="portrait")
    add(section, "pgMar", top=1134, bottom=1134, left=SIDE_MARGIN, right=SIDE_MARGIN, header=520, footer=520, gutter=0)
    for el in list(relationships):
        if el.get("Type", "").endswith("/footnotes"):
            relationships.remove(el)
    for el in list(contents):
        if el.get("PartName") == "/word/footnotes.xml":
            contents.remove(el)
    for name in ("word/footnotes.xml", "word/_rels/footnotes.xml.rels"):
        entries.pop(name, None)
    settings = etree.fromstring(entries["word/settings.xml"])
    for el in settings.findall(w("footnotePr")):
        settings.remove(el)
    entries.update({"[Content_Types].xml": dump(contents), "word/_rels/document.xml.rels": dump(relationships), "word/settings.xml": dump(settings)})


def dump(element):
    return etree.tostring(element, encoding="utf-8", xml_declaration=True)


def format_docx(blob):
    with ZipFile(BytesIO(blob)) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    doc, styles, nums, fonts = (
        etree.fromstring(entries[f"word/{key}.xml"]) for key in ("document", "styles", "numbering", "fontTable")
    )
    body = doc.find(w("body"))
    original = text(body)
    title = text(body.find(w("p")))
    classification = ""
    for row in body.xpath("./w:tbl/w:tr", namespaces=NS):
        cells = row.findall(w("tc"))
        if len(cells) == 2 and text(cells[0]) == "保密等级":
            classification = text(cells[1])
    set_styles(styles)
    format_paragraphs(body, add_heading_numbers(nums))
    for table in body.findall(w("tbl")):
        format_table(table)
    add(add(fonts, "font", name=NOTE_FONT), "altName", val="KaiTi")
    page_parts(entries, body, title, classification)
    for xml in (doc, styles, nums):
        for color in xml.iter(w("color")):
            color.attrib.clear()
            color.set(w("val"), BLACK)
    without_notes = "".join(
        text(x) for x in body if not x.xpath("./w:pPr/w:pStyle[@w:val='BPTypeNote']", namespaces=NS)
    )
    if original != without_notes or list(doc.iter(w("footnoteReference"))) or re.search(r"\[\^\d+\]", without_notes):
        raise ValueError("Word 内容对照或脚注检查未通过，请核对源稿。")
    for name, element in (("document", doc), ("styles", styles), ("numbering", nums), ("fontTable", fonts)):
        entries[f"word/{name}.xml"] = dump(element)
    output = BytesIO()
    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            archive.writestr(name, data)
    return output.getvalue()


def export(args):
    args.source = args.source.resolve(strict=True)
    source_bytes, markdown, filename, project, version, status = load_source(args)
    input_ast = filter_report(markdown, args.source.parent)
    blob = format_docx(pandoc(["-f", "json", "-t", "docx", "-o", "-"], input_ast, args.source.parent))
    if args.source.read_bytes() != source_bytes:
        raise ValueError("导出过程中源报告有变化，请核对后重试。")
    if args.output_dir is not None:
        folder = args.output_dir.resolve()
        folder.mkdir(parents=True, exist_ok=True)
    else:
        root = args.source.parent / "90_按需导出"
        stamp = datetime.now().strftime("%Y%m%d-%H%M")
        folder = root / stamp
        n = 2
        while True:
            try:
                folder.mkdir(parents=True)
                break
            except FileExistsError:
                folder = root / f"{stamp}-{n:02d}"
                n += 1
    target = folder / filename
    with target.open("xb") as f:
        f.write(blob)
    return {
        "source": str(args.source), "docx": str(target), "version": version, "status": status,
        "source_sha256": sha256(source_bytes).hexdigest(), "docx_sha256": sha256(blob).hexdigest(),
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("--trigger", choices=("review", "requested"), required=True)
    parser.add_argument("--project")
    parser.add_argument("--version")
    parser.add_argument("--status", choices=("待复核", "已复核", "已冻结"))
    parser.add_argument("--output-dir", type=Path)
    args = parser.parse_args()
    try:
        result = export(args)
    except subprocess.CalledProcessError as exc:
        parser.exit(1, f"Pandoc 转换失败：{exc.stderr.decode('utf-8', errors='replace').strip()}\n")
    except (OSError, ValueError, yaml.YAMLError) as exc:
        parser.exit(1, f"{exc}\n")
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()

