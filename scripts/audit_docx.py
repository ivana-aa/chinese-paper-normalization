"""Read-only DOCX formatting inventory; does not validate rendered pages."""

import argparse
import hashlib
import json
import re
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path
from zipfile import ZipFile


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}
W = "{" + NS["w"] + "}"
CJK = r"\u3400-\u9fff"
BOUNDARY_SPACE = re.compile(
    rf"[{CJK}][ \u00a0\u3000]+(?=[A-Za-z0-9])|[A-Za-z0-9][ \u00a0\u3000]+(?=[{CJK}])"
)
HEADING = re.compile(r"^([1-9]\d?(?:\.\d+){0,2})(?:\.(?!\d))?\s*([\u3400-\u9fffA-Za-z].*)$")


def attrs(element):
    return {key.rsplit("}", 1)[-1]: value for key, value in element.attrib.items()}


def merge_properties(target, element):
    if element is not None:
        for child in element:
            target.setdefault(child.tag.rsplit("}", 1)[-1], {}).update(attrs(child))
    return target


def merge_profile(base, changes):
    result = dict(base)
    for key, value in changes.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = merge_profile(result[key], value)
        else:
            result[key] = value
    return result


def text_of(paragraph):
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


class Styles:
    def __init__(self, root):
        self.root = root
        self.by_id = {s.get(W + "styleId"): s for s in root.findall("w:style", NS)}
        self.default_id = next(
            (key for key, s in self.by_id.items()
             if s.get(W + "type") == "paragraph" and s.get(W + "default") == "1"),
            "Normal",
        )

    def chain(self, style_id, seen=None):
        seen = set() if seen is None else seen
        if style_id in seen or style_id not in self.by_id:
            return []
        seen.add(style_id)
        style = self.by_id[style_id]
        base = style.find("w:basedOn", NS)
        return (self.chain(base.get(W + "val"), seen) if base is not None else []) + [style]

    def paragraph_style(self, paragraph):
        node = paragraph.find("w:pPr/w:pStyle", NS)
        return node.get(W + "val") if node is not None else self.default_id

    def properties(self, paragraph, run=None):
        kind = "pPr" if run is None else "rPr"
        result = {}
        defaults = self.root.find(f"w:docDefaults/w:{kind}Default/w:{kind}", NS)
        merge_properties(result, defaults)
        for style in self.chain(self.paragraph_style(paragraph)):
            merge_properties(result, style.find("w:" + kind, NS))
        if run is None:
            merge_properties(result, paragraph.find("w:pPr", NS))
        else:
            character = run.find("w:rPr/w:rStyle", NS)
            if character is not None:
                for style in self.chain(character.get(W + "val")):
                    merge_properties(result, style.find("w:rPr", NS))
            merge_properties(result, run.find("w:rPr", NS))
        return result

    def name(self, paragraph):
        style = self.by_id.get(self.paragraph_style(paragraph))
        name = style.find("w:name", NS) if style is not None else None
        return name.get(W + "val", "") if name is not None else ""


def classify(paragraph, styles, in_body, in_references):
    text = text_of(paragraph).strip()
    style = styles.name(paragraph).lower()
    if style.startswith(("toc", "目录")):
        return "toc"
    if "code" in style or "代码" in style:
        return "code"
    if paragraph.find(".//m:oMath", NS) is not None:
        return "equation"
    if re.match(r"^[图表]\s*\d+[-－]\d+", text):
        return "caption"
    if style.startswith(("heading", "标题")):
        level = re.search(r"([123])$", style)
        if level:
            return "heading" + level.group(1)
    match = HEADING.match(text)
    if match and len(text) <= 100:
        return "heading" + str(match.group(1).count(".") + 1)
    if in_references and re.match(r"^\[\d+\]", text):
        return "references"
    return "body" if in_body else "front_matter"


def inspect_paragraph(paragraph, role, styles, profile, location, issues):
    text = text_of(paragraph)
    if not text.strip():
        return
    expected = profile.get(role)

    def note(check, expected_value, actual):
        issues.append({"location": location, "role": role, "check": check,
                       "expected": expected_value, "actual": actual, "text": text[:90]})

    if role not in {"code", "equation", "references", "front_matter", "toc"}:
        spaces = [m.group(0) for m in BOUNDARY_SPACE.finditer(text)]
        if spaces and profile.get("spaces", {}).get("remove_cjk_latin_boundaries", True):
            note("cjk_latin_space_candidate", "复核并清理正文交界空格", spaces)
    if expected is None:
        return

    pprops = styles.properties(paragraph)
    alignment = pprops.get("jc", {}).get("val", "left")
    if alignment != expected.get("alignment", alignment):
        note("alignment", expected["alignment"], alignment)
    if role == "body":
        spacing = pprops.get("spacing", {})
        line = spacing.get("line", "240")
        rule = spacing.get("lineRule", "auto")
        if rule != "auto" or float(line) != expected["line_spacing"] * 240:
            note("line_spacing", expected["line_spacing"], {"line": line, "rule": rule})
        ind = pprops.get("ind", {})
        if "firstLineChars" in ind:
            good_indent = ind["firstLineChars"] == str(expected["first_line_chars"] * 100)
        else:
            good_indent = ind.get("firstLine") == str(int(expected["size_pt"] * 20 * expected["first_line_chars"]))
        if not good_indent:
            note("first_line_indent", f"{expected['first_line_chars']}字符", ind)
        for label in ("before", "after"):
            if spacing.get(label, "0") != "0" or spacing.get(label + "Lines", "0") != "0":
                note("space_" + label, "0", spacing)

    for index, run in enumerate(paragraph.findall(".//w:r", NS), 1):
        value = "".join(t.text or "" for t in run.findall("w:t", NS))
        if not value:
            continue
        rprops = styles.properties(paragraph, run)
        fonts = rprops.get("rFonts", {})
        for label, pattern, key, profile_key in (
            ("chinese_font", rf"[{CJK}]", "eastAsia", "chinese_font"),
            ("latin_font", r"[A-Za-z0-9]", "ascii", "latin_font"),
        ):
            if re.search(pattern, value):
                actual = fonts.get(key)
                if fonts.get(key + "Theme"):
                    note(label + "_theme_review", expected[profile_key], fonts)
                elif actual is None:
                    note(label + "_unresolved", expected[profile_key], "字体无法从显式属性确定")
                elif actual != expected[profile_key]:
                    note(label, expected[profile_key], actual)
        size = rprops.get("sz", {}).get("val")
        if size is None:
            note("size_unresolved", expected["size_pt"], None)
        elif float(size) / 2 != expected["size_pt"]:
            note("size", expected["size_pt"], float(size) / 2)
        if expected.get("bold") and rprops.get("b", {}).get("val", "1" if "b" in rprops else "0") in {"0", "false", "off"}:
            note("heading_bold", True, False)
        char_spacing = rprops.get("spacing", {}).get("val", "0")
        if char_spacing != "0" and role != "code":
            note("character_spacing", "0", char_spacing)


def inspect_table(table, number, styles, profile, issues):
    records = []
    for edge in table.findall("w:tblPr/w:tblBorders/*", NS):
        records.append({"scope": "table", "edge": edge.tag.rsplit("}", 1)[-1], **attrs(edge)})
    for row_index, row in enumerate(table.findall("w:tr", NS), 1):
        for cell_index, cell in enumerate(row.findall("w:tc", NS), 1):
            location = f"表{number}/行{row_index}/列{cell_index}"
            valign = cell.find("w:tcPr/w:vAlign", NS)
            actual = valign.get(W + "val") if valign is not None else "top"
            if actual != profile["table"]["vertical_alignment"]:
                issues.append({"location": location, "check": "table_vertical_alignment",
                               "expected": "center", "actual": actual})
            for edge in cell.findall("w:tcPr/w:tcBorders/*", NS):
                records.append({"scope": location, "edge": edge.tag.rsplit("}", 1)[-1], **attrs(edge)})
            for paragraph in cell.findall("w:p", NS):
                inspect_paragraph(paragraph, "table", styles, profile, location, issues)
    style_id = table.find("w:tblPr/w:tblStyle", NS)
    vertical = [r for r in records if r["edge"] in {"left", "right", "insideV", "start", "end"}
                and r.get("val", "nil") not in {"nil", "none"}]
    if vertical:
        issues.append({"location": f"表{number}", "check": "visible_vertical_border_candidate",
                       "expected": "无竖线", "actual": vertical})
    return {"table": number, "style_id": style_id.get(W + "val") if style_id is not None else None,
            "direct_border_records": records,
            "requires_visual_review": True,
            "note": "完整三线表判定需结合条件表格样式、单元格覆盖和渲染结果；此处只盘点显式边框。"}


def audit(path, profile, scope="auto"):
    original_hash = hashlib.sha256(path.read_bytes()).hexdigest()
    with ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
        styles = Styles(ET.fromstring(archive.read("word/styles.xml")))
        issues, table_inventory, headings = [], [], []
        in_body, in_references = scope == "all", False
        paragraph_no, table_no = 0, 0
        for node in root.find("w:body", NS):
            if node.tag == W + "p":
                paragraph_no += 1
                value = text_of(node).strip()
                if value == "参考文献":
                    in_references = True
                if value == "致谢":
                    in_body, in_references = False, False
                role = classify(node, styles, in_body, in_references)
                if role.startswith("heading"):
                    in_body = True
                    headings.append({"paragraph": paragraph_no, "role": role, "text": value})
                if not in_references or role == "references":
                    inspect_paragraph(node, role, styles, profile, f"正文段落{paragraph_no}", issues)
            elif node.tag == W + "tbl":
                table_no += 1
                # Front matter can contain layout tables that are not academic data tables.
                if in_body and not in_references:
                    table_inventory.append(inspect_table(node, table_no, styles, profile, issues))

        sections = []
        for number, section in enumerate(root.findall(".//w:sectPr", NS), 1):
            size = section.find("w:pgSz", NS)
            margins = section.find("w:pgMar", NS)
            info = {"section": number, "size": attrs(size) if size is not None else {},
                    "margins": attrs(margins) if margins is not None else {}}
            sections.append(info)
            expected_page = profile["page"]
            expected = {"w": expected_page["width_cm"], "h": expected_page["height_cm"]}
            for key, cm in expected.items():
                actual = info["size"].get(key)
                if actual is None or abs(float(actual) - cm / 2.54 * 1440) > 4:
                    issues.append({"location": f"分节{number}", "check": "page_size_" + key,
                                   "expected_cm": cm, "actual_twips": actual})
            for key, cm in expected_page["margins_cm"].items():
                actual = info["margins"].get(key)
                if actual is None or abs(float(actual) - cm / 2.54 * 1440) > 4:
                    issues.append({"location": f"分节{number}", "check": "margin_" + key,
                                   "expected_cm": cm, "actual_twips": actual})

        media = {name: hashlib.sha256(archive.read(name)).hexdigest()
                 for name in archive.namelist() if name.startswith("word/media/") and not name.endswith("/")}
        fields = [node.text or "" for node in root.findall(".//w:instrText", NS)]
        report = {"input": str(path.resolve()), "sha256": original_hash, "scope": scope,
                  "issue_count": len(issues), "checks": dict(Counter(i["check"] for i in issues)),
                  "issues": issues, "headings": headings, "sections": sections,
                  "tables": table_inventory, "media_sha256": media,
                  "native_math_count": len(root.findall(".//m:oMath", NS)), "field_instructions": fields,
                  "visual_qa": "未执行：此脚本只读OOXML，不能证明分页、图像位置、数学字形或实际字体替代。",
                  "coverage_notes": "自动角色分类属于候选；前置页、嵌套表、文本框、目录和全部参考文献标点仍需人工核查。"}
    if hashlib.sha256(path.read_bytes()).hexdigest() != original_hash:
        raise RuntimeError("输入文件在审计期间发生变化，请复查文件是否被其他应用保存。")
    return report


def main():
    parser = argparse.ArgumentParser(description="只读盘点中文论文DOCX格式，不修改输入文件。")
    parser.add_argument("input", type=Path)
    parser.add_argument("--report", type=Path, help="将UTF-8 JSON审计结果写入指定位置")
    parser.add_argument("--profile", type=Path, help="JSON局部覆盖默认配置；学校模板和当次要求优先")
    parser.add_argument("--scope", choices=["auto", "all"], default="auto",
                        help="auto从首个编号标题起审计正文；all用于无封面的小节片段")
    args = parser.parse_args()
    default_path = Path(__file__).resolve().parents[1] / "references" / "default-profile.json"
    profile = json.loads(default_path.read_text(encoding="utf-8"))
    if args.profile:
        profile = merge_profile(profile, json.loads(args.profile.read_text(encoding="utf-8-sig")))
    report = audit(args.input, profile, args.scope)
    rendered = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report:
        if args.report.resolve() == args.input.resolve():
            parser.error("审计结果路径不能与输入DOCX相同")
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(rendered, encoding="utf-8")
        print(json.dumps({"report": str(args.report), "issue_count": report["issue_count"],
                          "checks": report["checks"], "visual_qa": "未执行"}, ensure_ascii=False))
    else:
        print(rendered)


if __name__ == "__main__":
    main()
