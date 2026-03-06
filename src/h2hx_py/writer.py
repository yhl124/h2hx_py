from __future__ import annotations

from pathlib import Path
import zipfile

from lxml import etree as ET

from .model import (
    BorderFill,
    BorderLine,
    CharShape,
    ColumnsDef,
    Equation,
    FieldBegin,
    FieldEnd,
    HwpDocument,
    NoteShape,
    PageBorderFill,
    PageDef,
    PageHide,
    PageNum,
    ParaShape,
    Run,
    Section,
    SectionDef,
    Style,
    TabDef,
    Table,
    TableCell,
)

XML_NS = "http://www.w3.org/XML/1998/namespace"

NS = {
    "ha": "http://www.hancom.co.kr/hwpml/2011/app",
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hp10": "http://www.hancom.co.kr/hwpml/2016/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "hhs": "http://www.hancom.co.kr/hwpml/2011/history",
    "hm": "http://www.hancom.co.kr/hwpml/2011/master-page",
    "hpf": "http://www.hancom.co.kr/schema/2011/hpf",
    "dc": "http://purl.org/dc/elements/1.1/",
    "opf": "http://www.idpf.org/2007/opf/",
    "ooxmlchart": "http://www.hancom.co.kr/hwpml/2016/ooxmlchart",
    "hwpunitchar": "http://www.hancom.co.kr/hwpml/2016/HwpUnitChar",
    "epub": "http://www.idpf.org/2007/ops",
    "config": "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
    "ocf": "urn:oasis:names:tc:opendocument:xmlns:container",
    "odf": "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
    "rdf": "http://www.w3.org/1999/02/22-rdf-syntax-ns#",
    "hv": "http://www.hancom.co.kr/hwpml/2011/version",
    "pkg": "http://www.hancom.co.kr/hwpml/2016/meta/pkg#",
}

DOC_NSMAP = {
    "ha": NS["ha"],
    "hp": NS["hp"],
    "hp10": NS["hp10"],
    "hs": NS["hs"],
    "hc": NS["hc"],
    "hh": NS["hh"],
    "hhs": NS["hhs"],
    "hm": NS["hm"],
    "hpf": NS["hpf"],
    "dc": NS["dc"],
    "opf": NS["opf"],
    "ooxmlchart": NS["ooxmlchart"],
    "hwpunitchar": NS["hwpunitchar"],
    "epub": NS["epub"],
    "config": NS["config"],
}


def write_hwpx(document: HwpDocument, destination: str | Path) -> Path:
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)

    files = {
        "mimetype": b"application/hwp+zip",
        "version.xml": _xml_bytes(_build_version_xml(document)),
        "Contents/content.hpf": _xml_bytes(_build_content_hpf(document)),
        "Contents/header.xml": _xml_bytes(_build_header_xml(document)),
        "settings.xml": _xml_bytes(_build_settings_xml(document)),
        "META-INF/container.xml": _xml_bytes(_build_container_xml()),
        "META-INF/manifest.xml": _xml_bytes(_build_manifest_xml(document)),
        "Preview/PrvText.txt": document.metadata.preview_text.encode("utf-8"),
    }
    for section in document.sections:
        files[f"Contents/section{section.index}.xml"] = _xml_bytes(_build_section_xml(section))

    with zipfile.ZipFile(destination, "w") as archive:
        archive.writestr("mimetype", files.pop("mimetype"), compress_type=zipfile.ZIP_DEFLATED)
        for name, payload in files.items():
            archive.writestr(name, payload, compress_type=zipfile.ZIP_DEFLATED)

    return destination


def _build_version_xml(document: HwpDocument) -> ET.Element:
    major, minor, micro, build = document.version
    return ET.Element(
        _q("hv", "HCFVersion"),
        {
            "tagetApplication": "WORDPROCESSOR",
            "major": str(major),
            "minor": str(minor),
            "micro": str(micro),
            "buildNumber": str(build),
            "os": "1",
            "xmlVersion": "1.4",
            "application": "Hancom Office Hangul",
            "appVersion": "9, 1, 1, 5656 WIN32LEWindows_Unknown_Version",
        },
        nsmap={"hv": NS["hv"]},
    )


def _build_content_hpf(document: HwpDocument) -> ET.Element:
    root = ET.Element(_q("opf", "package"), {"version": "", "unique-identifier": "", "id": ""}, nsmap=DOC_NSMAP)
    metadata = ET.SubElement(root, _q("opf", "metadata"))
    ET.SubElement(metadata, _q("opf", "title")).text = document.metadata.title
    ET.SubElement(metadata, _q("opf", "language")).text = "ko"
    for name, value in (
        ("creator", document.metadata.author),
        ("subject", document.metadata.subject),
        ("description", document.metadata.comments),
        ("lastsaveby", document.metadata.last_saved_by),
        ("CreatedDate", document.metadata.created_at),
        ("ModifiedDate", document.metadata.modified_at),
        ("date", document.metadata.date_string),
        ("keyword", document.metadata.keywords),
    ):
        node = ET.SubElement(metadata, _q("opf", "meta"), {"name": name, "content": "text"})
        node.text = value

    manifest = ET.SubElement(root, _q("opf", "manifest"))
    ET.SubElement(manifest, _q("opf", "item"), {"id": "header", "href": "Contents/header.xml", "media-type": "application/xml"})
    for section in document.sections:
        ET.SubElement(
            manifest,
            _q("opf", "item"),
            {
                "id": f"section{section.index}",
                "href": f"Contents/section{section.index}.xml",
                "media-type": "application/xml",
            },
        )
    ET.SubElement(manifest, _q("opf", "item"), {"id": "settings", "href": "settings.xml", "media-type": "application/xml"})

    spine = ET.SubElement(root, _q("opf", "spine"))
    ET.SubElement(spine, _q("opf", "itemref"), {"idref": "header", "linear": "yes"})
    for section in document.sections:
        ET.SubElement(spine, _q("opf", "itemref"), {"idref": f"section{section.index}", "linear": "yes"})
    return root


def _build_header_xml(document: HwpDocument) -> ET.Element:
    root = ET.Element(_q("hh", "head"), {"version": "1.4", "secCnt": str(document.section_count)}, nsmap=DOC_NSMAP)
    ET.SubElement(
        root,
        _q("hh", "beginNum"),
        {"page": "1", "footnote": "1", "endnote": "1", "pic": "1", "tbl": "1", "equation": "1"},
    )
    ref_list = ET.SubElement(root, _q("hh", "refList"))
    _fontfaces(ref_list, document)
    _borderfills(ref_list, document.border_fills)
    _char_properties(ref_list, document.char_shapes)
    _tab_properties(ref_list, document.tab_defs)
    _para_properties(ref_list, document.para_shapes)
    _styles(ref_list, document.styles)

    compatible = ET.SubElement(root, _q("hh", "compatibleDocument"), {"targetProgram": "HWP201X"})
    ET.SubElement(compatible, _q("hh", "layoutCompatibility"))
    doc_option = ET.SubElement(root, _q("hh", "docOption"))
    ET.SubElement(doc_option, _q("hh", "linkinfo"), {"path": "", "pageInherit": "0", "footnoteInherit": "0"})
    ET.SubElement(root, _q("hh", "trackchageConfig"), {"flags": "56"})
    return root


def _build_section_xml(section: Section) -> ET.Element:
    root = ET.Element(_q("hs", "sec"), nsmap=DOC_NSMAP)
    for paragraph in section.paragraphs:
        root.append(_paragraph_xml(paragraph))
    return root


def _build_settings_xml(document: HwpDocument) -> ET.Element:
    root = ET.Element(_q("ha", "HWPApplicationSetting"), nsmap={"ha": NS["ha"]})
    ET.SubElement(
        root,
        _q("ha", "CaretPosition"),
        {
            "listIDRef": str(document.caret.list_id),
            "paraIDRef": str(document.caret.paragraph_id),
            "pos": str(document.caret.position_in_paragraph),
        },
    )
    return root


def _build_container_xml() -> ET.Element:
    root = ET.Element(_q("ocf", "container"), nsmap={"ocf": NS["ocf"], "hpf": NS["hpf"]})
    rootfiles = ET.SubElement(root, _q("ocf", "rootfiles"))
    ET.SubElement(rootfiles, _q("ocf", "rootfile"), {"full-path": "Contents/content.hpf", "media-type": "application/hwpml-package+xml"})
    ET.SubElement(rootfiles, _q("ocf", "rootfile"), {"full-path": "Preview/PrvText.txt", "media-type": "text/plain"})
    return root


def _build_manifest_xml(document: HwpDocument) -> ET.Element:
    return ET.Element(_q("odf", "manifest"), nsmap={"odf": NS["odf"]})

def _fontfaces(parent: ET.Element, document: HwpDocument) -> None:
    container = ET.SubElement(parent, _q("hh", "fontfaces"), {"itemCnt": str(len(document.font_faces))})
    for _, group_key, group_name in (
        ("ko_fonts", "hangul", "HANGUL"),
        ("en_fonts", "latin", "LATIN"),
        ("cn_fonts", "hanja", "HANJA"),
        ("jp_fonts", "japanese", "JAPANESE"),
        ("other_fonts", "other", "OTHER"),
        ("symbol_fonts", "symbol", "SYMBOL"),
        ("user_fonts", "user", "USER"),
    ):
        fonts = document.font_faces.get(group_key, [])
        fontface = ET.SubElement(container, _q("hh", "fontface"), {"lang": group_name, "fontCnt": str(len(fonts))})
        for font in fonts:
            ET.SubElement(fontface, _q("hh", "font"), {"id": str(font.id), "face": font.name, "type": "TTF", "isEmbedded": "0"})


def _borderfills(parent: ET.Element, border_fills: list[BorderFill]) -> None:
    container = ET.SubElement(parent, _q("hh", "borderFills"), {"itemCnt": str(len(border_fills))})
    for border_fill in border_fills:
        node = ET.SubElement(
            container,
            _q("hh", "borderFill"),
            {
                "id": str(border_fill.id),
                "threeD": "1" if border_fill.three_d else "0",
                "shadow": "1" if border_fill.shadow else "0",
                "centerLine": border_fill.center_line,
                "breakCellSeparateLine": "1" if border_fill.break_cell_separate_line else "0",
            },
        )
        ET.SubElement(
            node,
            _q("hh", "slash"),
            {
                "type": border_fill.slash_type,
                "Crooked": "1" if border_fill.slash_crooked else "0",
                "isCounter": "1" if border_fill.slash_counter else "0",
            },
        )
        ET.SubElement(
            node,
            _q("hh", "backSlash"),
            {
                "type": border_fill.backslash_type,
                "Crooked": "1" if border_fill.backslash_crooked else "0",
                "isCounter": "1" if border_fill.backslash_counter else "0",
            },
        )
        for tag, line in (
            ("leftBorder", border_fill.left_border),
            ("rightBorder", border_fill.right_border),
            ("topBorder", border_fill.top_border),
            ("bottomBorder", border_fill.bottom_border),
            ("diagonal", border_fill.diagonal_border),
        ):
            _border_line(node, tag, line)
        if border_fill.has_fill:
            fill_brush = ET.SubElement(node, _q("hc", "fillBrush"))
            attrs = {
                "faceColor": border_fill.fill_face_color,
                "hatchColor": border_fill.fill_hatch_color,
                "alpha": _alpha_text(border_fill.fill_alpha),
            }
            if border_fill.fill_hatch_style:
                attrs["hatchStyle"] = border_fill.fill_hatch_style
            ET.SubElement(fill_brush, _q("hc", "winBrush"), attrs)


def _char_properties(parent: ET.Element, char_shapes: list[CharShape]) -> None:
    container = ET.SubElement(parent, _q("hh", "charProperties"), {"itemCnt": str(len(char_shapes))})
    for char_shape in char_shapes:
        node = ET.SubElement(
            container,
            _q("hh", "charPr"),
            {
                "id": str(char_shape.id),
                "height": str(char_shape.basesize),
                "textColor": char_shape.text_color,
                "shadeColor": char_shape.shade_color,
                "useFontSpace": "0",
                "useKerning": "0",
                "symMark": "NONE",
                "borderFillIDRef": str(char_shape.border_fill_id),
            },
        )
        ET.SubElement(
            node,
            _q("hh", "fontRef"),
            {
                "hangul": str(char_shape.font_face["hangul"]),
                "latin": str(char_shape.font_face["latin"]),
                "hanja": str(char_shape.font_face["hanja"]),
                "japanese": str(char_shape.font_face["japanese"]),
                "other": str(char_shape.font_face["other"]),
                "symbol": str(char_shape.font_face["symbol"]),
                "user": str(char_shape.font_face["user"]),
            },
        )
        for tag, values in (
            ("ratio", char_shape.letter_width_expansion),
            ("spacing", char_shape.letter_spacing),
            ("relSz", char_shape.relative_size),
            ("offset", char_shape.position),
        ):
            ET.SubElement(
                node,
                _q("hh", tag),
                {
                    "hangul": str(values["hangul"]),
                    "latin": str(values["latin"]),
                    "hanja": str(values["hanja"]),
                    "japanese": str(values["japanese"]),
                    "other": str(values["other"]),
                    "symbol": str(values["symbol"]),
                    "user": str(values["user"]),
                },
            )
        ET.SubElement(node, _q("hh", "underline"), {"type": "NONE", "shape": "SOLID", "color": "#000000"})
        ET.SubElement(node, _q("hh", "strikeout"), {"shape": "NONE", "color": "#000000"})
        ET.SubElement(node, _q("hh", "outline"), {"type": "NONE"})
        ET.SubElement(node, _q("hh", "shadow"), {"type": "NONE", "color": "#C0C0C0", "offsetX": "10", "offsetY": "10"})


def _tab_properties(parent: ET.Element, tab_defs: list[TabDef]) -> None:
    container = ET.SubElement(parent, _q("hh", "tabProperties"), {"itemCnt": str(len(tab_defs))})
    for tab_def in tab_defs:
        ET.SubElement(
            container,
            _q("hh", "tabPr"),
            {
                "id": str(tab_def.id),
                "autoTabLeft": "1" if tab_def.auto_tab_left else "0",
                "autoTabRight": "1" if tab_def.auto_tab_right else "0",
            },
        )


def _para_properties(parent: ET.Element, para_shapes: list[ParaShape]) -> None:
    container = ET.SubElement(parent, _q("hh", "paraProperties"), {"itemCnt": str(len(para_shapes))})
    for para_shape in para_shapes:
        node = ET.SubElement(
            container,
            _q("hh", "paraPr"),
            {
                "id": str(para_shape.id),
                "tabPrIDRef": str(para_shape.tabdef_id),
                "condense": str(para_shape.condense),
                "fontLineHeight": "1" if para_shape.font_line_height else "0",
                "snapToGrid": "1" if para_shape.snap_to_grid else "0",
                "suppressLineNumbers": "1" if para_shape.suppress_line_numbers else "0",
                "checked": "0",
            },
        )
        ET.SubElement(node, _q("hh", "align"), {"horizontal": para_shape.align, "vertical": para_shape.vertical_align})
        ET.SubElement(
            node,
            _q("hh", "heading"),
            {"type": para_shape.heading_type, "idRef": str(para_shape.heading_id_ref), "level": str(para_shape.heading_level)},
        )
        ET.SubElement(
            node,
            _q("hh", "breakSetting"),
            {
                "breakLatinWord": para_shape.break_latin_word,
                "breakNonLatinWord": para_shape.break_non_latin_word,
                "widowOrphan": "1" if para_shape.widow_orphan else "0",
                "keepWithNext": "1" if para_shape.keep_with_next else "0",
                "keepLines": "1" if para_shape.keep_lines else "0",
                "pageBreakBefore": "1" if para_shape.page_break_before else "0",
                "lineWrap": para_shape.line_wrap,
            },
        )
        ET.SubElement(
            node,
            _q("hh", "autoSpacing"),
            {
                "eAsianEng": "1" if para_shape.auto_spacing_easian_eng else "0",
                "eAsianNum": "1" if para_shape.auto_spacing_easian_num else "0",
            },
        )
        switch = ET.SubElement(node, _q("hp", "switch"))
        case = ET.SubElement(switch, _q("hp", "case"), {_q("hp", "required-namespace"): NS["hwpunitchar"]})
        default = ET.SubElement(switch, _q("hp", "default"))
        for parent_node, divisor in ((case, 2), (default, 1)):
            margin = ET.SubElement(parent_node, _q("hh", "margin"))
            ET.SubElement(margin, _q("hc", "intent"), {"value": str(int(para_shape.indent / divisor)), "unit": "HWPUNIT"})
            ET.SubElement(margin, _q("hc", "left"), {"value": str(int(para_shape.margin_left / divisor)), "unit": "HWPUNIT"})
            ET.SubElement(margin, _q("hc", "right"), {"value": str(int(para_shape.margin_right / divisor)), "unit": "HWPUNIT"})
            ET.SubElement(margin, _q("hc", "prev"), {"value": str(int(para_shape.margin_top / divisor)), "unit": "HWPUNIT"})
            ET.SubElement(margin, _q("hc", "next"), {"value": str(int(para_shape.margin_bottom / divisor)), "unit": "HWPUNIT"})
            line_spacing_value = para_shape.line_spacing if divisor == 1 or para_shape.line_spacing_type == "PERCENT" else int(para_shape.line_spacing / divisor)
            ET.SubElement(
                parent_node,
                _q("hh", "lineSpacing"),
                {"type": para_shape.line_spacing_type, "value": str(line_spacing_value), "unit": "HWPUNIT"},
            )
        ET.SubElement(
            node,
            _q("hh", "border"),
            {
                "borderFillIDRef": str(para_shape.borderfill_id),
                "offsetLeft": str(para_shape.border_offset_left),
                "offsetRight": str(para_shape.border_offset_right),
                "offsetTop": str(para_shape.border_offset_top),
                "offsetBottom": str(para_shape.border_offset_bottom),
                "connect": "1" if para_shape.border_connect else "0",
                "ignoreMargin": "1" if para_shape.border_ignore_margin else "0",
            },
        )


def _styles(parent: ET.Element, styles: list[Style]) -> None:
    container = ET.SubElement(parent, _q("hh", "styles"), {"itemCnt": str(len(styles))})
    for style in styles:
        ET.SubElement(
            container,
            _q("hh", "style"),
            {
                "id": str(style.id),
                "type": style.kind,
                "name": style.local_name or style.name or f"Style {style.id}",
                "engName": style.name or style.local_name or f"Style {style.id}",
                "paraPrIDRef": str(style.para_shape_id),
                "charPrIDRef": str(style.char_shape_id),
                "nextStyleIDRef": str(style.next_style_id),
                "langID": str(style.lang_id),
                "lockForm": "0",
            },
        )


def _append_run(parent: ET.Element, run: Run) -> None:
    run_el = ET.SubElement(parent, _q("hp", "run"), {"charPrIDRef": str(run.char_shape_id)})
    text_el: ET.Element | None = None
    last_inline: ET.Element | None = None
    processed_piece = False
    needs_trailing_text = False

    def ensure_text() -> ET.Element:
        nonlocal text_el
        if text_el is None:
            text_el = ET.SubElement(run_el, _q("hp", "t"))
        return text_el

    def add_text(text: str) -> None:
        nonlocal last_inline
        target = ensure_text()
        if last_inline is None:
            target.text = (target.text or "") + text
        else:
            last_inline.tail = (last_inline.tail or "") + text

    def add_inline(tag: str, attributes: dict[str, str] | None = None) -> None:
        nonlocal last_inline
        target = ensure_text()
        last_inline = ET.SubElement(target, _q("hp", tag), attributes or {})

    for piece in run.pieces:
        processed_piece = True
        if piece.kind == "text":
            add_text(str(piece.value or ""))
        elif piece.kind == "tab":
            add_inline("tab", {"width": "4000", "leader": "NONE", "type": "LEFT"})
        elif piece.kind == "linebreak":
            add_inline("lineBreak")
        elif piece.kind == "hyphen":
            add_inline("hyphen")
        elif piece.kind == "nbspace":
            add_inline("nbSpace")
        elif piece.kind == "fwspace":
            add_inline("fwSpace")
        elif piece.kind == "section_def" and isinstance(piece.value, SectionDef):
            run_el.append(_section_def_xml(piece.value))
            text_el = None
            last_inline = None
            needs_trailing_text = False
        elif piece.kind == "columns_def" and isinstance(piece.value, ColumnsDef):
            ctrl = ET.SubElement(run_el, _q("hp", "ctrl"))
            ET.SubElement(ctrl, _q("hp", "colPr"), {"id": "", "type": "NEWSPAPER", "layout": "LEFT", "colCount": str(piece.value.col_count), "sameSz": "1", "sameGap": str(piece.value.spacing)})
            text_el = None
            last_inline = None
            needs_trailing_text = False
        elif piece.kind == "page_num" and isinstance(piece.value, PageNum):
            ctrl = ET.SubElement(run_el, _q("hp", "ctrl"))
            ET.SubElement(
                ctrl,
                _q("hp", "pageNum"),
                {"pos": piece.value.position, "formatType": piece.value.format_type, "sideChar": piece.value.side_char},
            )
            text_el = None
            last_inline = None
            needs_trailing_text = False
        elif piece.kind == "page_hide" and isinstance(piece.value, PageHide):
            ctrl = ET.SubElement(run_el, _q("hp", "ctrl"))
            ET.SubElement(
                ctrl,
                _q("hp", "pageHiding"),
                {
                    "hideHeader": "1" if piece.value.hide_header else "0",
                    "hideFooter": "1" if piece.value.hide_footer else "0",
                    "hideMasterPage": "1" if piece.value.hide_master_page else "0",
                    "hideBorder": "1" if piece.value.hide_border else "0",
                    "hideFill": "1" if piece.value.hide_fill else "0",
                    "hidePageNum": "1" if piece.value.hide_page_num else "0",
                },
            )
            text_el = None
            last_inline = None
            needs_trailing_text = False
        elif piece.kind == "table" and isinstance(piece.value, Table):
            run_el.append(_table_xml(piece.value))
            text_el = None
            last_inline = None
            needs_trailing_text = False
        elif piece.kind == "equation" and isinstance(piece.value, Equation):
            run_el.append(_equation_xml(piece.value))
            text_el = None
            last_inline = None
            needs_trailing_text = False
        elif piece.kind == "field_begin" and isinstance(piece.value, FieldBegin):
            ctrl = ET.SubElement(run_el, _q("hp", "ctrl"))
            ctrl.append(_field_begin_xml(piece.value))
            text_el = None
            last_inline = None
            needs_trailing_text = False
        elif piece.kind == "field_end" and isinstance(piece.value, FieldEnd):
            ctrl = ET.SubElement(run_el, _q("hp", "ctrl"))
            ctrl.append(_field_end_xml(piece.value))
            text_el = None
            last_inline = None
            needs_trailing_text = True

    if processed_piece and text_el is None and needs_trailing_text:
        ET.SubElement(run_el, _q("hp", "t"))


def _paragraph_xml(paragraph) -> ET.Element:
    para_el = ET.Element(
        _q("hp", "p"),
        {
            "id": str(paragraph.instance_id & 0xFFFFFFFF),
            "paraPrIDRef": str(paragraph.para_shape_id),
            "styleIDRef": str(paragraph.style_id),
            "pageBreak": "1" if paragraph.page_break else "0",
            "columnBreak": "1" if paragraph.column_break else "0",
            "merged": "0",
        },
    )
    for run in paragraph.runs:
        _append_run(para_el, run)
    if paragraph.line_segs:
        line_seg_array = ET.SubElement(para_el, _q("hp", "linesegarray"))
        for line_seg in paragraph.line_segs:
            ET.SubElement(
                line_seg_array,
                _q("hp", "lineseg"),
                {
                    "textpos": str(line_seg.textpos),
                    "vertpos": str(line_seg.vertpos),
                    "vertsize": str(line_seg.vertsize),
                    "textheight": str(line_seg.textheight),
                    "baseline": str(line_seg.baseline),
                    "spacing": str(line_seg.spacing),
                    "horzpos": str(line_seg.horzpos),
                    "horzsize": str(line_seg.horzsize),
                    "flags": str(line_seg.flags),
                },
            )
    return para_el


def _table_xml(table: Table) -> ET.Element:
    node = ET.Element(
        _q("hp", "tbl"),
        {
            "id": str(table.instance_id),
            "zOrder": str(table.z_order),
            "numberingType": table.numbering_type,
            "textWrap": table.text_wrap,
            "textFlow": table.text_flow,
            "lock": "0",
            "dropcapstyle": "None",
            "pageBreak": table.page_break,
            "repeatHeader": "1" if table.repeat_header else "0",
            "rowCnt": str(table.rows),
            "colCnt": str(table.cols),
            "cellSpacing": str(table.cell_spacing),
            "borderFillIDRef": str(table.border_fill_id),
            "noAdjust": "0",
        },
    )
    ET.SubElement(
        node,
        _q("hp", "sz"),
        {
            "width": str(table.width),
            "widthRelTo": table.width_rel_to,
            "height": str(table.height),
            "heightRelTo": table.height_rel_to,
            "protect": "0",
        },
    )
    ET.SubElement(
        node,
        _q("hp", "pos"),
        {
            "treatAsChar": "1" if table.treat_as_char else "0",
            "affectLSpacing": "1" if table.affect_line_spacing else "0",
            "flowWithText": "1" if table.flow_with_text else "0",
            "allowOverlap": "1" if table.allow_overlap else "0",
            "holdAnchorAndSO": "1" if table.hold_anchor_and_so else "0",
            "vertRelTo": table.vert_rel_to,
            "horzRelTo": table.horz_rel_to,
            "vertAlign": table.vert_align,
            "horzAlign": table.horz_align,
            "vertOffset": str(table.vert_offset),
            "horzOffset": str(table.horz_offset),
        },
    )
    ET.SubElement(
        node,
        _q("hp", "outMargin"),
        {
            "left": str(table.margin_left),
            "right": str(table.margin_right),
            "top": str(table.margin_top),
            "bottom": str(table.margin_bottom),
        },
    )
    ET.SubElement(
        node,
        _q("hp", "inMargin"),
        {
            "left": str(table.padding_left),
            "right": str(table.padding_right),
            "top": str(table.padding_top),
            "bottom": str(table.padding_bottom),
        },
    )
    if table.description:
        comment = ET.SubElement(node, _q("hp", "shapeComment"))
        comment.text = table.description
    current_row = None
    tr_el: ET.Element | None = None
    for cell in sorted(table.cells, key=lambda item: (item.row_addr, item.col_addr)):
        if current_row != cell.row_addr:
            tr_el = ET.SubElement(node, _q("hp", "tr"))
            current_row = cell.row_addr
        assert tr_el is not None
        tr_el.append(_table_cell_xml(cell))
    return node


def _equation_xml(equation: Equation) -> ET.Element:
    node = ET.Element(
        _q("hp", "equation"),
        {
            "id": str(equation.instance_id),
            "zOrder": str(equation.z_order),
            "numberingType": equation.numbering_type,
            "textWrap": equation.text_wrap,
            "textFlow": equation.text_flow,
            "lock": "0",
            "dropcapstyle": "None",
            "version": equation.version,
            "baseLine": str(equation.base_line),
            "textColor": equation.text_color,
            "baseUnit": str(equation.base_unit),
            "lineMode": equation.line_mode,
            "font": equation.font,
        },
    )
    ET.SubElement(
        node,
        _q("hp", "sz"),
        {
            "width": str(equation.width),
            "widthRelTo": equation.width_rel_to,
            "height": str(equation.height),
            "heightRelTo": equation.height_rel_to,
            "protect": "0",
        },
    )
    ET.SubElement(
        node,
        _q("hp", "pos"),
        {
            "treatAsChar": "1" if equation.treat_as_char else "0",
            "affectLSpacing": "1" if equation.affect_line_spacing else "0",
            "flowWithText": "1" if equation.flow_with_text else "0",
            "allowOverlap": "1" if equation.allow_overlap else "0",
            "holdAnchorAndSO": "1" if equation.hold_anchor_and_so else "0",
            "vertRelTo": equation.vert_rel_to,
            "horzRelTo": equation.horz_rel_to,
            "vertAlign": equation.vert_align,
            "horzAlign": equation.horz_align,
            "vertOffset": str(equation.vert_offset),
            "horzOffset": str(equation.horz_offset),
        },
    )
    ET.SubElement(
        node,
        _q("hp", "outMargin"),
        {
            "left": str(equation.margin_left),
            "right": str(equation.margin_right),
            "top": str(equation.margin_top),
            "bottom": str(equation.margin_bottom),
        },
    )
    if equation.description:
        comment = ET.SubElement(node, _q("hp", "shapeComment"))
        comment.text = equation.description
    script = ET.SubElement(node, _q("hp", "script"))
    script.text = equation.script
    return node


def _field_begin_xml(field_begin: FieldBegin) -> ET.Element:
    node = ET.Element(
        _q("hp", "fieldBegin"),
        {
            "id": str(field_begin.instance_id),
            "type": field_begin.field_type,
            "name": field_begin.name,
            "editable": "1" if field_begin.editable else "0",
            "dirty": "1" if field_begin.dirty else "0",
            "zorder": "-1",
            "fieldid": field_begin.field_id,
        },
    )
    params = ET.SubElement(node, _q("hp", "parameters"), {"cnt": "2", "name": ""})
    integer_param = ET.SubElement(params, _q("hp", "integerParam"), {"name": "Prop"})
    integer_param.text = "9"
    string_param = ET.SubElement(
        params,
        _q("hp", "stringParam"),
        {"name": "Command", f"{{{XML_NS}}}space": "preserve"},
    )
    if field_begin.command:
        string_param.text = field_begin.command
    return node


def _field_end_xml(field_end: FieldEnd) -> ET.Element:
    return ET.Element(
        _q("hp", "fieldEnd"),
        {"beginIDRef": str(field_end.begin_id_ref), "fieldid": field_end.field_id},
    )


def _table_cell_xml(cell: TableCell) -> ET.Element:
    node = ET.Element(
        _q("hp", "tc"),
        {
            "header": "1" if cell.header else "0",
            "hasMargin": "1" if cell.has_margin else "0",
            "protect": "1" if cell.protect else "0",
            "editable": "1" if cell.editable else "0",
            "dirty": "0",
            "borderFillIDRef": str(cell.border_fill_id),
        },
    )
    sub_list = ET.SubElement(
        node,
        _q("hp", "subList"),
        {
            "id": "",
            "textDirection": "HORIZONTAL",
            "lineWrap": "BREAK",
            "vertAlign": cell.vertical_align,
            "linkListIDRef": "0",
            "linkListNextIDRef": "0",
            "textWidth": str(cell.width),
            "textHeight": "0",
            "hasTextRef": "0",
            "hasNumRef": "0",
        },
    )
    for paragraph in cell.paragraphs:
        sub_list.append(_paragraph_xml(paragraph))
    ET.SubElement(node, _q("hp", "cellAddr"), {"colAddr": str(cell.col_addr), "rowAddr": str(cell.row_addr)})
    ET.SubElement(node, _q("hp", "cellSpan"), {"colSpan": str(cell.col_span), "rowSpan": str(cell.row_span)})
    ET.SubElement(node, _q("hp", "cellSz"), {"width": str(cell.width), "height": str(cell.height)})
    ET.SubElement(
        node,
        _q("hp", "cellMargin"),
        {
            "left": str(cell.padding_left),
            "right": str(cell.padding_right),
            "top": str(cell.padding_top),
            "bottom": str(cell.padding_bottom),
        },
    )
    return node


def _section_def_xml(section_def: SectionDef) -> ET.Element:
    page_def = section_def.page_def or PageDef(59528, 84186, 5668, 5668, 5668, 4252, 4252, 4252, 0)
    footnote_shape = section_def.footnote_shape or _default_note_shape()
    endnote_shape = section_def.endnote_shape or _default_note_shape()
    node = ET.Element(
        _q("hp", "secPr"),
        {
            "id": "",
            "textDirection": "HORIZONTAL",
            "spaceColumns": str(section_def.columnspacing),
            "tabStop": str(section_def.default_tab_stops),
            "tabStopVal": str(section_def.default_tab_stops // 2),
            "tabStopUnit": "HWPUNIT",
            "outlineShapeIDRef": str(section_def.numbering_shape_id),
            "memoShapeIDRef": "0",
            "textVerticalWidthHead": "0",
            "masterPageCnt": "0",
        },
    )
    ET.SubElement(node, _q("hp", "grid"), {"lineGrid": str(section_def.grid_vertical), "charGrid": str(section_def.grid_horizontal), "wonggojiFormat": "0"})
    ET.SubElement(node, _q("hp", "startNum"), {"pageStartsOn": "BOTH", "page": str(section_def.starting_pagenum), "pic": str(section_def.starting_picturenum), "tbl": str(section_def.starting_tablenum), "equation": str(section_def.starting_equationnum)})
    ET.SubElement(node, _q("hp", "visibility"), {"hideFirstHeader": "0", "hideFirstFooter": "0", "hideFirstMasterPage": "0", "border": "SHOW_ALL", "fill": "SHOW_ALL", "hideFirstPageNum": "0", "hideFirstEmptyLine": "0", "showLineNumber": "0"})
    ET.SubElement(node, _q("hp", "lineNumberShape"), {"restartType": "0", "countBy": "0", "distance": "0", "startNumber": "0"})
    page_pr = ET.SubElement(node, _q("hp", "pagePr"), {"landscape": "WIDELY" if page_def.height >= page_def.width else "NARROWLY", "width": str(page_def.width), "height": str(page_def.height), "gutterType": "LEFT_ONLY"})
    ET.SubElement(page_pr, _q("hp", "margin"), {"header": str(page_def.header), "footer": str(page_def.footer), "gutter": str(page_def.gutter), "left": str(page_def.left), "right": str(page_def.right), "top": str(page_def.top), "bottom": str(page_def.bottom)})
    node.append(_note_shape_xml("footNotePr", footnote_shape))
    node.append(_note_shape_xml("endNotePr", endnote_shape))
    for page_border_fill in section_def.page_border_fills or _default_page_border_fills():
        node.append(_page_border_fill_xml(page_border_fill))
    return node


def _note_shape_xml(tag: str, note_shape: NoteShape) -> ET.Element:
    node = ET.Element(_q("hp", tag))
    ET.SubElement(node, _q("hp", "autoNumFormat"), {"type": "DIGIT", "userChar": "", "prefixChar": "", "suffixChar": note_shape.suffix, "supscript": "0"})
    ET.SubElement(node, _q("hp", "noteLine"), {"length": str(note_shape.splitter_length), "type": "SOLID", "width": "0.12 mm", "color": note_shape.splitter_color})
    ET.SubElement(node, _q("hp", "noteSpacing"), {"betweenNotes": str(note_shape.splitter_margin_top), "belowLine": str(note_shape.splitter_margin_bottom), "aboveLine": str(note_shape.notes_spacing)})
    ET.SubElement(node, _q("hp", "numbering"), {"type": "CONTINUOUS", "newNum": str(note_shape.starting_number)})
    ET.SubElement(node, _q("hp", "placement"), {"place": "END_OF_DOCUMENT", "beneathText": "0"})
    return node


def _page_border_fill_xml(page_border_fill: PageBorderFill) -> ET.Element:
    node = ET.Element(
        _q("hp", "pageBorderFill"),
        {
            "type": page_border_fill.apply_type,
            "borderFillIDRef": str(page_border_fill.borderfill_id),
            "textBorder": "PAPER",
            "headerInside": "0",
            "footerInside": "0",
            "fillArea": "PAPER",
        },
    )
    ET.SubElement(node, _q("hp", "offset"), {"left": str(page_border_fill.left), "right": str(page_border_fill.right), "top": str(page_border_fill.top), "bottom": str(page_border_fill.bottom)})
    return node


def _default_note_shape() -> NoteShape:
    return NoteShape(")", 1, -1, 567, 567, 850, 1, 0, "#000000")


def _default_page_border_fills() -> list[PageBorderFill]:
    return [
        PageBorderFill("BOTH", 1, 1417, 1417, 1417, 1417),
        PageBorderFill("EVEN", 1, 1417, 1417, 1417, 1417),
        PageBorderFill("ODD", 1, 1417, 1417, 1417, 1417),
    ]


def _border_line(parent: ET.Element, tag: str, line: BorderLine) -> None:
    ET.SubElement(
        parent,
        _q("hh", tag),
        {"type": line.line_type, "width": line.width, "color": line.color},
    )


def _alpha_text(value: float) -> str:
    if float(value).is_integer():
        return str(int(value))
    return str(value)


def _q(prefix: str, local: str) -> str:
    return f"{{{NS[prefix]}}}{local}"


def _xml_bytes(element: ET.Element) -> bytes:
    return ET.tostring(
        element,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=True,
    )
