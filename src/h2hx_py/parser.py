from __future__ import annotations

from contextlib import closing
from pathlib import Path
import struct

from hwp5.xmlmodel import Hwp5File
import olefile

from .model import (
    BorderFill,
    BorderLine,
    CaretPosition,
    CharShape,
    ColumnsDef,
    DocumentMetadata,
    Equation,
    FieldBegin,
    FieldEnd,
    FontFace,
    HwpDocument,
    LineSeg,
    NoteShape,
    PageBorderFill,
    PageDef,
    Paragraph,
    ParaShape,
    PageHide,
    PageNum,
    Table,
    TableCell,
    Piece,
    Run,
    Section,
    SectionDef,
    Style,
    TabDef,
)

LANGUAGE_SPECS = (
    ("ko_fonts", "hangul", "HANGUL"),
    ("en_fonts", "latin", "LATIN"),
    ("cn_fonts", "hanja", "HANJA"),
    ("jp_fonts", "japanese", "JAPANESE"),
    ("other_fonts", "other", "OTHER"),
    ("symbol_fonts", "symbol", "SYMBOL"),
    ("user_fonts", "user", "USER"),
)

CONTROL_CHAR_MAP = {
    9: "tab",
    10: "linebreak",
    13: "paragraph_end",
    24: "hyphen",
    30: "nbspace",
    31: "fwspace",
}


def parse_hwp(source: str | Path) -> HwpDocument:
    path = Path(source)
    warnings: list[str] = []

    with closing(Hwp5File(str(path))) as hwp:
        metadata = _parse_metadata(hwp, path)
        docinfo_models = list(hwp.docinfo.models())
        document_properties = _first_content(docinfo_models, "DocumentProperties")
        id_mappings = _first_content(docinfo_models, "IdMappings")

        font_faces = _parse_font_faces(docinfo_models, id_mappings)
        border_fills = _parse_border_fills(docinfo_models)
        char_shapes = _parse_char_shapes(docinfo_models, border_fills)
        tab_defs = _parse_tab_defs(docinfo_models)
        para_shapes = _parse_para_shapes(docinfo_models)
        styles = _parse_styles(docinfo_models)
        sections = [
            _parse_section(hwp.bodytext.section(index), index, warnings)
            for index in hwp.bodytext.section_indexes()
        ]

        return HwpDocument(
            source=path,
            version=tuple(hwp.fileheader.version),
            metadata=metadata,
            caret=CaretPosition(
                list_id=document_properties.get("list_id", 0),
                paragraph_id=document_properties.get("paragraph_id", 0),
                position_in_paragraph=document_properties.get(
                    "character_unit_loc_in_paragraph", 0
                ),
            ),
            section_count=document_properties.get("section_count", len(sections)),
            font_faces=font_faces,
            border_fills=border_fills or [BorderFill(id=1, has_fill=False)],
            char_shapes=char_shapes or [_default_char_shape()],
            tab_defs=tab_defs or [TabDef(id=0)],
            para_shapes=para_shapes or [_default_para_shape()],
            styles=styles or [_default_style()],
            sections=sections,
            warnings=warnings,
        )


def _parse_metadata(hwp: Hwp5File, source_path: Path) -> DocumentMetadata:
    summary = hwp.summaryinfo
    return DocumentMetadata(
        title=_safe_text(summary.title),
        author=_safe_text(summary.author),
        subject=_safe_text(summary.subject),
        comments=_safe_text(summary.comments),
        last_saved_by=_safe_text(summary.lastSavedBy),
        created_at=str(summary.createdTime or ""),
        modified_at=str(summary.lastSavedTime or ""),
        date_string=_safe_text(summary.dateString),
        keywords=_safe_text(summary.keywords),
        preview_text=_safe_text(hwp.preview_text.get_text()),
        preview_image=_read_preview_image(str(source_path)),
    )


def _parse_font_faces(models: list[dict], id_mappings: dict) -> dict[str, list[FontFace]]:
    faces = [model["content"] for model in models if model["type"].__name__ == "FaceName"]
    grouped: dict[str, list[FontFace]] = {}
    offset = 0
    for count_key, group_key, _ in LANGUAGE_SPECS:
        count = int(id_mappings.get(count_key, 0))
        grouped[group_key] = [
            FontFace(id=index, name=faces[offset + index].get("name", ""))
            for index in range(count)
        ]
        offset += count
    return grouped


def _parse_border_fills(models: list[dict]) -> list[BorderFill]:
    border_models = [m["content"] for m in models if m["type"].__name__ == "BorderFill"]
    result: list[BorderFill] = []
    for index, content in enumerate(border_models):
        fill_pattern = content.get("fill_colorpattern")
        hatch_style = None
        if fill_pattern is not None:
            hatch_style = _hatch_style(fill_pattern["pattern_type_flags"].pattern_type)
        result.append(
            BorderFill(
                id=index + 1,
                has_fill=bool(content.get("fillflags", 0)),
                three_d=bool(content["borderflags"].effect_3d),
                shadow=bool(content["borderflags"].effect_shadow),
                center_line="NONE",
                break_cell_separate_line=False,
                slash_type=_slash_type(content["borderflags"].slash),
                slash_crooked=False,
                slash_counter=False,
                backslash_type=_slash_type(content["borderflags"].backslash),
                backslash_crooked=False,
                backslash_counter=False,
                left_border=_parse_border_line(content["left"]),
                right_border=_parse_border_line(content["right"]),
                top_border=_parse_border_line(content["top"]),
                bottom_border=_parse_border_line(content["bottom"]),
                diagonal_border=_parse_border_line(content["diagonal"]),
                fill_face_color=_hwp_color(fill_pattern.get("background_color", -1), "none")
                if fill_pattern is not None
                else "#FFFFFF",
                fill_hatch_color=_hwp_color(fill_pattern.get("pattern_color", 0), "#000000")
                if fill_pattern is not None
                else "#999999",
                fill_hatch_style=hatch_style,
                fill_alpha=0.0,
            )
        )
    return result


def _parse_char_shapes(models: list[dict], border_fills: list[BorderFill]) -> list[CharShape]:
    default_border_fill_id = border_fills[0].id if border_fills else 1
    result: list[CharShape] = []
    for model in models:
        if model["type"].__name__ != "CharShape":
            continue
        content = model["content"]
        flags = content.get("charshapeflags", 0)
        result.append(
            CharShape(
                id=len(result),
                font_face={
                    "hangul": int(content["font_face"].get("ko", 0)),
                    "latin": int(content["font_face"].get("en", 0)),
                    "hanja": int(content["font_face"].get("cn", 0)),
                    "japanese": int(content["font_face"].get("jp", 0)),
                    "other": int(content["font_face"].get("other", 0)),
                    "symbol": int(content["font_face"].get("symbol", 0)),
                    "user": int(content["font_face"].get("user", 0)),
                },
                basesize=int(content.get("basesize", 1000)),
                text_color=_hwp_color(content.get("text_color", 0), "none"),
                shade_color=_hwp_color(content.get("shade_color", -1), "none"),
                bold=bool(getattr(flags, "bold", 0)),
                italic=bool(getattr(flags, "italic", 0)),
                letter_spacing={
                    "hangul": int(content["letter_spacing"].get("ko", 0)),
                    "latin": int(content["letter_spacing"].get("en", 0)),
                    "hanja": int(content["letter_spacing"].get("cn", 0)),
                    "japanese": int(content["letter_spacing"].get("jp", 0)),
                    "other": int(content["letter_spacing"].get("other", 0)),
                    "symbol": int(content["letter_spacing"].get("symbol", 0)),
                    "user": int(content["letter_spacing"].get("user", 0)),
                },
                letter_width_expansion={
                    "hangul": int(content["letter_width_expansion"].get("ko", 100)),
                    "latin": int(content["letter_width_expansion"].get("en", 100)),
                    "hanja": int(content["letter_width_expansion"].get("cn", 100)),
                    "japanese": int(content["letter_width_expansion"].get("jp", 100)),
                    "other": int(content["letter_width_expansion"].get("other", 100)),
                    "symbol": int(content["letter_width_expansion"].get("symbol", 100)),
                    "user": int(content["letter_width_expansion"].get("user", 100)),
                },
                relative_size={
                    "hangul": int(content["relative_size"].get("ko", 100)),
                    "latin": int(content["relative_size"].get("en", 100)),
                    "hanja": int(content["relative_size"].get("cn", 100)),
                    "japanese": int(content["relative_size"].get("jp", 100)),
                    "other": int(content["relative_size"].get("other", 100)),
                    "symbol": int(content["relative_size"].get("symbol", 100)),
                    "user": int(content["relative_size"].get("user", 100)),
                },
                position={
                    "hangul": int(content["position"].get("ko", 0)),
                    "latin": int(content["position"].get("en", 0)),
                    "hanja": int(content["position"].get("cn", 0)),
                    "japanese": int(content["position"].get("jp", 0)),
                    "other": int(content["position"].get("other", 0)),
                    "symbol": int(content["position"].get("symbol", 0)),
                    "user": int(content["position"].get("user", 0)),
                },
                border_fill_id=default_border_fill_id,
            )
        )
    return result


def _parse_tab_defs(models: list[dict]) -> list[TabDef]:
    result: list[TabDef] = []
    for model in models:
        if model["type"].__name__ != "TabDef":
            continue
        flags = int(model["content"].get("flags", 0))
        result.append(
            TabDef(
                id=len(result),
                auto_tab_left=bool(flags & 0x1),
                auto_tab_right=bool(flags & 0x2),
            )
        )
    return result


def _parse_para_shapes(models: list[dict]) -> list[ParaShape]:
    result: list[ParaShape] = []
    for model in models:
        if model["type"].__name__ != "ParaShape":
            continue
        content = model["content"]
        flags = content["parashapeflags"]
        flags2 = content.get("flags2", 0)
        result.append(
            ParaShape(
                id=len(result),
                tabdef_id=int(content.get("tabdef_id", 0)),
                borderfill_id=max(1, int(content.get("borderfill_id", 1))),
                indent=int(content.get("indent", 0)),
                margin_left=int(content.get("doubled_margin_left", 0)),
                margin_right=int(content.get("doubled_margin_right", 0)),
                margin_top=int(content.get("doubled_margin_top", 0)),
                margin_bottom=int(content.get("doubled_margin_bottom", 0)),
                line_spacing=int(content.get("linespacing", 160)),
                align=_horizontal_align(flags.align),
                vertical_align=_vertical_align(flags.valign),
                heading_type=_heading_type(flags.head_shape),
                heading_id_ref=int(content.get("numbering_bullet_id", 0)),
                heading_level=int(flags.level),
                condense=int(flags.minimum_space),
                font_line_height=bool(flags.lineheight_along_fontsize),
                snap_to_grid=bool(flags.use_paper_grid),
                suppress_line_numbers=bool(getattr(flags2, "reserved", 0) and False),
                break_latin_word=_break_latin_word(flags.linebreak_alphabet),
                break_non_latin_word=_break_non_latin_word(flags.linebreak_hangul),
                widow_orphan=bool(flags.protect_single_line),
                keep_with_next=bool(flags.with_next_paragraph),
                keep_lines=bool(flags.protect),
                page_break_before=bool(flags.start_new_page),
                line_wrap="SQUEEZE" if bool(getattr(flags2, "in_single_line", 0)) else "BREAK",
                line_spacing_type=_line_spacing_type(flags.linespacing_type),
                auto_spacing_easian_eng=bool(getattr(flags2, "autospace_alphabet", 0)),
                auto_spacing_easian_num=bool(getattr(flags2, "autospace_number", 0)),
                border_offset_left=int(content.get("border_left", 0)),
                border_offset_right=int(content.get("border_right", 0)),
                border_offset_top=int(content.get("border_top", 0)),
                border_offset_bottom=int(content.get("border_bottom", 0)),
                border_connect=bool(flags.linked_border),
                border_ignore_margin=bool(flags.ignore_margin),
            )
        )
    return result


def _parse_styles(models: list[dict]) -> list[Style]:
    result: list[Style] = []
    for model in models:
        if model["type"].__name__ != "Style":
            continue
        content = model["content"]
        result.append(
            Style(
                id=len(result),
                local_name=_safe_text(content.get("local_name")),
                name=_safe_text(content.get("name")),
                kind="PARA",
                para_shape_id=int(content.get("parashape_id", 0)),
                char_shape_id=int(content.get("charshape_id", 0)),
                next_style_id=int(content.get("next_style_id", 0)),
                lang_id=int(content.get("lang_id", 1042)),
            )
        )
    return result


def _parse_section(section_stream, section_index: int, warnings: list[str]) -> Section:
    models = list(section_stream.models())
    paragraphs: list[Paragraph] = []
    index = 0

    while index < len(models):
        model = models[index]
        if model["type"].__name__ != "Paragraph" or int(model.get("level", 0)) != 0:
            index += 1
            continue

        header = model["content"]
        index += 1
        related: list[dict] = []
        while index < len(models):
            next_model = models[index]
            if next_model["type"].__name__ == "Paragraph" and int(next_model.get("level", 0)) == 0:
                break
            related.append(next_model)
            index += 1

        paragraphs.append(_parse_paragraph(header, related, warnings))

    return Section(index=section_index, paragraphs=paragraphs)


def _parse_paragraph(header: dict, related: list[dict], warnings: list[str]) -> Paragraph:
    chunks = []
    charshape_pairs: list[tuple[int, int]] = []
    line_segs: list[LineSeg] = []
    section_def: SectionDef | None = None
    columns_def: ColumnsDef | None = None
    page_num: PageNum | None = None
    prefix_pieces: list[Piece] = []
    pending_page_hides: list[PageHide] = []
    pending_tables: list[Table] = []
    pending_equation_controls: list[dict[str, object]] = []
    pending_equations: list[Equation] = []
    pending_field_begins: list[FieldBegin] = []
    paragraph_level = min((int(record.get("level", 1)) for record in related), default=1)

    index = 0
    while index < len(related):
        record = related[index]
        name = record["type"].__name__
        content = record["content"]
        if name == "ParaText":
            chunks = content.get("chunks", [])
        elif name == "ParaCharShape":
            charshape_pairs = list(content.get("charshapes", []))
        elif name == "ParaLineSeg":
            line_segs = [
                LineSeg(
                    textpos=int(item.get("chpos", 0)),
                    vertpos=int(item.get("y", 0)),
                    vertsize=int(item.get("height", 0)),
                    textheight=int(item.get("height_text", 0)),
                    baseline=int(item.get("height_baseline", 0)),
                    spacing=int(item.get("space_below", 0)),
                    horzpos=int(item.get("x", 0)),
                    horzsize=int(item.get("width", 0)),
                    flags=int(item.get("lineseg_flags", 0)),
                )
                for item in content.get("linesegs", [])
            ]
        elif name == "SectionDef":
            section_def = SectionDef(
                columnspacing=int(content.get("columnspacing", 0)),
                grid_vertical=int(content.get("grid_vertical", 0)),
                grid_horizontal=int(content.get("grid_horizontal", 0)),
                default_tab_stops=int(content.get("defaultTabStops", 8000)),
                numbering_shape_id=int(content.get("numbering_shape_id", 0)),
                starting_pagenum=int(content.get("starting_pagenum", 0)),
                starting_picturenum=int(content.get("starting_picturenum", 0)),
                starting_tablenum=int(content.get("starting_tablenum", 0)),
                starting_equationnum=int(content.get("starting_equationnum", 0)),
            )
            prefix_pieces.append(Piece("section_def", section_def))
        elif name == "PageDef" and section_def is not None:
            section_def.page_def = PageDef(
                width=int(content.get("width", 0)),
                height=int(content.get("height", 0)),
                left=int(content.get("left_offset", 0)),
                right=int(content.get("right_offset", 0)),
                top=int(content.get("top_offset", 0)),
                bottom=int(content.get("bottom_offset", 0)),
                header=int(content.get("header_offset", 0)),
                footer=int(content.get("footer_offset", 0)),
                gutter=int(content.get("bookbinding_offset", 0)),
            )
        elif name == "FootnoteShape" and section_def is not None:
            target = "footnote_shape" if section_def.footnote_shape is None else "endnote_shape"
            setattr(section_def, target, _make_note_shape(content))
        elif name == "PageBorderFill" and section_def is not None:
            apply_type = ("BOTH", "EVEN", "ODD")[len(section_def.page_border_fills) % 3]
            margin = content.get("margin", {})
            section_def.page_border_fills.append(
                PageBorderFill(
                    apply_type=apply_type,
                    borderfill_id=max(1, int(content.get("borderfill_id", 1))),
                    left=int(margin.get("left", 0)),
                    right=int(margin.get("right", 0)),
                    top=int(margin.get("top", 0)),
                    bottom=int(margin.get("bottom", 0)),
                )
            )
        elif name == "ColumnsDef":
            columns_def = ColumnsDef(
                spacing=int(content.get("spacing", 0)),
                col_count=1,
            )
            prefix_pieces.append(Piece("columns_def", columns_def))
        elif name == "PageNumberPosition":
            page_num = PageNum(
                position="BOTTOM_CENTER",
                format_type="DIGIT",
                side_char=chr(content.get("dash", 45)) if content.get("dash") else "-",
            )
            prefix_pieces.append(Piece("page_num", page_num))
        elif name == "PageHide":
            flags = int(content.get("flags", 0))
            pending_page_hides.append(
                PageHide(
                    hide_header=bool(flags & 0x1),
                    hide_footer=bool(flags & 0x2),
                    hide_master_page=bool(flags & 0x4),
                    hide_border=bool(flags & 0x8),
                    hide_fill=bool(flags & 0x10),
                    hide_page_num=bool(flags & 0x20),
                )
            )
        elif name == "Control":
            chid = content.get("chid")
            if chid == "eqed":
                pending_equation_controls.append(_parse_common_control(record))
            elif chid not in {"secd", "cold", "pgnp", "pghd"}:
                warnings.append(f"Unsupported control header `{chid}` was skipped.")
        elif name == "TableControl":
            table, index = _parse_table(related, index, warnings)
            pending_tables.append(table)
            continue
        elif name == "EqEdit":
            if pending_equation_controls:
                pending_equations.append(_parse_equation(record, pending_equation_controls.pop(0)))
            else:
                warnings.append("EqEdit was skipped because its control header was missing.")
        elif name == "FieldBookmark":
            pending_field_begins.append(_parse_field_begin(record))
        elif name == "ControlData":
            pass
        elif name == "Paragraph" and int(record.get("level", 0)) > paragraph_level:
            warnings.append("Nested paragraph was skipped outside of a table context.")
            nested_level = int(record.get("level", 0))
            index += 1
            while index < len(related) and int(related[index].get("level", 0)) > nested_level:
                index += 1
            continue
        elif name not in {"PageNumberPosition", "ParaRangeTag", "TableBody", "TableCell"}:
            warnings.append(f"Unsupported control `{name}` was skipped.")
        index += 1

    runs = _build_runs(
        chunks,
        charshape_pairs,
        prefix_pieces,
        pending_page_hides,
        pending_tables,
        pending_equations,
        pending_field_begins,
        warnings,
    )
    if not runs:
        default_charshape = charshape_pairs[0][1] if charshape_pairs else 0
        runs = [Run(char_shape_id=default_charshape)]

    return Paragraph(
        instance_id=int(header.get("instance_id", 0)),
        para_shape_id=int(header.get("parashape_id", 0)),
        style_id=int(header.get("style_id", 0)),
        split=int(header.get("split", 0)),
        page_break=_paragraph_page_break(header.get("split", 0)),
        column_break=False,
        runs=runs,
        line_segs=line_segs,
    )


def _build_runs(
    chunks: list,
    charshape_pairs: list[tuple[int, int]],
    prefix_pieces: list[Piece],
    page_hides: list[PageHide],
    tables: list[Table],
    equations: list[Equation],
    field_begins: list[FieldBegin],
    warnings: list[str],
) -> list[Run]:
    pieces_by_shape: list[tuple[int, Piece]] = []
    ranges = _shape_ranges(charshape_pairs)
    default_shape_id = charshape_pairs[0][1] if charshape_pairs else 0
    active_fields: list[FieldBegin] = []
    for piece in prefix_pieces:
        pieces_by_shape.append((default_shape_id, piece))

    for chunk_range, value in chunks:
        start, end = chunk_range
        if isinstance(value, str):
            for piece_start, piece_text in _split_text_piece(start, end, value, ranges):
                pieces_by_shape.append((_shape_id_at(piece_start, ranges), Piece("text", piece_text)))
            continue

        if isinstance(value, dict):
            code = int(value.get("code", -1))
            mapped = CONTROL_CHAR_MAP.get(code)
            if mapped == "paragraph_end":
                continue
            if mapped:
                pieces_by_shape.append((_shape_id_at(start, ranges), Piece(mapped)))
                continue

            chid = value.get("chid")
            if chid == "pghd" and page_hides:
                pieces_by_shape.append((_shape_id_at(start, ranges), Piece("page_hide", page_hides.pop(0))))
            elif chid == "tbl " and tables:
                pieces_by_shape.append((_shape_id_at(start, ranges), Piece("table", tables.pop(0))))
            elif chid == "eqed" and equations:
                pieces_by_shape.append((_shape_id_at(start, ranges), Piece("equation", equations.pop(0))))
            elif chid == "%bmk" and field_begins:
                field_begin = field_begins.pop(0)
                active_fields.append(field_begin)
                pieces_by_shape.append((_shape_id_at(start, ranges), Piece("field_begin", field_begin)))
            elif chid == "\x02bmk":
                if active_fields:
                    field_begin = active_fields.pop()
                    pieces_by_shape.append(
                        (
                            _shape_id_at(start, ranges),
                            Piece(
                                "field_end",
                                FieldEnd(begin_id_ref=field_begin.instance_id, field_id=field_begin.field_id),
                            ),
                        )
                    )
                else:
                    warnings.append("Field end was skipped because no matching field begin was active.")
            elif chid in {"secd", "cold", "pgnp"}:
                continue
            else:
                warnings.append(f"Unsupported inline control `{chid or code}` was skipped.")

    runs: list[Run] = []
    current: Run | None = None
    for char_shape_id, piece in pieces_by_shape:
        if current is None or current.char_shape_id != char_shape_id:
            current = Run(char_shape_id=char_shape_id)
            runs.append(current)
        current.pieces.append(piece)
    return runs


def _parse_table(related: list[dict], start_index: int, warnings: list[str]) -> tuple[Table, int]:
    table_record = related[start_index]
    table_content = table_record["content"]
    common = _common_control_from_content(table_content)
    table_level = int(table_record.get("level", 0))
    direct_child_level = table_level + 1

    body_content: dict = {}
    cells: list[TableCell] = []
    index = start_index + 1
    if (
        index < len(related)
        and related[index]["type"].__name__ == "TableBody"
        and int(related[index].get("level", 0)) == direct_child_level
    ):
        body_content = related[index]["content"]
        index += 1

    while index < len(related):
        record = related[index]
        level = int(record.get("level", 0))
        if level <= table_level:
            break
        if record["type"].__name__ != "TableCell" or level != direct_child_level:
            index += 1
            continue
        cell, index = _parse_table_cell(related, index, warnings)
        cells.append(cell)

    table = Table(
        instance_id=common["instance_id"],
        z_order=common["z_order"],
        width=common["width"],
        height=common["height"],
        margin_left=common["margin_left"],
        margin_right=common["margin_right"],
        margin_top=common["margin_top"],
        margin_bottom=common["margin_bottom"],
        rows=int(body_content.get("rows", 0)),
        cols=int(body_content.get("cols", 0)),
        cell_spacing=int(body_content.get("cellspacing", 0)),
        border_fill_id=max(1, int(body_content.get("borderfill_id", 1))),
        numbering_type=common["numbering_type"],
        text_wrap=common["text_wrap"],
        text_flow=common["text_flow"],
        repeat_header=bool(getattr(body_content.get("flags"), "repeat_header", True)) if body_content else True,
        page_break=_table_page_break(getattr(body_content.get("flags"), "split_page", None)),
        treat_as_char=common["treat_as_char"],
        affect_line_spacing=common["affect_line_spacing"],
        flow_with_text=common["flow_with_text"],
        allow_overlap=common["allow_overlap"],
        hold_anchor_and_so=False,
        vert_rel_to=common["vert_rel_to"],
        horz_rel_to=common["horz_rel_to"],
        vert_align=common["vert_align"],
        horz_align=common["horz_align"],
        vert_offset=common["vert_offset"],
        horz_offset=common["horz_offset"],
        width_rel_to=common["width_rel_to"],
        height_rel_to=common["height_rel_to"],
        padding_left=int(body_content.get("padding", {}).get("left", 0)),
        padding_right=int(body_content.get("padding", {}).get("right", 0)),
        padding_top=int(body_content.get("padding", {}).get("top", 0)),
        padding_bottom=int(body_content.get("padding", {}).get("bottom", 0)),
        description=common["description"],
        cells=cells,
    )
    return table, index


def _parse_table_cell(related: list[dict], start_index: int, warnings: list[str]) -> tuple[TableCell, int]:
    cell_record = related[start_index]
    cell_content = cell_record["content"]
    cell_level = int(cell_record.get("level", 0))
    index = start_index + 1
    paragraphs: list[Paragraph] = []
    paragraph_count = int(cell_content.get("paragraphs", 0))

    while index < len(related) and len(paragraphs) < paragraph_count:
        record = related[index]
        if record["type"].__name__ != "Paragraph" or int(record.get("level", 0)) != cell_level:
            if int(record.get("level", 0)) <= cell_level:
                break
            index += 1
            continue
        paragraph, index = _parse_nested_paragraph(related, index, warnings)
        paragraphs.append(paragraph)

    cell = TableCell(
        header=False,
        has_margin=False,
        protect=False,
        editable=False,
        border_fill_id=max(1, int(cell_content.get("borderfill_id", 1))),
        col_addr=int(cell_content.get("col", 0)),
        row_addr=int(cell_content.get("row", 0)),
        col_span=int(cell_content.get("colspan", 1)),
        row_span=int(cell_content.get("rowspan", 1)),
        width=int(cell_content.get("width", 0)),
        height=int(cell_content.get("height", 0)),
        padding_left=int(cell_content.get("padding", {}).get("left", 0)),
        padding_right=int(cell_content.get("padding", {}).get("right", 0)),
        padding_top=int(cell_content.get("padding", {}).get("top", 0)),
        padding_bottom=int(cell_content.get("padding", {}).get("bottom", 0)),
        vertical_align=_cell_vertical_align(cell_content.get("listflags")),
        paragraphs=paragraphs,
    )
    return cell, index


def _parse_nested_paragraph(related: list[dict], start_index: int, warnings: list[str]) -> tuple[Paragraph, int]:
    header_record = related[start_index]
    header = header_record["content"]
    paragraph_level = int(header_record.get("level", 0))
    index = start_index + 1
    nested_related: list[dict] = []
    while index < len(related):
        record = related[index]
        level = int(record.get("level", 0))
        if level <= paragraph_level:
            break
        nested_related.append(record)
        index += 1
    return _parse_paragraph(header, nested_related, warnings), index


def _make_note_shape(content: dict) -> NoteShape:
    suffix = chr(content.get("suffix", 0)) if content.get("suffix") else ""
    return NoteShape(
        suffix=suffix,
        starting_number=int(content.get("starting_number", 1)),
        splitter_length=int(content.get("splitter_length", -1)),
        splitter_margin_top=int(content.get("splitter_margin_top", 0)),
        splitter_margin_bottom=int(content.get("splitter_margin_bottom", 0)),
        notes_spacing=int(content.get("notes_spacing", 0)),
        splitter_stroke_type=int(content.get("splitter_stroke_type", 1)),
        splitter_width=int(content.get("splitter_width", 0)),
        splitter_color=_hwp_color(content.get("splitter_color", 0), "#000000"),
    )


def _parse_common_control(record: dict) -> dict[str, object]:
    payload = record.get("unparsed") or b""
    if not payload:
        full_payload = record.get("payload") or b""
        payload = full_payload[4:] if len(full_payload) >= 4 else full_payload

    offset = 0
    flags, offset = _read_u32(payload, offset)
    y, offset = _read_i32(payload, offset)
    x, offset = _read_i32(payload, offset)
    width, offset = _read_u32(payload, offset)
    height, offset = _read_u32(payload, offset)
    z_order, offset = _read_i16(payload, offset)
    _, offset = _read_i16(payload, offset)
    margin_left, offset = _read_u16(payload, offset)
    margin_right, offset = _read_u16(payload, offset)
    margin_top, offset = _read_u16(payload, offset)
    margin_bottom, offset = _read_u16(payload, offset)
    instance_id, offset = _read_u32(payload, offset)
    description = ""
    if offset + 2 <= len(payload):
        _, offset = _read_i16(payload, offset)
    if offset + 2 <= len(payload):
        description, offset = _read_bstr(payload, offset)
    if not description and offset + 2 <= len(payload):
        description, offset = _read_bstr(payload, offset)

    return _common_control_values(
        flags=flags,
        x=x,
        y=y,
        width=width,
        height=height,
        z_order=z_order,
        margin_left=margin_left,
        margin_right=margin_right,
        margin_top=margin_top,
        margin_bottom=margin_bottom,
        instance_id=instance_id,
        description=description,
    )


def _common_control_from_content(content: dict) -> dict[str, object]:
    flags = int(content.get("flags", 0))
    margin = content.get("margin", {})
    return _common_control_values(
        flags=flags,
        x=int(content.get("x", 0)),
        y=int(content.get("y", 0)),
        width=int(content.get("width", 0)),
        height=int(content.get("height", 0)),
        z_order=int(content.get("z_order", 0)),
        margin_left=int(margin.get("left", 0)),
        margin_right=int(margin.get("right", 0)),
        margin_top=int(margin.get("top", 0)),
        margin_bottom=int(margin.get("bottom", 0)),
        instance_id=int(content.get("instance_id", 0)),
        description=_safe_text(content.get("description")),
    )


def _common_control_values(
    *,
    flags: int,
    x: int,
    y: int,
    width: int,
    height: int,
    z_order: int,
    margin_left: int,
    margin_right: int,
    margin_top: int,
    margin_bottom: int,
    instance_id: int,
    description: str,
) -> dict[str, object]:
    return {
        "flags": flags,
        "x": x,
        "y": y,
        "width": width,
        "height": height,
        "z_order": z_order,
        "margin_left": margin_left,
        "margin_right": margin_right,
        "margin_top": margin_top,
        "margin_bottom": margin_bottom,
        "instance_id": instance_id,
        "description": description,
        "treat_as_char": _flag_bits(flags, 0, 0) == 1,
        "affect_line_spacing": _flag_bits(flags, 2, 2) == 1,
        "vert_rel_to": _vert_rel_to(_flag_bits(flags, 3, 4)),
        "vert_align": _vert_align(_flag_bits(flags, 5, 7)),
        "horz_rel_to": _horz_rel_to(_flag_bits(flags, 8, 9)),
        "horz_align": _horz_align(_flag_bits(flags, 10, 12)),
        "flow_with_text": _flag_bits(flags, 13, 13) == 1,
        "allow_overlap": _flag_bits(flags, 14, 14) == 1,
        "width_rel_to": _width_rel_to(_flag_bits(flags, 15, 17)),
        "height_rel_to": _height_rel_to(_flag_bits(flags, 18, 19)),
        "text_wrap": _text_wrap(_flag_bits(flags, 21, 23)),
        "text_flow": _text_flow(_flag_bits(flags, 24, 25)),
        "numbering_type": _numbering_type(_flag_bits(flags, 26, 27)),
        "vert_offset": y,
        "horz_offset": x,
    }


def _parse_equation(record: dict, control: dict[str, object]) -> Equation:
    payload = record.get("payload") or record.get("unparsed") or b""
    offset = 0
    property_flags, offset = _read_u32(payload, offset)
    script, offset = _read_bstr(payload, offset)
    base_unit, offset = _read_u32(payload, offset)
    text_color, offset = _read_u32(payload, offset)
    base_line, offset = _read_u32(payload, offset)
    version, offset = _read_bstr(payload, offset)
    font, offset = _read_bstr(payload, offset)
    return Equation(
        instance_id=int(control["instance_id"]),
        z_order=int(control["z_order"]),
        width=int(control["width"]),
        height=int(control["height"]),
        margin_left=int(control["margin_left"]),
        margin_right=int(control["margin_right"]),
        margin_top=int(control["margin_top"]),
        margin_bottom=int(control["margin_bottom"]),
        numbering_type=str(control["numbering_type"]),
        text_wrap=str(control["text_wrap"]),
        text_flow=str(control["text_flow"]),
        treat_as_char=bool(control["treat_as_char"]),
        affect_line_spacing=bool(control["affect_line_spacing"]),
        flow_with_text=bool(control["flow_with_text"]),
        allow_overlap=bool(control["allow_overlap"]),
        hold_anchor_and_so=False,
        vert_rel_to=str(control["vert_rel_to"]),
        horz_rel_to=str(control["horz_rel_to"]),
        vert_align=str(control["vert_align"]),
        horz_align=str(control["horz_align"]),
        vert_offset=int(control["vert_offset"]),
        horz_offset=int(control["horz_offset"]),
        width_rel_to=str(control["width_rel_to"]),
        height_rel_to=str(control["height_rel_to"]),
        version=version,
        base_line=int(base_line),
        text_color=_hwp_color(text_color, "#000000"),
        base_unit=int(base_unit),
        line_mode="LINE" if property_flags & 0x1 else "CHAR",
        font=font,
        script=script,
        description=str(control["description"]),
    )


def _parse_field_begin(record: dict) -> FieldBegin:
    content = record["content"]
    payload = record.get("payload") or b""
    field_id = 0
    if len(payload) >= 4:
        field_id = struct.unpack_from("<I", payload, 0)[0]
    flags = int(content.get("flags", 0))
    chid = content.get("chid")
    return FieldBegin(
        instance_id=int(content.get("id", 0)),
        field_type="BOOKMARK" if chid == "%bmk" else "UNKNOWN",
        name="",
        editable=bool(flags & 0x1),
        dirty=bool(flags & 0x8000),
        field_id=str(field_id),
        command=_safe_text(content.get("command")),
    )


def _parse_border_line(content: dict) -> BorderLine:
    return BorderLine(
        line_type=_line_type(content["stroke_flags"].stroke_type),
        width=_line_width(content["width_flags"].width),
        color=_hwp_color(content.get("color", 0), "#000000"),
    )


def _shape_ranges(charshape_pairs: list[tuple[int, int]]) -> list[tuple[int, int, int]]:
    if not charshape_pairs:
        return [(0, 2**31 - 1, 0)]
    result: list[tuple[int, int, int]] = []
    for index, (start, shape_id) in enumerate(charshape_pairs):
        end = charshape_pairs[index + 1][0] if index + 1 < len(charshape_pairs) else 2**31 - 1
        result.append((int(start), int(end), int(shape_id)))
    return result


def _shape_id_at(position: int, ranges: list[tuple[int, int, int]]) -> int:
    for start, end, shape_id in ranges:
        if start <= position < end:
            return shape_id
    return ranges[-1][2]


def _split_text_piece(
    start: int,
    end: int,
    text: str,
    ranges: list[tuple[int, int, int]],
) -> list[tuple[int, str]]:
    if not text:
        return []
    result: list[tuple[int, str]] = []
    current = start
    offset = 0
    while current < end and offset < len(text):
        boundary = min(
            [range_end for range_start, range_end, _ in ranges if range_start <= current < range_end]
            or [end]
        )
        piece_len = min(boundary - current, len(text) - offset)
        if piece_len <= 0:
            break
        result.append((current, text[offset : offset + piece_len]))
        current += piece_len
        offset += piece_len
    if offset < len(text):
        result.append((current, text[offset:]))
    return result


def _first_content(models: list[dict], type_name: str) -> dict:
    for model in models:
        if model["type"].__name__ == type_name:
            return model["content"]
    return {}


def _safe_text(value) -> str:
    return "" if value is None else str(value)


def _read_u16(payload: bytes, offset: int) -> tuple[int, int]:
    return struct.unpack_from("<H", payload, offset)[0], offset + 2


def _read_i16(payload: bytes, offset: int) -> tuple[int, int]:
    return struct.unpack_from("<h", payload, offset)[0], offset + 2


def _read_u32(payload: bytes, offset: int) -> tuple[int, int]:
    return struct.unpack_from("<I", payload, offset)[0], offset + 4


def _read_i32(payload: bytes, offset: int) -> tuple[int, int]:
    return struct.unpack_from("<i", payload, offset)[0], offset + 4


def _read_bstr(payload: bytes, offset: int) -> tuple[str, int]:
    length, offset = _read_u16(payload, offset)
    byte_length = length * 2
    text = payload[offset : offset + byte_length].decode("utf-16le", errors="ignore")
    return text, offset + byte_length


def _flag_bits(value: int, start: int, end: int) -> int:
    width = end - start + 1
    return (int(value) >> start) & ((1 << width) - 1)


def _text_wrap(value: int) -> str:
    return {
        0: "SQUARE",
        1: "TOP_AND_BOTTOM",
        2: "BEHIND_TEXT",
        3: "IN_FRONT_OF_TEXT",
    }.get(value, "TOP_AND_BOTTOM")


def _text_flow(value: int) -> str:
    return {
        0: "BOTH_SIDES",
        1: "LEFT_ONLY",
        2: "RIGHT_ONLY",
        3: "LARGEST_ONLY",
    }.get(value, "BOTH_SIDES")


def _numbering_type(value: int) -> str:
    return {
        0: "NONE",
        1: "PICTURE",
        2: "TABLE",
        3: "EQUATION",
    }.get(value, "NONE")


def _vert_rel_to(value: int) -> str:
    return {0: "PAPER", 1: "PAGE", 2: "PARA"}.get(value, "PAPER")


def _horz_rel_to(value: int) -> str:
    return {0: "PAPER", 1: "PAGE", 2: "COLUMN", 3: "PARA"}.get(value, "PAPER")


def _vert_align(value: int) -> str:
    return {0: "TOP", 1: "CENTER", 2: "BOTTOM", 3: "INSIDE", 4: "OUTSIDE"}.get(value, "TOP")


def _horz_align(value: int) -> str:
    return {0: "LEFT", 1: "CENTER", 2: "RIGHT", 3: "INSIDE", 4: "OUTSIDE"}.get(value, "LEFT")


def _width_rel_to(value: int) -> str:
    return {0: "PAPER", 1: "PAGE", 2: "COLUMN", 3: "PARA", 4: "ABSOLUTE"}.get(value, "PAPER")


def _height_rel_to(value: int) -> str:
    return {0: "PAPER", 1: "PAGE", 2: "ABSOLUTE"}.get(value, "PAPER")


def _line_type(value) -> str:
    key = _enum_name(value)
    return {
        "none": "NONE",
        "solid": "SOLID",
        "dashed": "DASH",
        "dotted": "DOT",
        "dash-dot": "DASH_DOT",
        "dash-dot-dot": "DASH_DOT_DOT",
        "long-dash": "LONG_DASH",
        "large-dot": "CIRCLE",
        "double": "DOUBLE_SLIM",
        "double-2": "SLIM_THICK",
        "double-3": "THICK_SLIM",
        "triple": "SLIM_THICK_SLIM",
        "wave": "WAVE",
        "double-wave": "DOUBLEWAVE",
    }.get(key, "NONE")


def _line_width(value) -> str:
    key = _enum_name(value)
    return f"{key.replace('mm', ' mm')}" if key.endswith("mm") else "0.1 mm"


def _slash_type(value) -> str:
    return {
        0: "NONE",
        1: "CENTER",
        2: "CENTER_BELOW",
        3: "CENTER_ABOVE",
        4: "ALL",
    }.get(int(value), "NONE")


def _hatch_style(value) -> str | None:
    key = _enum_name(value)
    return {
        "NONE": None,
        "HORIZONTAL": "HORIZONTAL",
        "VERTICAL": "VERTICAL",
        "BACKSLASH": "BACK_SLASH",
        "SLASH": "SLASH",
        "GRID": "CROSS_DIAGONAL",
        "CROSS": "CROSS",
    }.get(key, None)


def _horizontal_align(value) -> str:
    return {
        "BOTH": "JUSTIFY",
        "LEFT": "LEFT",
        "RIGHT": "RIGHT",
        "CENTER": "CENTER",
        "DISTRIBUTE": "DISTRIBUTE",
        "DISTRIBUTE_SPACE": "DISTRIBUTE_SPACE",
    }.get(_enum_name(value), "JUSTIFY")


def _vertical_align(value) -> str:
    return {
        "FONT": "BASELINE",
        "TOP": "TOP",
        "CENTER": "CENTER",
        "BOTTOM": "BOTTOM",
    }.get(_enum_name(value), "BASELINE")


def _heading_type(value) -> str:
    return {
        "NONE": "NONE",
        "OUTLINE": "OUTLINE",
        "NUMBER": "NUMBER",
        "BULLET": "BULLET",
    }.get(_enum_name(value), "NONE")


def _break_latin_word(value) -> str:
    return {
        "WORD": "KEEP_WORD",
        "HYPHEN": "HYPHENATION",
        "CHAR": "BREAK_WORD",
    }.get(_enum_name(value), "BREAK_WORD")


def _break_non_latin_word(value) -> str:
    return {
        "CHAR": "KEEP_WORD",
        "WORD": "BREAK_WORD",
    }.get(_enum_name(value), "BREAK_WORD")


def _line_spacing_type(value) -> str:
    return {
        "RATIO": "PERCENT",
        "FIXED": "PERCENT",
        "SPACEONLY": "BETWEEN_LINES",
        "MINIMUM": "AT_LEAST",
    }.get(_enum_name(value), "PERCENT")


def _cell_vertical_align(value) -> str:
    if value is None:
        return "CENTER"
    name = _enum_name(value.valign)
    return {"TOP": "TOP", "MIDDLE": "CENTER", "BOTTOM": "BOTTOM"}.get(name, "CENTER")


def _table_page_break(value) -> str:
    name = _enum_name(value) if value is not None else "SPLIT"
    return {
        "NONE": "NONE",
        "BY_CELL": "CELL",
        "SPLIT": "TABLE",
    }.get(name, "TABLE")


def _paragraph_page_break(value) -> bool:
    return bool(int(value or 0) & 0x4)


def _enum_name(value) -> str:
    name = getattr(value, "name", None)
    if name is not None:
        return str(name)
    text = str(value)
    return text.split(".")[-1]


def _hwp_color(value: int, none_value: str) -> str:
    if value is None or int(value) < 0:
        return none_value
    color = int(value) & 0xFFFFFF
    red = (color >> 16) & 0xFF
    green = (color >> 8) & 0xFF
    blue = color & 0xFF
    return f"#{blue:02X}{green:02X}{red:02X}"


def _default_char_shape() -> CharShape:
    keys = [group_key for _, group_key, _ in LANGUAGE_SPECS]
    return CharShape(
        id=0,
        font_face={key: 0 for key in keys},
        basesize=1000,
        text_color="#000000",
        shade_color="none",
        bold=False,
        italic=False,
        letter_spacing={key: 0 for key in keys},
        letter_width_expansion={key: 100 for key in keys},
        relative_size={key: 100 for key in keys},
        position={key: 0 for key in keys},
        border_fill_id=1,
    )


def _default_para_shape() -> ParaShape:
    return ParaShape(
        id=0,
        tabdef_id=0,
        borderfill_id=1,
        indent=0,
        margin_left=0,
        margin_right=0,
        margin_top=0,
        margin_bottom=0,
        line_spacing=160,
    )


def _default_style() -> Style:
    return Style(
        id=0,
        local_name="Default",
        name="Default",
        kind="PARA",
        para_shape_id=0,
        char_shape_id=0,
        next_style_id=0,
        lang_id=1042,
    )


def _read_preview_image(path: str) -> bytes | None:
    try:
        ole = olefile.OleFileIO(path)
        try:
            if ole.exists("PrvImage"):
                return ole.openstream("PrvImage").read()
        finally:
            ole.close()
    except OSError:
        return None
    return None
