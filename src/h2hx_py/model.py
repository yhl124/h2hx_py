from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(slots=True)
class DocumentMetadata:
    title: str = ""
    author: str = ""
    subject: str = ""
    comments: str = ""
    last_saved_by: str = ""
    created_at: str = ""
    modified_at: str = ""
    date_string: str = ""
    keywords: str = ""
    preview_text: str = ""
    preview_image: bytes | None = None


@dataclass(slots=True)
class FontFace:
    id: int
    name: str


@dataclass(slots=True)
class BorderLine:
    line_type: str = "NONE"
    width: str = "0.1 mm"
    color: str = "#000000"


@dataclass(slots=True)
class BorderFill:
    id: int
    has_fill: bool = False
    three_d: bool = False
    shadow: bool = False
    center_line: str = "NONE"
    break_cell_separate_line: bool = False
    slash_type: str = "NONE"
    slash_crooked: bool = False
    slash_counter: bool = False
    backslash_type: str = "NONE"
    backslash_crooked: bool = False
    backslash_counter: bool = False
    left_border: BorderLine = field(default_factory=BorderLine)
    right_border: BorderLine = field(default_factory=BorderLine)
    top_border: BorderLine = field(default_factory=BorderLine)
    bottom_border: BorderLine = field(default_factory=BorderLine)
    diagonal_border: BorderLine = field(default_factory=lambda: BorderLine(line_type="SOLID"))
    fill_face_color: str = "#FFFFFF"
    fill_hatch_color: str = "#999999"
    fill_hatch_style: str | None = None
    fill_alpha: float = 0.0


@dataclass(slots=True)
class CharShape:
    id: int
    font_face: dict[str, int]
    basesize: int
    text_color: str
    shade_color: str
    bold: bool
    italic: bool
    letter_spacing: dict[str, int]
    letter_width_expansion: dict[str, int]
    relative_size: dict[str, int]
    position: dict[str, int]
    border_fill_id: int


@dataclass(slots=True)
class TabDef:
    id: int
    auto_tab_left: bool = False
    auto_tab_right: bool = False


@dataclass(slots=True)
class ParaShape:
    id: int
    tabdef_id: int
    borderfill_id: int
    indent: int
    margin_left: int
    margin_right: int
    margin_top: int
    margin_bottom: int
    line_spacing: int
    align: str = "JUSTIFY"
    vertical_align: str = "BASELINE"
    heading_type: str = "NONE"
    heading_id_ref: int = 0
    heading_level: int = 0
    condense: int = 0
    font_line_height: bool = False
    snap_to_grid: bool = True
    suppress_line_numbers: bool = False
    break_latin_word: str = "KEEP_WORD"
    break_non_latin_word: str = "BREAK_WORD"
    widow_orphan: bool = False
    keep_with_next: bool = False
    keep_lines: bool = False
    page_break_before: bool = False
    line_wrap: str = "BREAK"
    line_spacing_type: str = "PERCENT"
    auto_spacing_easian_eng: bool = False
    auto_spacing_easian_num: bool = False
    border_offset_left: int = 0
    border_offset_right: int = 0
    border_offset_top: int = 0
    border_offset_bottom: int = 0
    border_connect: bool = False
    border_ignore_margin: bool = False


@dataclass(slots=True)
class Style:
    id: int
    local_name: str
    name: str
    kind: str
    para_shape_id: int
    char_shape_id: int
    next_style_id: int
    lang_id: int


@dataclass(slots=True)
class LineSeg:
    textpos: int
    vertpos: int
    vertsize: int
    textheight: int
    baseline: int
    spacing: int
    horzpos: int
    horzsize: int
    flags: int


@dataclass(slots=True)
class PageDef:
    width: int
    height: int
    left: int
    right: int
    top: int
    bottom: int
    header: int
    footer: int
    gutter: int


@dataclass(slots=True)
class NoteShape:
    suffix: str
    starting_number: int
    splitter_length: int
    splitter_margin_top: int
    splitter_margin_bottom: int
    notes_spacing: int
    splitter_stroke_type: int
    splitter_width: int
    splitter_color: str


@dataclass(slots=True)
class PageBorderFill:
    apply_type: str
    borderfill_id: int
    left: int
    right: int
    top: int
    bottom: int


@dataclass(slots=True)
class SectionDef:
    columnspacing: int
    grid_vertical: int
    grid_horizontal: int
    default_tab_stops: int
    numbering_shape_id: int
    starting_pagenum: int
    starting_picturenum: int
    starting_tablenum: int
    starting_equationnum: int
    page_def: PageDef | None = None
    footnote_shape: NoteShape | None = None
    endnote_shape: NoteShape | None = None
    page_border_fills: list[PageBorderFill] = field(default_factory=list)


@dataclass(slots=True)
class ColumnsDef:
    spacing: int
    col_count: int = 1


@dataclass(slots=True)
class PageNum:
    position: str = "BOTTOM_CENTER"
    format_type: str = "DIGIT"
    side_char: str = "-"


@dataclass(slots=True)
class PageHide:
    hide_header: bool = False
    hide_footer: bool = False
    hide_master_page: bool = False
    hide_border: bool = False
    hide_fill: bool = False
    hide_page_num: bool = False


@dataclass(slots=True)
class TableCell:
    name: str = ""
    header: bool = False
    has_margin: bool = False
    protect: bool = False
    editable: bool = False
    border_fill_id: int = 1
    col_addr: int = 0
    row_addr: int = 0
    col_span: int = 1
    row_span: int = 1
    width: int = 0
    height: int = 0
    padding_left: int = 0
    padding_right: int = 0
    padding_top: int = 0
    padding_bottom: int = 0
    vertical_align: str = "CENTER"
    paragraphs: list["Paragraph"] = field(default_factory=list)


@dataclass(slots=True)
class Table:
    instance_id: int
    z_order: int
    width: int
    height: int
    margin_left: int
    margin_right: int
    margin_top: int
    margin_bottom: int
    rows: int
    cols: int
    cell_spacing: int
    border_fill_id: int
    numbering_type: str = "TABLE"
    text_wrap: str = "TOP_AND_BOTTOM"
    text_flow: str = "BOTH_SIDES"
    repeat_header: bool = True
    page_break: str = "TABLE"
    treat_as_char: bool = True
    affect_line_spacing: bool = False
    flow_with_text: bool = True
    allow_overlap: bool = False
    hold_anchor_and_so: bool = False
    vert_rel_to: str = "PARA"
    horz_rel_to: str = "PARA"
    vert_align: str = "TOP"
    horz_align: str = "LEFT"
    vert_offset: int = 0
    horz_offset: int = 0
    width_rel_to: str = "ABSOLUTE"
    height_rel_to: str = "ABSOLUTE"
    padding_left: int = 0
    padding_right: int = 0
    padding_top: int = 0
    padding_bottom: int = 0
    description: str = ""
    cells: list[TableCell] = field(default_factory=list)


@dataclass(slots=True)
class Equation:
    instance_id: int
    z_order: int
    width: int
    height: int
    margin_left: int
    margin_right: int
    margin_top: int
    margin_bottom: int
    numbering_type: str = "EQUATION"
    text_wrap: str = "TOP_AND_BOTTOM"
    text_flow: str = "BOTH_SIDES"
    treat_as_char: bool = True
    affect_line_spacing: bool = False
    flow_with_text: bool = True
    allow_overlap: bool = False
    hold_anchor_and_so: bool = False
    vert_rel_to: str = "PARA"
    horz_rel_to: str = "PARA"
    vert_align: str = "TOP"
    horz_align: str = "LEFT"
    vert_offset: int = 0
    horz_offset: int = 0
    width_rel_to: str = "ABSOLUTE"
    height_rel_to: str = "ABSOLUTE"
    version: str = ""
    base_line: int = 0
    text_color: str = "#000000"
    base_unit: int = 1000
    line_mode: str = "CHAR"
    font: str = ""
    script: str = ""
    description: str = ""


@dataclass(slots=True)
class FieldBegin:
    instance_id: int
    field_type: str = "BOOKMARK"
    name: str = ""
    editable: bool = False
    dirty: bool = True
    field_id: str = ""
    command: str = ""


@dataclass(slots=True)
class FieldEnd:
    begin_id_ref: int
    field_id: str = ""


@dataclass(slots=True)
class Piece:
    kind: str
    value: str | SectionDef | ColumnsDef | PageNum | PageHide | Table | Equation | FieldBegin | FieldEnd | None = None


@dataclass(slots=True)
class Run:
    char_shape_id: int
    pieces: list[Piece] = field(default_factory=list)


@dataclass(slots=True)
class Paragraph:
    instance_id: int
    para_shape_id: int
    style_id: int
    split: int
    page_break: bool = False
    column_break: bool = False
    runs: list[Run] = field(default_factory=list)
    line_segs: list[LineSeg] = field(default_factory=list)


@dataclass(slots=True)
class Section:
    index: int
    paragraphs: list[Paragraph] = field(default_factory=list)


@dataclass(slots=True)
class CaretPosition:
    list_id: int
    paragraph_id: int
    position_in_paragraph: int


@dataclass(slots=True)
class HwpDocument:
    source: Path
    version: tuple[int, int, int, int]
    metadata: DocumentMetadata
    caret: CaretPosition
    section_count: int
    font_faces: dict[str, list[FontFace]]
    border_fills: list[BorderFill]
    char_shapes: list[CharShape]
    tab_defs: list[TabDef]
    para_shapes: list[ParaShape]
    styles: list[Style]
    sections: list[Section]
    warnings: list[str] = field(default_factory=list)
