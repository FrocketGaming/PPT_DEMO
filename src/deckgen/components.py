"""Slide components. Each one draws itself into a Box.

To add a new component type:

    @register("my_type")
    def render_my_type(slide, box: Box, spec: dict, ctx: Context) -> None:
        ...

and use `type: my_type` in the YAML.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from . import data as dq
from .layout import Box
from .theme import Theme


@dataclass
class Context:
    theme: Theme
    data: dict[str, list[dict]]
    base_dir: Path


Renderer = Callable[[object, Box, dict, Context], None]
REGISTRY: dict[str, Renderer] = {}


def register(name: str):
    def deco(fn: Renderer) -> Renderer:
        REGISTRY[name] = fn
        return fn
    return deco


def render(slide, box: Box, spec: dict, ctx: Context) -> None:
    kind = spec.get("type")
    if kind not in REGISTRY:
        raise ValueError(f"unknown component type '{kind}'; available: {sorted(REGISTRY)}")
    REGISTRY[kind](slide, box, spec, ctx)


# --------------------------------------------------------------------------- helpers

ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}


def _emu(box: Box) -> tuple[Emu, Emu, Emu, Emu]:
    return Inches(box.x), Inches(box.y), Inches(box.w), Inches(box.h)


def add_rich_text(paragraph, text: str, theme: Theme, size: float | None,
                  color: str | None = "text", bold: bool = False) -> None:
    """Adds text to a paragraph; **double asterisks** become bold accent-colored runs.

    size/color None leave the run unstyled so it inherits from a template placeholder.
    """
    inherit = size is None
    for i, chunk in enumerate(re.split(r"\*\*(.+?)\*\*", str(text))):
        if not chunk:
            continue
        run = paragraph.add_run()
        run.text = chunk
        emphasized = i % 2 == 1
        if not inherit:
            run.font.name = theme.font
            run.font.size = Pt(size)
        if bold or emphasized:
            run.font.bold = True
        if emphasized:
            run.font.color.rgb = theme.rgb("accent")
        elif color is not None:
            run.font.color.rgb = theme.rgb(color)


def _textbox(slide, box: Box, anchor=MSO_ANCHOR.TOP, margin: float = 0.05):
    tb = slide.shapes.add_textbox(*_emu(box))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", Inches(margin))
    return tf


def _bullet(paragraph, theme: Theme, numbered: bool) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(Inches(0.28)))
    pPr.set("indent", str(-Inches(0.28)))
    clr = pPr.makeelement(qn("a:buClr"), {})
    srgb = clr.makeelement(qn("a:srgbClr"), {"val": str(theme.rgb("accent"))})
    clr.append(srgb)
    pPr.append(clr)
    if numbered:
        pPr.append(pPr.makeelement(qn("a:buAutoNum"), {"type": "arabicPeriod"}))
    else:
        pPr.append(pPr.makeelement(qn("a:buChar"), {"char": "•"}))


def _fill(shape, color) -> None:
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def fmt(value, pattern: str | None) -> str:
    if pattern is None:
        if isinstance(value, float):
            return f"{value:,.2f}"
        if isinstance(value, int):
            return f"{value:,}"
        return str(value)
    return pattern.format(value)


# --------------------------------------------------------------------------- text components

@register("insights")
def render_insights(slide, box: Box, spec: dict, ctx: Context) -> None:
    """Bullet list of takeaways. style: bullets (default) | numbered | callout."""
    theme = ctx.theme
    style = spec.get("style", "bullets")
    size = spec.get("font_size", 16)
    inner = box
    if style == "callout":
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, *_emu(box))
        _fill(bg, theme.rgb("light"))
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, *_emu(Box(box.x, box.y, 0.08, box.h)))
        _fill(bar, theme.rgb("accent"))
        inner = Box(box.x + 0.3, box.y + 0.2, box.w - 0.5, box.h - 0.4)

    tf = _textbox(slide, inner, anchor=MSO_ANCHOR.MIDDLE if style == "callout" else MSO_ANCHOR.TOP)
    first = True
    if heading := spec.get("heading"):
        p = tf.paragraphs[0]
        add_rich_text(p, heading, theme, size + 2, color="primary", bold=True)
        p.space_after = Pt(8)
        first = False
    for item in spec.get("items", []):
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(spec.get("spacing", 10))
        _bullet(p, theme, numbered=style == "numbered")
        add_rich_text(p, item, theme, size)


@register("text")
def render_text(slide, box: Box, spec: dict, ctx: Context) -> None:
    tf = _textbox(slide, box, anchor={"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE,
                                      "bottom": MSO_ANCHOR.BOTTOM}[spec.get("valign", "top")])
    for i, para in enumerate(str(spec.get("text", "")).strip().split("\n\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ALIGN[spec.get("align", "left")]
        p.space_after = Pt(8)
        add_rich_text(p, para.replace("\n", " "), ctx.theme, spec.get("font_size", 16),
                      color=spec.get("color", "text"), bold=spec.get("bold", False))


@register("kpis")
def render_kpis(slide, box: Box, spec: dict, ctx: Context) -> None:
    """Row of metric cards. `value`/`delta` may be literals or computed metric specs."""
    theme = ctx.theme
    items = spec.get("items", [])
    cards = box.split_cols([1] * len(items), gap=0.25)
    for item, card in zip(items, cards):
        bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, *_emu(card))
        bg.adjustments[0] = 0.08
        _fill(bg, theme.rgb("light"))

        value = _resolve_metric(item.get("value"), ctx)
        delta = _resolve_metric(item.get("delta"), ctx)
        tf = _textbox(slide, Box(card.x + 0.2, card.y + 0.1, card.w - 0.4, card.h - 0.2),
                      anchor=MSO_ANCHOR.MIDDLE)
        p = tf.paragraphs[0]
        add_rich_text(p, item.get("label", ""), theme, 12, color="muted")
        p = tf.add_paragraph()
        add_rich_text(p, fmt(value, item.get("format")), theme,
                      item.get("value_size", spec.get("value_size", 30)),
                      color="primary", bold=True)
        if delta is not None:
            text = fmt(delta, item.get("delta_format"))
            is_up = not str(text).lstrip().startswith("-")
            good = is_up if item.get("higher_is_better", True) else not is_up
            p = tf.add_paragraph()
            add_rich_text(p, ("▲ " if is_up else "▼ ") + text.lstrip("+-") +
                          (f"  {item['delta_label']}" if item.get("delta_label") else ""),
                          theme, 12, color="positive" if good else "negative")


def _resolve_metric(value, ctx: Context):
    if isinstance(value, dict):
        rows = _source_rows(value["source"], ctx)
        return dq.metric(rows, value)
    return value


# --------------------------------------------------------------------------- data components

def _source_rows(name: str, ctx: Context) -> list[dict]:
    if name not in ctx.data:
        raise ValueError(f"unknown data source '{name}'; defined: {sorted(ctx.data)}")
    return ctx.data[name]


def _chart_series(spec: dict, ctx: Context) -> tuple[list, dict[str, list]]:
    """Returns (categories, {series_name: values}) from inline data or a named source."""
    src = spec.get("data")
    if isinstance(src, dict):  # inline: {categories: [...], series: {name: [...]}}
        return src["categories"], src["series"]

    rows = dq.query(_source_rows(src, ctx), spec)
    x = spec["x"]
    ys = spec["y"] if isinstance(spec["y"], list) else [spec["y"]]
    if series_by := spec.get("series_by"):
        dq.check_columns(rows, [x, ys[0], series_by], src)
        return dq.pivot(rows, x, ys[0], series_by)
    dq.check_columns(rows, [x, *ys], src)
    labels = spec.get("series_names", {})
    return [r[x] for r in rows], {labels.get(y, y.replace("_", " ").title()): [r[y] for r in rows]
                                  for y in ys}


CHART_TYPES = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "stacked_column": XL_CHART_TYPE.COLUMN_STACKED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "stacked_bar": XL_CHART_TYPE.BAR_STACKED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "area": XL_CHART_TYPE.AREA,
    "pie": XL_CHART_TYPE.PIE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
    "scatter": XL_CHART_TYPE.XY_SCATTER,
}
LEGEND = {"bottom": XL_LEGEND_POSITION.BOTTOM, "right": XL_LEGEND_POSITION.RIGHT,
          "top": XL_LEGEND_POSITION.TOP, "left": XL_LEGEND_POSITION.LEFT}


@register("chart")
def render_chart(slide, box: Box, spec: dict, ctx: Context) -> None:
    """Native (editable) PowerPoint chart. kind: see CHART_TYPES."""
    theme = ctx.theme
    kind = spec.get("kind", "column")
    if kind not in CHART_TYPES:
        raise ValueError(f"unknown chart kind '{kind}'; available: {sorted(CHART_TYPES)}")
    number_format = spec.get("number_format", "General")

    if kind == "scatter":
        rows = dq.query(_source_rows(spec["data"], ctx), spec)
        dq.check_columns(rows, [spec["x"], spec["y"]], spec["data"])
        cd = XyChartData()
        groups = {}
        for r in rows:
            groups.setdefault(r[spec["series_by"]] if spec.get("series_by") else spec["y"], []).append(r)
        for name, members in groups.items():
            s = cd.add_series(str(name), number_format=number_format)
            for r in members:
                s.add_data_point(r[spec["x"]], r[spec["y"]])
    else:
        categories, series = _chart_series(spec, ctx)
        cd = CategoryChartData(number_format=number_format)
        cd.categories = categories
        for name, values in series.items():
            cd.add_series(name, values)

    chart = slide.shapes.add_chart(CHART_TYPES[kind], *_emu(box), cd).chart
    chart.font.name = theme.font
    chart.font.size = Pt(spec.get("font_size", 11))
    chart.font.color.rgb = theme.rgb("text")

    if title := spec.get("title"):
        chart.has_title = True
        chart.chart_title.text_frame.text = title
        run = chart.chart_title.text_frame.paragraphs[0].runs[0]
        run.font.size, run.font.bold = Pt(13), True
        run.font.color.rgb = theme.rgb("primary")
    else:
        chart.has_title = False

    multi_series = len(chart.plots[0].series) > 1
    legend = spec.get("legend", "bottom" if multi_series or kind in ("pie", "doughnut") else "none")
    chart.has_legend = legend != "none"
    if chart.has_legend:
        chart.legend.position = LEGEND[legend]
        chart.legend.include_in_layout = False

    plot = chart.plots[0]
    if kind in ("pie", "doughnut"):
        for i, point in enumerate(plot.series[0].points):
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = theme.series_color(i)
    else:
        for i, s in enumerate(plot.series):
            color = theme.series_color(i)
            if kind in ("line", "scatter"):
                s.format.line.color.rgb = color
                s.format.line.width = Pt(2.25)
                s.marker.format.fill.solid()
                s.marker.format.fill.fore_color.rgb = color
                s.marker.format.line.color.rgb = color
                s.smooth = False
                if kind == "scatter":
                    s.format.line.fill.background()
                    s.marker.size = 10
            else:
                s.format.fill.solid()
                s.format.fill.fore_color.rgb = color
        if kind in ("column", "bar", "stacked_column", "stacked_bar"):
            plot.gap_width = spec.get("gap_width", 60)
            if kind.startswith("stacked"):
                plot.overlap = 100
        va = chart.value_axis
        va.has_major_gridlines = spec.get("gridlines", True)
        if va.has_major_gridlines:
            va.major_gridlines.format.line.color.rgb = theme.rgb("light")
        va.format.line.fill.background()
        va.tick_labels.number_format = spec.get("axis_format", number_format)
        va.tick_labels.number_format_is_linked = False
        ca = chart.category_axis
        ca.format.line.color.rgb = theme.rgb("muted")
        ca.has_major_gridlines = False
        if kind in ("bar", "stacked_bar"):
            ca.reverse_order = True  # keep first row at the top

    if spec.get("data_labels"):
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.number_format = spec.get("label_format", number_format)
        dl.number_format_is_linked = False
        dl.font.size = Pt(10)
        if kind in ("pie", "doughnut"):
            dl.show_percentage = spec.get("show_percentage", False)


@register("table")
def render_table(slide, box: Box, spec: dict, ctx: Context) -> None:
    theme = ctx.theme
    rows = dq.query(_source_rows(spec["data"], ctx), spec)
    columns = spec.get("columns") or list(rows[0])
    if isinstance(columns, list):
        columns = {c: c.replace("_", " ").title() for c in columns}
    dq.check_columns(rows, list(columns), spec["data"])
    formats = spec.get("formats", {})
    size = spec.get("font_size", 12)

    row_h = min(0.45, box.h / (len(rows) + 1))
    shape = slide.shapes.add_table(len(rows) + 1, len(columns), Inches(box.x), Inches(box.y),
                                   Inches(box.w), Inches(row_h * (len(rows) + 1)))
    table = shape.table
    highlight = spec.get("highlight_top", 0)

    def write(cell, text, *, header=False, numeric=False, row_i=0):
        cell.text = ""
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT if numeric else PP_ALIGN.LEFT
        add_rich_text(p, text, theme, size, color="light" if header else "text",
                      bold=header or row_i < highlight)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = cell.margin_right = Inches(0.1)
        cell.fill.solid()
        if header:
            cell.fill.fore_color.rgb = theme.rgb("primary")
        else:
            cell.fill.fore_color.rgb = theme.rgb("light") if row_i % 2 else theme.rgb("FFFFFF")

    numeric_cols = {c for c in columns if isinstance(rows[0][c], (int, float))}
    for j, (col, header) in enumerate(columns.items()):
        write(table.cell(0, j), header, header=True, numeric=col in numeric_cols)
    for i, row in enumerate(rows, start=1):
        for j, col in enumerate(columns):
            write(table.cell(i, j), fmt(row[col], formats.get(col)),
                  numeric=col in numeric_cols, row_i=i - 1)


@register("image")
def render_image(slide, box: Box, spec: dict, ctx: Context) -> None:
    path = (ctx.base_dir / spec["path"]).resolve()
    if not path.exists():
        raise FileNotFoundError(f"image not found: {path}")
    slide.shapes.add_picture(str(path), Inches(box.x), Inches(box.y), height=Inches(box.h))
