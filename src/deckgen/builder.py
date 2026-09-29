"""Reads a deck YAML and produces a .pptx. All slide content comes from the config.

Two modes:
- no `deck.template`: slides are drawn from scratch with the built-in theme
- `deck.template: <name>`: slides use the brand template's layouts, placeholders,
  colours and fonts; regions are laid out inside the template's content area
"""

from __future__ import annotations

import copy
import sys
from pathlib import Path

import yaml
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Emu, Inches

from . import components
from .components import Context, add_rich_text, render
from .data import load_sources
from .layout import Box, compute_regions
from .template import Template, default_date, find_template
from .theme import Theme

SLIDE_W, SLIDE_H = 13.333, 7.5
MARGIN = 0.5
CONTENT = Box(MARGIN, 1.55, SLIDE_W - 2 * MARGIN, SLIDE_H - 1.55 - 0.6)
REGION_KEYS = ("left", "center", "right", "top", "bottom", "full")
SLIDE_TYPES = ("title", "section", "content", "closing")
SLIDE_KEYS = {"type", "title", "subtitle", "source", "notes", "skip", "gap", "regions",
              "font_size", "layout", "placeholders", "tag", "date", *REGION_KEYS}


class ConfigError(Exception):
    pass


def warn(msg: str) -> None:
    print(f"warning: {msg}", file=sys.stderr)


def _merge_defaults(spec: dict, defaults: dict) -> dict:
    merged = copy.deepcopy(defaults.get(spec.get("type"), {}))
    merged.update(spec)
    return merged


def _regions(slide_cfg: dict) -> dict:
    regions = {k: slide_cfg[k] for k in REGION_KEYS if k in slide_cfg}
    regions.update(slide_cfg.get("regions", {}))
    return regions


def _render_regions(slide, regions: dict, area: Box, ctx: Context, defaults: dict,
                    gap: float) -> None:
    boxes = compute_regions(regions, area, gap=gap)
    for name, region in regions.items():
        try:
            specs = region["stack"] if "stack" in region else [region]
            weights = [s.get("weight", 1) for s in specs]
            for spec, box in zip(specs, boxes[name].split_rows(weights, gap=0.2)):
                render(slide, box, _merge_defaults(spec, defaults), ctx)
        except Exception as exc:
            raise ValueError(f"region '{name}': {exc}") from exc


def _small_text(slide, box: Box, text: str, theme: Theme, size: float = 10,
                color: str = "muted", align: str = "left") -> None:
    tf = components._textbox(slide, box)
    tf.paragraphs[0].alignment = components.ALIGN[align]
    add_rich_text(tf.paragraphs[0], text, theme, size, color=color)


# --------------------------------------------------------------------------- built-in look

def _plain_title_slide(prs, slide_cfg: dict, theme: Theme, dark: bool = True):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = theme.rgb("primary" if dark else "light")
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(MARGIN), Inches(3.55),
                                  Inches(1.2), Inches(0.08))
    band.fill.solid()
    band.fill.fore_color.rgb = theme.rgb("accent")
    band.line.fill.background()

    tf = components._textbox(slide, Box(MARGIN, 1.5, SLIDE_W - 2 * MARGIN, 1.95),
                             anchor=MSO_ANCHOR.BOTTOM)
    add_rich_text(tf.paragraphs[0], slide_cfg.get("title", ""), theme,
                  slide_cfg.get("font_size", 40), color="FFFFFF" if dark else "primary", bold=True)
    if sub := slide_cfg.get("subtitle"):
        tf2 = components._textbox(slide, Box(MARGIN, 3.85, SLIDE_W - 2 * MARGIN, 1.5))
        add_rich_text(tf2.paragraphs[0], sub, theme, 20, color="D0D7E1" if dark else "muted")
    return slide


def _plain_content_slide(prs, slide_cfg: dict, ctx: Context, defaults: dict, number: int):
    theme = ctx.theme
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    tf = components._textbox(slide, Box(MARGIN, 0.35, SLIDE_W - 2 * MARGIN, 0.75),
                             anchor=MSO_ANCHOR.BOTTOM)
    add_rich_text(tf.paragraphs[0], slide_cfg.get("title", ""), theme, 28, color="primary",
                  bold=True)
    if sub := slide_cfg.get("subtitle"):
        _small_text(slide, Box(MARGIN, 1.08, SLIDE_W - 2 * MARGIN, 0.4), sub, theme, 15)
    if src := slide_cfg.get("source"):
        _small_text(slide, Box(MARGIN, SLIDE_H - 0.45, SLIDE_W - 2 * MARGIN - 1, 0.3),
                    f"Source: {src}", theme)
    _small_text(slide, Box(SLIDE_W - MARGIN - 0.6, SLIDE_H - 0.45, 0.6, 0.3), str(number),
                theme, align="right")

    regions = _regions(slide_cfg)
    if not regions:
        raise ValueError("content slide has no regions (use left/right/top/bottom/center/full)")
    _render_regions(slide, regions, CONTENT, ctx, defaults, slide_cfg.get("gap", 0.3))
    return slide


# --------------------------------------------------------------------------- template look

def _fill_text(placeholder, value, theme: Theme) -> None:
    """Writes text into a placeholder, keeping the template's fonts, sizes and bullets."""
    lines = value if isinstance(value, list) else str(value).split("\n")
    tf = placeholder.text_frame
    tf.clear()
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        add_rich_text(p, line, theme, size=None, color=None)


def _fill_placeholders(slide, fills: dict, ctx: Context, defaults: dict) -> None:
    """`placeholders: {idx: text | [lines] | component}` targets the layout's own boxes."""
    by_idx = {ph.placeholder_format.idx: ph for ph in slide.placeholders}
    for idx, value in fills.items():
        if (ph := by_idx.get(int(idx))) is None:
            raise ValueError(f"placeholder {idx} not on layout '{slide.slide_layout.name}'; "
                             f"available: {sorted(by_idx)}")
        if not isinstance(value, dict):
            _fill_text(ph, value, ctx.theme)
            continue
        spec = _merge_defaults(value, defaults)
        if spec.get("type") == "image" and hasattr(ph, "insert_picture"):
            ph.insert_picture(str((ctx.base_dir / spec["path"]).resolve()))
            continue
        box = Box(*(Emu(v).inches for v in (ph.left, ph.top, ph.width, ph.height)))
        ph._element.getparent().remove(ph._element)
        try:
            render(slide, box, spec, ctx)
        except Exception as exc:
            raise ValueError(f"placeholder {idx}: {exc}") from exc


def _template_slide(tpl: Template, kind: str, slide_cfg: dict, ctx: Context, defaults: dict,
                    section: str | None, deck: dict):
    slide, roles = tpl.add_slide(kind, slide_cfg)
    values = {
        "title": slide_cfg.get("title"),
        "subtitle": slide_cfg.get("subtitle"),
        "tag": slide_cfg.get("tag", section if kind == "content" else None),
        "date": slide_cfg.get("date", deck.get("date") or default_date())
        if kind == "title" else None,
    }
    for role, value in values.items():
        if value is None:
            continue
        if role in roles:
            _fill_text(roles[role], value, ctx.theme)
        elif role == "title" or (role == "subtitle" and kind != "content"):
            warn(f"layout '{slide.slide_layout.name}' has no {role} placeholder; "
                 f"'{value}' not shown (map it in {tpl.path.stem}.yaml)")

    _fill_placeholders(slide, slide_cfg.get("placeholders", {}), ctx, defaults)

    regions = _regions(slide_cfg)
    if kind == "content" or regions:
        area = tpl.content_area(slide, roles.get("title"))
        if (sub := values["subtitle"]) and "subtitle" not in roles and kind == "content":
            _small_text(slide, Box(area.x, area.y, area.w, 0.4), sub, ctx.theme, 15)
            area = Box(area.x, area.y + 0.5, area.w, area.h - 0.5)
        if src := slide_cfg.get("source"):
            _small_text(slide, Box(area.x, area.y + area.h + 0.08, area.w, 0.28),
                        f"Source: {src}", ctx.theme)
        if regions:
            _render_regions(slide, regions, area, ctx, defaults, slide_cfg.get("gap", 0.3))
        elif not slide_cfg.get("placeholders"):
            raise ValueError("content slide has no regions or placeholders")

    tpl.remove_empty_placeholders(slide)
    return slide


# --------------------------------------------------------------------------- entry point

def _template_dirs(deck: dict, base_dir: Path) -> list[Path]:
    dirs = [base_dir / deck["templates_dir"]] if deck.get("templates_dir") else []
    return dirs + [base_dir / "templates", base_dir.parent / "templates", Path.cwd() / "templates"]


def build(config_path: str | Path, output: str | Path | None = None) -> Path:
    config_path = Path(config_path).resolve()
    cfg = yaml.safe_load(config_path.read_text(encoding="utf-8"))
    base_dir = config_path.parent
    deck = cfg.get("deck", {})

    tpl = None
    if name := deck.get("template"):
        tpl = Template(find_template(name, _template_dirs(deck, base_dir)))
        prs = tpl.prs
        theme = Theme.from_config({**tpl.theme, **(deck.get("theme") or {})})
    else:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(SLIDE_W), Inches(SLIDE_H)
        theme = Theme.from_config(deck.get("theme"))

    ctx = Context(theme=theme, data=load_sources(cfg.get("data", {}), base_dir), base_dir=base_dir)
    defaults = cfg.get("defaults", {})
    prs.core_properties.title = deck.get("title", "")
    prs.core_properties.author = deck.get("author", "")

    section = None
    for i, slide_cfg in enumerate(cfg.get("slides", []), start=1):
        if slide_cfg.get("skip"):
            continue
        kind = slide_cfg.get("type", "content")
        try:
            if unknown := sorted(set(slide_cfg) - SLIDE_KEYS):
                raise ValueError(f"unknown key(s) {unknown}; valid: {sorted(SLIDE_KEYS)}")
            if kind not in SLIDE_TYPES:
                raise ValueError(f"unknown slide type '{kind}'; valid: {list(SLIDE_TYPES)}")
            if kind == "section":
                section = slide_cfg.get("title")
            if tpl:
                slide = _template_slide(tpl, kind, slide_cfg, ctx, defaults, section, deck)
            elif slide_cfg.get("layout") or slide_cfg.get("placeholders"):
                raise ValueError("'layout'/'placeholders' need a template (set deck.template)")
            elif kind in ("title", "closing"):
                slide = _plain_title_slide(prs, slide_cfg, theme)
            elif kind == "section":
                slide = _plain_title_slide(prs, slide_cfg, theme, dark=False)
            else:
                slide = _plain_content_slide(prs, slide_cfg, ctx, defaults, len(prs.slides) + 1)
        except Exception as exc:
            raise ConfigError(f"slide {i} ('{slide_cfg.get('title', '')}'): {exc}") from exc
        if notes := slide_cfg.get("notes"):
            slide.notes_slide.notes_text_frame.text = notes

    output = Path(output) if output else base_dir.parent / "out" / f"{config_path.stem}.pptx"
    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output)
    return output
