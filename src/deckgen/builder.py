"""Reads a deck YAML and produces a .pptx. All slide content comes from the config."""

from __future__ import annotations

import copy
from pathlib import Path

import yaml
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches, Pt

from . import components
from .components import Context, add_rich_text, render
from .data import load_sources
from .layout import Box, compute_regions
from .theme import Theme

SLIDE_W, SLIDE_H = 13.333, 7.5
MARGIN = 0.5
CONTENT = Box(MARGIN, 1.55, SLIDE_W - 2 * MARGIN, SLIDE_H - 1.55 - 0.6)
REGION_KEYS = ("left", "center", "right", "top", "bottom", "full")
SLIDE_KEYS = {"type", "title", "subtitle", "source", "notes", "skip", "gap", "regions",
              "font_size", *REGION_KEYS}


class ConfigError(Exception):
    pass


def _merge_defaults(spec: dict, defaults: dict) -> dict:
    merged = copy.deepcopy(defaults.get(spec.get("type"), {}))
    merged.update(spec)
    return merged


def _title_slide(prs, slide_cfg: dict, theme: Theme, dark: bool = True):
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


def _content_slide(prs, slide_cfg: dict, ctx: Context, defaults: dict, number: int):
    theme = ctx.theme
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    tf = components._textbox(slide, Box(MARGIN, 0.35, SLIDE_W - 2 * MARGIN, 0.75),
                             anchor=MSO_ANCHOR.BOTTOM)
    add_rich_text(tf.paragraphs[0], slide_cfg.get("title", ""), theme, 28, color="primary",
                  bold=True)
    if sub := slide_cfg.get("subtitle"):
        tf = components._textbox(slide, Box(MARGIN, 1.08, SLIDE_W - 2 * MARGIN, 0.4))
        add_rich_text(tf.paragraphs[0], sub, theme, 15, color="muted")

    # footer: source note + slide number
    if src := slide_cfg.get("source"):
        tf = components._textbox(slide, Box(MARGIN, SLIDE_H - 0.45, SLIDE_W - 2 * MARGIN - 1, 0.3))
        add_rich_text(tf.paragraphs[0], f"Source: {src}", theme, 10, color="muted")
    tf = components._textbox(slide, Box(SLIDE_W - MARGIN - 0.6, SLIDE_H - 0.45, 0.6, 0.3))
    tf.paragraphs[0].alignment = components.ALIGN["right"]
    add_rich_text(tf.paragraphs[0], str(number), theme, 10, color="muted")

    regions = {k: slide_cfg[k] for k in REGION_KEYS if k in slide_cfg}
    regions.update(slide_cfg.get("regions", {}))
    if not regions:
        raise ValueError("content slide has no regions (use left/right/top/bottom/center/full)")
    boxes = compute_regions(regions, CONTENT, gap=slide_cfg.get("gap", 0.3))

    for name, region in regions.items():
        try:
            specs = region["stack"] if "stack" in region else [region]
            weights = [s.get("weight", 1) for s in specs]
            for spec, box in zip(specs, boxes[name].split_rows(weights, gap=0.2)):
                render(slide, box, _merge_defaults(spec, defaults), ctx)
        except Exception as exc:
            raise ValueError(f"region '{name}': {exc}") from exc
    return slide


def build(config_path: str | Path, output: str | Path | None = None) -> Path:
    config_path = Path(config_path).resolve()
    cfg = yaml.safe_load(config_path.read_text(encoding="utf-8"))
    base_dir = config_path.parent

    deck = cfg.get("deck", {})
    theme = Theme.from_config(deck.get("theme"))
    ctx = Context(theme=theme, data=load_sources(cfg.get("data", {}), base_dir), base_dir=base_dir)
    defaults = cfg.get("defaults", {})

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(SLIDE_W), Inches(SLIDE_H)
    prs.core_properties.title = deck.get("title", "")
    prs.core_properties.author = deck.get("author", "")

    for i, slide_cfg in enumerate(cfg.get("slides", []), start=1):
        if slide_cfg.get("skip"):
            continue
        kind = slide_cfg.get("type", "content")
        try:
            if unknown := sorted(set(slide_cfg) - SLIDE_KEYS):
                raise ValueError(f"unknown key(s) {unknown}; valid: {sorted(SLIDE_KEYS)}")
            if kind == "title":
                slide = _title_slide(prs, slide_cfg, theme)
            elif kind == "section":
                slide = _title_slide(prs, slide_cfg, theme, dark=False)
            elif kind == "content":
                slide = _content_slide(prs, slide_cfg, ctx, defaults, len(prs.slides) + 1)
            else:
                raise ValueError(f"unknown slide type '{kind}' (title, section, content)")
        except Exception as exc:
            raise ConfigError(f"slide {i} ('{slide_cfg.get('title', '')}'): {exc}") from exc
        if notes := slide_cfg.get("notes"):
            slide.notes_slide.notes_text_frame.text = notes

    output = Path(output) if output else base_dir.parent / "out" / f"{config_path.stem}.pptx"
    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output)
    return output
