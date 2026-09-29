"""Theme defaults. Anything here can be overridden under `deck.theme` in the YAML."""

from __future__ import annotations

from dataclasses import dataclass, field, fields

from pptx.dml.color import RGBColor


@dataclass
class Theme:
    font: str = "Calibri"
    primary: str = "1F3A5F"  # titles, title-slide background
    accent: str = "E07A5F"  # highlights, callouts
    text: str = "2B2D42"
    muted: str = "6C757D"
    light: str = "F1F4F8"  # card / callout backgrounds
    positive: str = "2A9D8F"
    negative: str = "D1495B"
    palette: list[str] = field(
        default_factory=lambda: [
            "1F3A5F", "E07A5F", "3D8BD3", "81B29A", "F2CC8F", "8D6A9F", "6C757D",
        ]
    )

    @classmethod
    def from_config(cls, cfg: dict | None) -> Theme:
        cfg = cfg or {}
        known = {f.name for f in fields(cls)}
        unknown = set(cfg) - known
        if unknown:
            raise ValueError(f"unknown theme keys {sorted(unknown)}; valid: {sorted(known)}")
        return cls(**{k: (v.lstrip("#") if isinstance(v, str) and k != "font" else v)
                      for k, v in cfg.items()})

    def rgb(self, name_or_hex: str) -> RGBColor:
        value = getattr(self, name_or_hex, name_or_hex)
        return RGBColor.from_string(str(value).lstrip("#").upper())

    def series_color(self, i: int) -> RGBColor:
        return RGBColor.from_string(self.palette[i % len(self.palette)].lstrip("#").upper())
