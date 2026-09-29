"""Theme defaults. Anything here can be overridden under `deck.theme` in the YAML."""

from __future__ import annotations

from dataclasses import dataclass, field, fields, replace

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
    subtle: str = "C5CDD6"  # de-emphasised series next to a `highlight`
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

    def hex(self, name_or_hex: str) -> str:
        value = getattr(self, name_or_hex, name_or_hex)
        return str(value).lstrip("#").upper()

    def rgb(self, name_or_hex: str) -> RGBColor:
        return RGBColor.from_string(self.hex(name_or_hex))

    def mix(self, a: str, b: str, t: float) -> str:
        """Hex colour t of the way from colour a to colour b (names or hex)."""
        ca, cb = self.rgb(a), self.rgb(b)
        return "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(ca, cb))

    def is_dark(self, name_or_hex: str) -> bool:
        r, g, b = self.rgb(name_or_hex)
        return 0.299 * r + 0.587 * g + 0.114 * b < 140

    def on_background(self, bg: str) -> Theme:
        """The theme to draw with on a slide whose background is `bg`.

        On a dark background, text turns white and cards become a lighter tint of it.
        """
        bg = self.hex(bg)
        if not self.is_dark(bg):
            return replace(self, light=self.mix(bg, "text", 0.07))
        return replace(
            self, text="FFFFFF", primary="FFFFFF", muted=self.mix(bg, "FFFFFF", 0.7),
            light=self.mix(bg, "FFFFFF", 0.1), subtle=self.mix(bg, "FFFFFF", 0.3),
            palette=["FFFFFF" if c.lstrip("#").upper() == bg else c for c in self.palette],
        )

    def series_color(self, i: int) -> RGBColor:
        return RGBColor.from_string(self.palette[i % len(self.palette)].lstrip("#").upper())
