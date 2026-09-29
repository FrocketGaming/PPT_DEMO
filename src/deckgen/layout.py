"""Turns region names (left/right/top/bottom/...) into rectangles on the slide.

The content area is split into up to three horizontal bands:

    +-----------------------------+
    |            top              |   <- full width
    +---------+---------+---------+
    |  left   | center  |  right  |   <- "middle" band, columns
    +---------+---------+---------+
    |           bottom            |   <- full width
    +-----------------------------+

- `top` / `bottom` take `size` as a fraction of the height (default 0.3)
  when a middle band exists; otherwise they split the height by weight.
- `left` / `center` / `right` take `size` as a relative width weight
  (default 1), e.g. left: 1, right: 2 -> one third / two thirds.
- `full` takes the whole content area and cannot be combined.
"""

from __future__ import annotations

from dataclasses import dataclass

COLUMNS = ("left", "center", "right")
BANDS = ("top", "bottom")
VALID_REGIONS = {*COLUMNS, *BANDS, "full"}


@dataclass(frozen=True)
class Box:
    """A rectangle in inches."""

    x: float
    y: float
    w: float
    h: float

    def split_rows(self, weights: list[float], gap: float) -> list[Box]:
        usable = self.h - gap * (len(weights) - 1)
        total = sum(weights)
        boxes, y = [], self.y
        for wt in weights:
            h = usable * wt / total
            boxes.append(Box(self.x, y, self.w, h))
            y += h + gap
        return boxes

    def split_cols(self, weights: list[float], gap: float) -> list[Box]:
        usable = self.w - gap * (len(weights) - 1)
        total = sum(weights)
        boxes, x = [], self.x
        for wt in weights:
            w = usable * wt / total
            boxes.append(Box(x, self.y, w, self.h))
            x += w + gap
        return boxes


def _size(spec, default: float) -> float:
    if isinstance(spec, dict) and "size" in spec:
        return float(spec["size"])
    return default


def compute_regions(regions: dict, area: Box, gap: float = 0.3) -> dict[str, Box]:
    names = list(regions)
    unknown = [n for n in names if n not in VALID_REGIONS]
    if unknown:
        raise ValueError(
            f"unknown region(s) {unknown}; valid regions are {sorted(VALID_REGIONS)}"
        )
    if "full" in names:
        if len(names) > 1:
            raise ValueError("'full' cannot be combined with other regions")
        return {"full": area}

    cols = [n for n in COLUMNS if n in names]
    rows: list[tuple[str, float]] = []
    if cols:
        fractions = {n: _size(regions[n], 0.3) for n in BANDS if n in names}
        middle = 1.0 - sum(fractions.values())
        if middle <= 0.05:
            raise ValueError("top/bottom sizes leave no room for left/center/right")
        if "top" in fractions:
            rows.append(("top", fractions["top"]))
        rows.append(("middle", middle))
        if "bottom" in fractions:
            rows.append(("bottom", fractions["bottom"]))
    else:
        rows = [(n, _size(regions[n], 1.0)) for n in BANDS if n in names]

    result: dict[str, Box] = {}
    row_boxes = area.split_rows([w for _, w in rows], gap)
    for (name, _), box in zip(rows, row_boxes):
        if name == "middle":
            col_boxes = box.split_cols([_size(regions[c], 1.0) for c in cols], gap)
            result.update(zip(cols, col_boxes))
        else:
            result[name] = box
    return result
