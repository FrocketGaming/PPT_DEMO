"""Loads data sources and applies the small query language used in the YAML.

Query keys a component may use (all optional, applied in this order):

    where:      {column: value}  or  {column: [v1, v2]}
    group_by:   column or [columns]      (with agg: sum|mean|count|min|max)
    sort_by:    column                   (with descending: true)
    limit:      N
"""

from __future__ import annotations

import csv
from pathlib import Path

Rows = list[dict]

QUERY_KEYS = ("where", "group_by", "agg", "sort_by", "descending", "limit")


def _coerce(value: str):
    try:
        f = float(value.replace(",", ""))
    except ValueError:
        return value
    return int(f) if f.is_integer() and "." not in value else f


def read_csv(path: Path) -> Rows:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        return [{k: _coerce(v) for k, v in row.items()} for row in csv.DictReader(fh)]


def load_sources(sources: dict, base_dir: Path) -> dict[str, Rows]:
    """`sources` maps a name to a CSV path (relative to the config file) or inline rows."""
    loaded = {}
    for name, src in (sources or {}).items():
        if isinstance(src, list):
            loaded[name] = src
        else:
            path = (base_dir / src).resolve()
            if not path.exists():
                raise FileNotFoundError(f"data source '{name}': {path} not found")
            loaded[name] = read_csv(path)
    return loaded


def _aggregate(values: list, how: str):
    nums = [v for v in values if isinstance(v, (int, float))]
    match how:
        case "sum":
            return sum(nums)
        case "mean":
            return sum(nums) / len(nums) if nums else 0
        case "count":
            return len(values)
        case "min":
            return min(nums)
        case "max":
            return max(nums)
    raise ValueError(f"unknown agg '{how}' (use sum, mean, count, min, max)")


def group(rows: Rows, by: list[str], how: str = "sum") -> Rows:
    buckets: dict[tuple, list[dict]] = {}
    for row in rows:
        buckets.setdefault(tuple(row[c] for c in by), []).append(row)
    out = []
    for key, members in buckets.items():
        agg_row = dict(zip(by, key))
        for col in members[0]:
            if col not in by and isinstance(members[0][col], (int, float)):
                agg_row[col] = _aggregate([m[col] for m in members], how)
        out.append(agg_row)
    return out


def query(rows: Rows, spec: dict) -> Rows:
    for col, want in (spec.get("where") or {}).items():
        allowed = want if isinstance(want, list) else [want]
        rows = [r for r in rows if r.get(col) in allowed]
    if by := spec.get("group_by"):
        rows = group(rows, [by] if isinstance(by, str) else by, spec.get("agg", "sum"))
    if key := spec.get("sort_by"):
        rows = sorted(rows, key=lambda r: r[key], reverse=bool(spec.get("descending")))
    if limit := spec.get("limit"):
        rows = rows[: int(limit)]
    return rows


def check_columns(rows: Rows, columns: list[str], source: str) -> None:
    if not rows:
        raise ValueError(f"data source '{source}' returned no rows")
    missing = [c for c in columns if c not in rows[0]]
    if missing:
        raise ValueError(
            f"column(s) {missing} not in '{source}'; available: {list(rows[0])}"
        )


def pivot(rows: Rows, x: str, y: str, series_by: str) -> tuple[list, dict[str, list]]:
    """Long -> wide: one series per distinct `series_by` value, summed."""
    categories = list(dict.fromkeys(r[x] for r in rows))
    series_names = list(dict.fromkeys(r[series_by] for r in rows))
    table = {s: {c: 0 for c in categories} for s in series_names}
    for r in rows:
        table[r[series_by]][r[x]] += r[y]
    return categories, {s: [table[s][c] for c in categories] for s in series_names}


def metric(rows: Rows, spec: dict):
    """Single number for KPI cards: {source, column, agg, where, ...}.

    `agg` is the final aggregation; a `group_by` step uses `group_agg` (default sum),
    so e.g. `group_by: quarter, agg: growth` means quarterly totals, last vs first.
    """
    rows = query(rows, {**spec, "agg": spec.get("group_agg", "sum")})
    col = spec["column"]
    check_columns(rows, [col], spec["source"])
    values = [r[col] for r in rows]
    how = spec.get("agg", "sum")
    if how == "last":
        return values[-1]
    if how == "first":
        return values[0]
    if how == "growth":  # last vs first, as a fraction
        return (values[-1] - values[0]) / values[0]
    return _aggregate(values, how)
