# deckgen — config-driven PowerPoint

Generic code builds the deck; a YAML file decides what goes on each slide and where.

```bash
uv sync
uv run deckgen configs/demo.yaml            # -> out/demo.pptx
uv run deckgen configs/demo.yaml -o my.pptx
```

Charts are native PowerPoint charts, so they stay editable (right-click → Edit Data).

## Config shape

```yaml
deck:
  title: ...
  theme: {primary: "#1F3A5F", accent: "#E07A5F", font: Calibri}   # optional

data:                          # named sources, paths relative to the YAML
  monthly: ../data/monthly_sales.csv

defaults:                      # optional per-component-type defaults
  chart: {font_size: 11}

slides:
  - type: title                # title | section | content (default)
    title: ...
    subtitle: ...

  - title: Revenue beat target
    subtitle: optional line under the title
    source: footnote text
    notes: speaker notes
    left:  {type: insights, items: [...]}
    right: {type: chart, kind: line, data: monthly, x: month, y: [revenue, target]}
```

## Regions (where things go)

```
+-----------------------------+
|            top              |   size = fraction of height (default 0.3)
+---------+---------+---------+
|  left   | center  |  right  |   size = relative width (default 1)
+---------+---------+---------+
|           bottom            |
+-----------------------------+
```

Use any combination, or `full` alone. `left: {size: 1}`, `right: {size: 2}` gives a 1/3–2/3 split.
To put several components in one region, use `stack: [...]` (each item can have a `weight`).

## Components (what goes there)

| type       | key options |
|------------|-------------|
| `insights` | `items`, `heading`, `style: bullets \| numbered \| callout`, `font_size`. `**text**` is bold and in the accent color |
| `text`     | `text`, `align`, `valign`, `font_size`, `bold`, `color` |
| `kpis`     | `items: [{label, value, format, delta, delta_label, higher_is_better}]`. `value`/`delta` can be computed: `{source, column, agg: sum\|mean\|min\|max\|count\|first\|last\|growth}` |
| `chart`    | `kind: column \| stacked_column \| bar \| stacked_bar \| line \| area \| pie \| doughnut \| scatter`, `data`, `x`, `y` (one or a list), `series_by` (long → wide pivot), `title`, `legend`, `data_labels`, `number_format`, `axis_format`, `gridlines` |
| `table`    | `data`, `columns` (list or `{col: Header}`), `formats: {col: "${:,.0f}"}`, `highlight_top` |
| `image`    | `path` |

`chart` and `table` also accept the query keys `where`, `group_by` + `agg`, `sort_by` + `descending`, and `limit`.
`data:` can be inline instead of a source name: `{categories: [...], series: {Name: [...]}}`.

## Extending

Add a component type in `src/deckgen/components.py`:

```python
@register("quote")
def render_quote(slide, box, spec, ctx):
    ...
```

Then use `type: quote` in YAML. Layout, theming, data loading and validation come for free.

## Layout of the code

- `layout.py`: region names → rectangles
- `data.py`: CSV loading, where/group/sort/limit, pivot, metrics
- `components.py`: the renderers, plus the registry
- `builder.py`: reads the YAML, draws titles/footers, dispatches regions, validates
- `theme.py`: colors, fonts and the chart palette
