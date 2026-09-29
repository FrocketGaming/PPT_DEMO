# deckgen — config-driven PowerPoint

Generic code builds the deck; a YAML file decides what goes on each slide and where.

```bash
uv sync
uv run deckgen configs/demo.yaml            # -> out/demo.pptx
uv run deckgen configs/demo.yaml -o my.pptx
uv run deckgen configs/demo_branded.yaml    # built on a brand template (see below)
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
  - type: title                # title | section | content (default) | closing
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

## Brand templates

Put a `.pptx` or `.potx` in `templates/` and name it in the config:

```yaml
deck:
  template: Brand Template        # -> templates/Brand Template.pptx (or .potx)
```

The deck is then built on that file's layouts, so logos, backgrounds, footer bars,
fonts and colours come from the template. The template's sample slides are removed.
Chart and table colours are read from the template's theme, and `deck.theme` can
override them. Regions (`left`, `right`, ...) are laid out inside the template's
content area. See `configs/demo_branded.yaml`, which has the same slides as `demo.yaml`
plus the template line.

**Profile.** `templates/<name>.yaml`, next to the template, says which layout each slide
type uses and which placeholder is the title, subtitle, date or tag. It also sets the
content area. Without a profile, deckgen guesses from layout names such as "Title Only".
To see a template's layouts and placeholder numbers when writing a profile:

```bash
uv run deckgen layouts "Brand Template"
```

**Per-slide layouts.** Any slide can use another layout and fill its placeholders by
number. Text keeps the template's fonts and bullet styles. A component is drawn in the
placeholder's position:

```yaml
- title: Revenue beat target
  layout: Diagram/Chart with Text
  placeholders:
    25: "Revenue beat target in **11 of 12** months"
    26: [First point, Second point]
    27: {type: chart, kind: column, data: monthly, x: month, y: revenue}
```

Slide types with a template: `title`, `section`, `content`, `closing`. On content
slides, the `tag` role defaults to the most recent section title.

### Testing against a template

Templates are git-ignored, so put your own in `templates/` first. `demo_branded.yaml`
expects `templates/Brand Template.pptx`; if your file has another name, change its
`template:` line.

```bash
uv run deckgen configs/demo_branded.yaml
start out/demo_branded.pptx          # Windows; `open` on macOS
```

Close the deck in PowerPoint before rebuilding, because Windows locks open files.

Check in PowerPoint:

- Logo, footer bar, backgrounds and fonts come from the template.
- Titles use the template's title style, not the built-in look.
- Slide numbers appear in the footer.
- Charts, KPI cards and tables sit between the title and the footer, with no overlap.
- Chart colours follow the brand palette.
- Right-click a chart → Edit Data opens the numbers.

If something is off, adjust the profile (`templates/<name>.yaml`) and rebuild:

| Problem | Fix |
|---|---|
| Content too close to the title or footer | Change `content_area: [x, y, w, h]` (inches) |
| Want a different title or section design | Change the `layout:` names for that slide type |
| Chart colours wrong or in the wrong order | Set `theme: palette: [...]` |
| Text lands in the wrong box | Check placeholder numbers with `deckgen layouts "<name>"` |

To check that auto-detection works on a new template, build against it without a profile.

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
- `template.py`: finds and opens brand templates (.pptx/.potx), reads their theme,
  maps layouts and placeholders, and removes sample slides
- `theme.py`: colors, fonts and the chart palette
