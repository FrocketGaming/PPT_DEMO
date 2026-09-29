# deckgen — config-driven PowerPoint

Generic code builds the deck; a YAML file decides what goes on each slide and where.

```bash
uv sync
uv run deckgen configs/demo.yaml            # -> out/demo.pptx
uv run deckgen configs/demo.yaml -o my.pptx
uv run deckgen configs/demo_branded.yaml    # built on a brand template (see below)
uv run deckgen layouts "Brand Template"     # list a template's layouts and placeholders
```

Charts are native PowerPoint charts, so they stay editable (right-click → Edit Data).
The YAML is the source of truth. Each build writes a fresh `.pptx`, so edits made in
PowerPoint are overwritten on the next build.

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

Every option is listed in the [Reference](#reference) section below.

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

## Brand templates

Put a `.pptx` or `.potx` in `templates/` and name it in the config:

```yaml
deck:
  template: Brand Template        # -> templates/Brand Template.pptx (or .potx)
```

`deck.template` is what makes the deck branded. The deck is built on that file's layouts,
so logos, backgrounds, footer bars, fonts and colours come from the template. The
template's sample slides are removed. Chart and table colours are read from the
template's theme, and `deck.theme` can override them. Regions (`left`, `right`, ...) are
laid out inside the template's content area. `configs/demo_branded.yaml` has the same
slides as `demo.yaml`, plus the template line.

**Profile.** `templates/<name>.yaml`, next to the template, sets the default layout for
each slide type and which placeholder holds the title, subtitle, date or tag. It also sets
the content area (see [Template profile](#template-profile)). Without a profile, deckgen
guesses from layout names such as "Title Only".

### Finding layouts

```bash
uv run deckgen layouts "Brand Template"
```

This lists every layout and the numbered placeholders (boxes) on each one:

```
Diagram/Chart with Text
  idx  22  body       at (0.44, 0.35) 10.49 x 1.08  Text Placeholder 7
  idx  27  chart      at (5.09, 1.29) 7.81 x 5.39  Chart Placeholder 6
  ...
```

The first line of each block is the name to use in `layout:`. Matching ignores case,
extra spaces and `–` vs `-`. `idx` is the number to use in `placeholders:`, and
`(x, y) w x h` is the box's position and size in inches. Each layout has its own numbers.
To see what the layouts look like, open the template in PowerPoint and check
**Home → Layout**, or **View → Slide Master** for full-size layouts. A mistyped layout
name fails the build and lists every valid name.

### Per-slide layouts

A slide can use any layout in the template and fill its placeholders by number:

```yaml
- title: Revenue beat target
  layout: Diagram/Chart with Text
  placeholders:
    25: "Revenue beat target in **11 of 12** months"
    26: [First point, Second point]
    27: {type: chart, kind: column, data: monthly, x: month, y: revenue}
```

Plain text keeps the template's styling. To override only some of it, use the styled form:

```yaml
    25: {text: "Revenue beat target in **11 of 12** months", font_size: 20}
    26:
      text: [First point, Second point]
      font_size: 14         # also: bold, italic, color (theme name or hex), align
```

Placeholders you leave empty are removed, so no "Click to add text" prompts are left
behind. Use regions when you want deckgen to handle placement. Use placeholders when you
want the template designer's exact layout.

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

## Reference

### Commands

| Command | What it does |
|---|---|
| `deckgen <config.yaml>` | Build the deck to `./out/<config name>.pptx` |
| `deckgen <config.yaml> -o <file.pptx>` | Build to a specific file |
| `deckgen layouts "<template>"` | List a template's layouts and placeholder numbers. Takes a name in `./templates/` or a path |

Run them with `uv run` in front, or activate the virtual environment first. A config
mistake stops the build with a message naming the slide and the problem.

### `deck`

| Key | Default | Meaning |
|---|---|---|
| `title`, `author` | `""` | Written to the file's document properties |
| `template` | none | Brand template name (file in `templates/`, extension optional) |
| `templates_dir` | none | Extra folder to search for templates, relative to the config |
| `date` | current month and year | Date shown on title slides (templates with a `date` role) |
| `theme` | built-in, or read from the template | Colour and font overrides, see below |

Templates are searched for in `templates_dir`, then `templates/` next to the config, then
`templates/` one level up, then `./templates`.

`theme` keys: `font`, `primary` (titles, table headers), `accent` (highlights, `**bold**`),
`text`, `muted` (labels, footnotes), `light` (card and callout backgrounds), `positive`
and `negative` (KPI deltas), and `palette` (list of chart series colours). Colours are hex,
with or without `#`.

### `data`

```yaml
data:
  monthly: ../data/monthly_sales.csv      # CSV path, relative to the config
  targets:                                # or inline rows
    - {region: North, target: 400000}
    - {region: South, target: 250000}
```

Numbers in CSVs are detected automatically.

### `defaults`

Options applied to every component of a type, e.g. `defaults: {chart: {font_size: 11},
table: {font_size: 11}}`. Anything set on the component itself wins.

### Slides

| Key | Slide types | Meaning |
|---|---|---|
| `type` | all | `title`, `section`, `content` (default), `closing` |
| `title`, `subtitle` | all | Slide title and optional subtitle |
| `left` `center` `right` `top` `bottom` `full` | content | Regions, each holding one component or a `stack` |
| `regions` | content | Same as the region keys, grouped under one key |
| `gap` | content | Space between regions in inches (default `0.3`) |
| `source` | content | Footnote, shown as "Source: ..." |
| `notes` | all | Speaker notes |
| `skip` | all | `true` leaves the slide out of the build |
| `font_size` | title, section, closing | Title size (built-in look only; templates use their own) |
| `layout` | all (template only) | Use this template layout instead of the profile's default |
| `placeholders` | all (template only) | Fill the layout's numbered boxes, see below |
| `tag` | content (template only) | Small label, if the layout has one. Defaults to the latest section title |
| `date` | title (template only) | Overrides `deck.date` for this slide |

A region is either a component (`left: {type: chart, ...}`) or a stack:

```yaml
left:
  size: 1                 # region size (see Regions)
  stack:
    - {type: kpis, weight: 1, items: [...]}       # weight = share of the region's height
    - {type: insights, weight: 2, items: [...]}
```

### Components

**`insights`**: bullet points.

| Key | Default | Meaning |
|---|---|---|
| `items` | | List of strings. `**text**` is bold in the accent colour |
| `heading` | none | Bold line above the items |
| `style` | `bullets` | `bullets`, `numbered`, or `callout` (shaded box with an accent bar) |
| `font_size` | `16` | Item size in points; the heading is 2pt larger |
| `spacing` | `10` | Space after each item, in points |

**`text`**: free text.

| Key | Default | Meaning |
|---|---|---|
| `text` | | The text. A blank line starts a new paragraph |
| `align` | `left` | `left`, `center`, `right` |
| `valign` | `top` | `top`, `middle`, `bottom` |
| `font_size` | `16` | Points |
| `bold` | `false` | |
| `color` | `text` | Theme colour name or hex |

**`kpis`**: a row of metric cards.

| Key | Default | Meaning |
|---|---|---|
| `items` | | One card per item, see below |
| `value_size` | `30` | Size of the big number on every card |

Each item takes:

| Key | Meaning |
|---|---|
| `label` | Small text above the value |
| `value` | A literal (`187.4`, `"Holiday Blitz"`) or a computed metric (below) |
| `format` | Python format for the value, e.g. `"${:,.0f}"`, `"{:.0%}"`, `"{:,}"` |
| `delta` | Change line under the value, literal or computed. A leading `-` shows ▼, otherwise ▲ |
| `delta_format` | Python format for the delta |
| `delta_label` | Text after the delta, e.g. `vs. plan` |
| `higher_is_better` | Default `true`. Set `false` to show increases in the `negative` colour |
| `value_size` | Overrides the row's `value_size` for this card |

A computed metric is `{source, column, agg}`, and it also accepts the
[query keys](#data-queries). `agg` is one of `sum` (default), `mean`, `count`, `min`,
`max`, `first`, `last`, or `growth` (last vs first, as a fraction).

```yaml
value: {source: monthly, column: revenue, agg: sum}
```

**`chart`**: a native, editable chart.

| Key | Default | Meaning |
|---|---|---|
| `kind` | `column` | `column`, `stacked_column`, `bar`, `stacked_bar`, `line`, `area`, `pie`, `doughnut`, `scatter` |
| `data` | | Source name, or inline `{categories: [...], series: {Name: [...]}}` |
| `x` | | Category column (the x values for `scatter`) |
| `y` | | One value column or a list; each becomes a series |
| `series_by` | none | Split one `y` column into a series per value of this column (long → wide). For `scatter`, it groups points into coloured series |
| `series_names` | column names | Rename series: `{revenue: "Revenue ($)"}` |
| `title` | none | Chart title |
| `legend` | `bottom` for multi-series, pie and doughnut; else `none` | `bottom`, `top`, `left`, `right`, `none` |
| `data_labels` | `false` | Show values on the chart |
| `number_format` | `General` | Excel number format for the data, e.g. `$#,##0`, `0%`, `$#,##0,"K"` |
| `axis_format` | `number_format` | Excel format for the value axis |
| `label_format` | `number_format` | Excel format for the data labels |
| `show_percentage` | `false` | Pie and doughnut labels show percentages |
| `gridlines` | `true` | Horizontal gridlines |
| `gap_width` | `60` | Space between bars, as a % of bar width |
| `font_size` | `11` | Points |

Charts also accept the [query keys](#data-queries).

**`table`**

| Key | Default | Meaning |
|---|---|---|
| `data` | | Source name |
| `columns` | all | List of columns, or `{column: Header}` to rename them |
| `formats` | none | Python format per column: `{revenue: "${:,.0f}"}` |
| `highlight_top` | `0` | Make the first N rows bold |
| `font_size` | `12` | Points |

Tables also accept the [query keys](#data-queries). Number columns are right-aligned.

**`image`**

| Key | Meaning |
|---|---|
| `path` | Image file, relative to the config. It fills the region's height and keeps its proportions |

### Data queries

Charts, tables and computed KPI values accept these keys. They are applied in this order:

| Key | Example | Meaning |
|---|---|---|
| `where` | `{region: West}` or `{region: [West, North]}` | Keep matching rows |
| `group_by` | `region` or `[region, quarter]` | One row per group |
| `agg` | `sum` | How `group_by` combines number columns: `sum`, `mean`, `count`, `min`, `max` |
| `sort_by` | `revenue` | Sort rows |
| `descending` | `true` | Sort largest first |
| `limit` | `5` | Keep the first N rows |

### Placeholders

`placeholders:` maps a placeholder number (from `deckgen layouts`) to one of:

| Value | Result |
|---|---|
| `"text"` | Text with the template's styling. `\n` starts a new line |
| `[line, line]` | One paragraph per item, using the template's bullet style |
| `{text: ..., font_size: 14}` | Styled text. Also `bold`, `italic`, `color`, `align` |
| `{type: chart, ...}` (any component) | The component, drawn where the placeholder sits |
| `{type: image, path: ...}` in a picture placeholder | The image, cropped to fill the placeholder |

### Template profile

`templates/<name>.yaml`:

```yaml
slides:                                  # one entry per slide type: title, section, content, closing
  title:
    layout: Presentation Title Cool Gray # layout name from `deckgen layouts`
    title: 10                            # placeholder number for each role
    subtitle: 11
    date: 12
  content:
    layout: Content Slide – Title Only
    title: 22
    tag: 21
content_area: [0.44, 1.6, 12.47, 5.1]   # x, y, width, height in inches for regions
theme:                                   # optional overrides of colours read from the template
  palette: ["1F4456", "E8542C", "448790"]
```

Roles are `title`, `subtitle`, `date` and `tag`. Without a profile, layouts are guessed
from their names ("Title Slide", "Section Header", "Title Only", "Thank You"). The title
is the template's title placeholder, or else the topmost wide text box. The
content area is the space between the title and any footer bar.

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
- `cli.py`: the `deckgen` command
