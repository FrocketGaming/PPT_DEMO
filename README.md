# deckgen — config-driven PowerPoint

Generic code builds the deck; a YAML file decides what goes on each slide and where.

```bash
uv sync
uv run deckgen configs/demo.yaml            # -> out/demo.pptx
uv run deckgen configs/demo.yaml -o my.pptx
uv run deckgen configs/demo_advanced.yaml   # a full business review: dashboards, scorecards, small multiples
uv run deckgen configs/recipes.yaml         # one slide per common use case, to copy from
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

Every option is listed in the full reference, [`docs/index.html`](docs/index.html).

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

## Building a good slide

deckgen handles layout and styling. What is left for you is the message. A slide that
works usually follows four rules:

1. **The title is the conclusion.** "Revenue beat target in 11 of 12 months" says
   something. "Monthly revenue" does not.
2. **One slide, one message.** If you need two headlines, make two slides.
3. **Show the point, grey out the rest.** Use `highlight` on a chart, `active` on a
   timeline, `panel` on the one thing that matters.
4. **Words explain, they don't repeat.** Insights say why the chart matters. Keep them to
   three short points.

`configs/recipes.yaml` builds one slide per pattern below. Copy the one closest to what you
need.

```
uv run deckgen configs/recipes.yaml    # -> out/recipes.pptx
```

### Choosing a layout

| You want to... | Use | Regions |
|---|---|---|
| Open with the answer | KPI row plus a callout | `top` (kpis) + `bottom` (insights `callout`) |
| Prove one point | One large chart | `full` |
| Show a trend and explain it | Chart plus short bullets | `left` (chart, `size: 2`) + `right` (insights) |
| Compare two things | Two matching charts, same axis | `left` + `right`, both with `axis_min` and `axis_max` |
| Compare three things | Three panels and a takeaway | `left` + `center` + `right` + `bottom` |
| Rank against a target | Progress bars | `full` (progress with `target`) |
| Explain a process or plan | Timeline plus the main risk | `top` (timeline) + `bottom` (callout) |
| Explain a repeating loop | Cycle plus rules | `left` (cycle) + `right` (numbered insights) |
| Report a scorecard | KPI row plus a table | `top` (kpis `dark`) + `bottom` (table) |
| Break the deck into chapters | Section slide | `type: section` with `number` and `background` |
| Land a human voice | Quote | `full` (quote) |

### Recipes

**Executive summary.** The first content slide: the answer, the proof, the ask.

```yaml
- title: Revenue beat plan, but Trade Show is dragging returns
  top:
    type: kpis
    size: 0.32
    style: accent
    items:
      - label: Revenue
        value: {source: monthly, column: revenue, agg: sum}
        format: "${:,.0f}"
        delta: "+4.2%"
        delta_label: vs. plan
  bottom:
    type: insights
    style: callout
    items:
      - "**Decision needed:** move budget from Trade Show into online retargeting in Q1."
```

**One chart, one message.** Grey out everything except the bar the title is about, and draw
the goal as a dashed line. The title should say what the highlighted bar shows.

```yaml
- title: December hit $276K, 38% above the $200K goal
  full:
    type: chart
    kind: column
    data: monthly
    x: month
    y: revenue
    highlight: Dec                                  # the bar the title is about
    reference: {value: 200000, label: $200K goal}
    number_format: '$#,##0,"K"'
    data_labels: true
    gridlines: false
```

**Chart plus interpretation.** Give the chart two thirds of the width. Use `direct_labels`
so the reader does not have to match a legend to a line.

```yaml
- title: Q4 accelerated sharply
  left:  {type: chart, kind: line, size: 2, data: monthly, x: month, y: [revenue, target],
          direct_labels: true}
  right: {type: insights, heading: Why it matters, panel: true,
          items: [December alone was **$276K**, Growth was **volume-led**]}
```

**Side-by-side comparison.** Give both charts the same `axis_min` and `axis_max`, or the
bars will not be comparable. Use `where` to slice one data source twice.

```yaml
- title: West outsells South in both products
  left:
    type: chart
    kind: bar
    title: Core
    data: regional
    where: {product: Core}
    group_by: region
    x: region
    y: revenue
    axis_min: 0
    axis_max: 600000
  right: {type: chart, kind: bar, title: Pro, data: regional, where: {product: Pro},
          group_by: region, x: region, y: revenue, axis_min: 0, axis_max: 600000}
```

**Ranking against a target.** `progress` bars read faster than a table when the question is
"who cleared the bar?".

```yaml
- title: Three of five campaigns cleared 4x return on spend
  full:
    type: progress
    max: 6
    items:
      - {label: Webinar Series, value: 5.39, format: "{:.1f}x", target: 4, highlight: true}
      - {label: Trade Show, value: 2.2, format: "{:.1f}x", target: 4}
```

**Plan or process.** Set `active` so the reader sees where you are. Earlier steps show as
done and later steps are greyed.

```yaml
- title: Rollout takes three quarters
  top:
    type: timeline
    size: 1.4
    active: 1
    items:
      - {label: Q1, title: Pilot, text: "Two regions, **10%** of spend"}
      - {label: Q2, title: Expand, text: All regions}
  bottom:
    type: insights
    style: callout
    items: ["**Risk:** the pilot regions are not representative of the South."]
```

**Chapter divider.** A dark background flips text to white and tints cards automatically.

```yaml
- type: section
  title: Customer support
  number: "02"
  background: primary
```

**Scorecard.** KPIs on top for the headline numbers, a sorted table below for the detail.
`highlight_top: 1` bolds the row to look at first.

```yaml
- title: Support quality is steady; North escalations need a look
  top:    {type: kpis, size: 0.3, style: dark, items: [...]}
  bottom: {type: table, data: support, group_by: region, agg: sum,
           columns: {region: Region, escalations: Escalations},
           sort_by: escalations, descending: true, highlight_top: 1}
```

### Habits worth copying

- **Put numbers in `data`, not in the slide.** KPI values, chart series and table rows
  computed from a source stay correct when the data changes. Type a literal only for
  something that is not in the data, such as a plan variance.
- **Wrap the key phrase in `**bold**`.** It renders in the accent colour, so a bullet's
  point survives a skim.
- **Use `defaults`** for font sizes instead of repeating them on every chart.
- **Set `source:` on any slide with data.** It adds a footnote and takes no space from the
  content.
- **Put the story in `notes:`.** Speaker notes are written to the file, so the config holds
  the script as well as the slides.
- **Use `stack` to fit a KPI above bullets** in one column, and `weight` to split the
  height.
- **Repeat the `agenda`** with a different `active` between sections, so the audience
  always knows where they are.

### Things to avoid

- More than about five bullets, or bullets longer than a line and a half. Split the slide.
- Pie or doughnut charts with more than five slices. Use a sorted `bar` instead.
- Two charts side by side with different axis ranges.
- A legend when `direct_labels` or `highlight` would do.
- Every KPI card with a `delta`. Keep deltas for the numbers that changed meaningfully.

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

**Colours are detected from the template.** deckgen reads the template's theme
(**Design → Variants → Colors** in PowerPoint) and maps it as follows:

| deckgen colour | Taken from the template's |
|---|---|
| `primary` (titles, table headers) | Dark 2, or Dark 1 if there is no Dark 2 |
| `accent` (highlights, `**bold**`) | Accent 1 |
| `text` | Dark 1 |
| `light` (cards, callouts) | Light 2 |
| `palette` (chart series, in order) | Dark 2, then Accent 1 to 6. Very light colours are skipped, because they vanish on a white chart |
| `font` | The theme's body font |

Not detected: `muted`, `subtle`, `positive` and `negative`. They keep the built-in greys,
teal and red. If those clash with the brand, set them in `deck.theme`.

Colours are applied in this order, and later ones win: the template's theme, then `theme:`
in the profile, then `deck.theme` in the config. If a chart uses the wrong colours, check
what is in the template's Colors scheme, then override just the `palette`.

**Profile.** `templates/<name>.yaml`, next to the template, sets the default layout for
each slide type and which placeholder holds the title, subtitle, date or tag. It also sets
the content area (see [Template profile](docs/index.html#profile)). Without a profile, deckgen
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

The full reference is in [`docs/index.html`](docs/index.html). Open it in a browser
(`start docs/index.html` on Windows, `open` on macOS). It is a single searchable page that
covers every option, its default and an example:

- **Formatting:** which options take Python formats (KPIs, progress bars, tables) and
  which take Excel number formats (charts), with cheat sheets for both
- **Data:** sources, queries (`where`, `group_by`, `sort_by`, ...) and computed metrics
  (`sum`, `mean`, `growth`, ...)
- **Components:** `insights`, `text`, `quote`, `kpis`, `progress`, `chart`, `table`,
  `agenda`, `timeline`, `cycle`, `image`
- **Deck and slides:** `deck`, `theme`, `defaults`, slide keys, regions and stacks
- **Brand templates:** placeholders and the template profile
- **Errors and fixes:** what each build error means

When you add or change an option, update `docs/index.html` in the same commit.

## Extending

Add a component type in `src/deckgen/components.py`:

```python
@register("quote")
def render_quote(slide, box, spec, ctx):
    ...
```

Then use `type: quote` in YAML. Layout, theming, data loading and validation come for free.
Add a section for the new component to `docs/index.html`.

## Layout of the code

- `layout.py`: region names → rectangles
- `data.py`: CSV loading, where/group/sort/limit, pivot, metrics
- `components.py`: the renderers, plus the registry
- `builder.py`: reads the YAML, draws titles/footers, dispatches regions, validates
- `template.py`: finds and opens brand templates (.pptx/.potx), reads their theme,
  maps layouts and placeholders, and removes sample slides
- `theme.py`: colors, fonts and the chart palette
- `cli.py`: the `deckgen` command
