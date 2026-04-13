# Bain-Style Visual Design Reference

This document defines the visual design system used by the brain-to-deck skill. All outputs — PPTX, HTML, PDF — must follow these specifications. The style is modeled after the presentation conventions of top-tier management consulting firms, particularly Bain & Company: clean, structured, message-driven, and free of decoration.

## Core Philosophy

**"The plainer it looks, the more attention goes to the content."**

Every visual choice serves the argument. Nothing is decorative. If a design element doesn't clarify the message, it doesn't belong. The goal is not to impress with aesthetics but to move decision-makers from confusion to clarity in the shortest possible time.

---

## Color Palette

Use exactly these colors. Do not invent new ones or add gradients.

| Role | Name | Hex (no #) | Usage |
|------|------|-----------|-------|
| **Primary** | Bain Red | `CC2229` | Accent bar, key highlights, chart emphasis, section dividers |
| **Text Primary** | Near Black | `1A1A1A` | All body text, titles, labels |
| **Text Secondary** | Dark Gray | `4A4A4A` | Subtitles, supporting text, axis labels |
| **Text Tertiary** | Medium Gray | `7F7F7F` | Source lines, footnotes, captions |
| **Background** | White | `FFFFFF` | Default slide/page background |
| **Background Alt** | Light Gray | `F2F2F2` | Alternate section backgrounds, table stripes |
| **Border / Grid** | Subtle Gray | `D9D9D9` | Table borders, chart grid lines, divider lines |
| **Chart Color 1** | Bain Red | `CC2229` | Primary data series, "our" data, focus item |
| **Chart Color 2** | Charcoal | `3D3D3D` | Secondary data series, comparison |
| **Chart Color 3** | Steel Gray | `8C8C8C` | Tertiary data series |
| **Chart Color 4** | Light Steel | `BFBFBF` | Quaternary data series, "other" category |
| **Positive** | Muted Green | `3D8B37` | Positive change indicators only (use sparingly) |
| **Negative** | Bain Red | `CC2229` | Negative change indicators |

**Rules:**
- White background is the default. Never use dark backgrounds unless quoting the user's explicit preference.
- Red is the accent, not the dominant color. Use it for one element per slide/section to draw focus.
- Charts use a grayscale-dominant palette with red to spotlight the key insight.
- Never use gradients, glow effects, drop shadows, or transparency.

---

## Typography

### For PPTX (PowerPoint)
| Element | Font | Size | Weight | Color |
|---------|------|------|--------|-------|
| **Action Title** | Arial | 18-20pt | Bold | `1A1A1A` |
| **Subtitle / Context** | Arial | 12-14pt | Regular | `4A4A4A` |
| **Body Text** | Arial | 11-12pt | Regular | `1A1A1A` |
| **Bullet Points** | Arial | 11-12pt | Regular | `1A1A1A` |
| **Chart Labels** | Arial | 9-10pt | Regular | `4A4A4A` |
| **Data Callout** | Arial | 24-36pt | Bold | `CC2229` |
| **Source Line** | Arial | 7-8pt | Regular | `7F7F7F` |
| **Page Number** | Arial | 8pt | Regular | `7F7F7F` |

### For HTML
| Element | Font Stack | Size | Weight | Color |
|---------|-----------|------|--------|-------|
| **Action Title** | `Arial, Helvetica, sans-serif` | 22-28px | 700 | `#1A1A1A` |
| **Section Heading** | `Arial, Helvetica, sans-serif` | 18-20px | 700 | `#1A1A1A` |
| **Body** | `Arial, Helvetica, sans-serif` | 14-15px | 400 | `#1A1A1A` |
| **Small / Caption** | `Arial, Helvetica, sans-serif` | 12px | 400 | `#7F7F7F` |
| **Data Callout** | `Georgia, 'Times New Roman', serif` | 36-48px | 700 | `#CC2229` |

### For PDF (reportlab)
| Element | Font | Size | Color |
|---------|------|------|-------|
| **Title** | Helvetica-Bold (or CJK Bold) | 18-20pt | `1A1A1A` |
| **Body** | Helvetica (or CJK Regular) | 10-11pt | `1A1A1A` |
| **Source** | Helvetica (or CJK Regular) | 7pt | `7F7F7F` |

**Rules:**
- Arial is the primary typeface across all formats. Never use decorative or unusual fonts.
- Georgia is used only for large data callout numbers to add subtle visual distinction.
- Bullet characters: use standard round bullet `•`, never emoji or custom symbols.
- Text alignment: left-align everything except data callout numbers (center-align).
- No italic text except for direct quotes. No underlines. Use bold sparingly — only for titles and key emphasis.
- CJK content: use Noto Sans CJK or the system's best available CJK sans-serif font.

---

## Slide / Page Structure (The Consulting Slide Anatomy)

Every slide or content section follows this exact anatomy, top to bottom:

```
┌──────────────────────────────────────────────┐
│ [Red accent bar: 3-4px, full width]          │
│                                              │
│ ACTION TITLE (1-2 lines, complete sentence)  │
│ Subtitle context (optional, 1 line)          │
│                                              │
│ ┌──────────────────────────────────────────┐ │
│ │                                          │ │
│ │         SINGLE EXHIBIT OR                │ │
│ │         CONTENT BLOCK                    │ │
│ │         (chart, table, or structured     │ │
│ │          text — NOT paragraphs)          │ │
│ │                                          │ │
│ └──────────────────────────────────────────┘ │
│                                              │
│ Source: [data source]           Page X of Y  │
└──────────────────────────────────────────────┘
```

**Key Rules:**
1. **Action Title, not topic title.** "Revenue declined 12% due to customer churn" — not "Revenue Analysis."
2. **One message per slide/section.** If you have two insights, make two slides.
3. **The title carries the message; the body provides evidence.** A reader should understand the argument by reading titles alone.
4. **Red accent bar** at the top of every slide — thin (3-4px), full-width. This is the only consistent use of red.
5. **Source line** at bottom-left for any data-driven content. Format: `Source: [description], [year]`
6. **Page numbers** at bottom-right.

---

## Charts & Data Visualization

### Preferred Chart Types (in order of preference)
1. **Column / Bar chart** — comparisons across categories
2. **Waterfall chart** — showing cumulative effect of sequential values
3. **Line chart** — trends over time
4. **Stacked bar (100%)** — composition / market share
5. **Simple table** — when precision matters more than shape
6. **Scatter plot** — relationships between two variables

### Chart Types to AVOID
- Pie charts (hard to read, low information density)
- 3D charts of any kind
- Donut charts
- Radar / spider charts
- Treemaps
- Infographic-style visuals

### Chart Formatting Rules
- **No chart borders or boxes** — charts float cleanly on the white background.
- **Gridlines**: only horizontal gridlines, light gray (`D9D9D9`), thin (0.5pt). No vertical gridlines.
- **Axis labels**: dark gray (`4A4A4A`), 9-10pt. Category axis at bottom, value axis at left.
- **Data labels**: show directly on bars/columns when possible, eliminating the need for a value axis.
- **Legend**: place below or to the right of the chart. Remove legend if only one data series.
- **Chart colors**: use the palette above. The focus item is red; everything else is gray.
- **Callout box**: use a small gray box with a bold insight sentence to annotate the key finding in a chart. This is a signature Bain technique — the chart shows the data, the callout tells you what to think about it.

### Data Callouts (Big Numbers)
When highlighting a key metric, display it as:
- **Large number** in 36-48pt Georgia Bold, Bain Red (`CC2229`)
- **Unit/label** directly below in 11pt Arial, dark gray
- **Context** (e.g., "vs. 3.2s last quarter") in 9pt, medium gray

---

## Layout Patterns

### Common Slide Layouts (PPTX)

**Layout 1: Chart + Bullets (most common)**
Left 60%: chart. Right 40%: 3-5 bullet points with supporting insights.

**Layout 2: Two-Column Comparison**
Left column and right column, each with a header. Used for before/after, option A vs. B, etc.

**Layout 3: Framework / Matrix**
2x2 or 3x3 grid. Clean borders, labels in each cell. Used for prioritization, categorization.

**Layout 4: Data Table**
Clean table with header row in Bain Red background + white text. Alternating row shading (white / `F2F2F2`). No excessive borders — only horizontal lines.

**Layout 5: Key Takeaway**
Large action title + 3 supporting data callouts arranged horizontally. Used for executive summary or key findings.

**Layout 6: Process / Timeline**
Horizontal arrow or numbered steps. No icons — just numbered circles and text labels.

### Spacing & Margins (PPTX, 10" x 5.625")
| Element | Value |
|---------|-------|
| Slide margins | 0.5" all sides |
| Between title and content | 0.3" |
| Between content blocks | 0.25-0.4" |
| Between bullet items | `paraSpaceAfter: 6` |
| Source line from bottom | 0.15" |

### HTML Layout
- Max content width: `960px`, centered
- Section padding: `48px 0`
- Card/block padding: `24px`
- Use CSS Grid or Flexbox for multi-column layouts
- No rounded corners exceeding `4px` (prefer `2px` or `0`)
- No box shadows. Use `1px solid #D9D9D9` borders instead.
- Background: `#FFFFFF` default. `#F2F2F2` for alternate sections.

---

## Tables

| Property | Specification |
|----------|--------------|
| Header row background | `CC2229` |
| Header row text | White, Bold, 11pt |
| Body row background | Alternating `FFFFFF` / `F2F2F2` |
| Body text | `1A1A1A`, Regular, 10-11pt |
| Borders | Horizontal only, `D9D9D9`, 0.5pt |
| Cell padding | 0.1" vertical, 0.15" horizontal |
| Alignment | Text left, numbers right |
| **Header bar text padding** | When placing text over a colored header bar (shape + text overlay), offset the text x by +0.2" from the shape's x to ensure visible left padding. Do NOT rely on margin alone — use coordinate offset. |

---

## Content Writing Rules

These rules govern how text content is written — they are just as important as visual formatting.

1. **Action titles must be complete sentences** that state the conclusion. "Revenue declined 12% QoQ driven by customer churn in the SMB segment" — not "Revenue Overview."
2. **Bullet points are sentence fragments**, not full paragraphs. Max 2 lines per bullet.
3. **No more than 5 bullets per slide.** If you need more, split the slide.
4. **Lead with the insight, not the data.** The takeaway comes first; the number supports it.
5. **Write like you talk.** No jargon soup. "Costs are rising faster than revenues" — not "The organization is experiencing unfavorable operational cost-revenue dynamics."
6. **Every number needs a source.** Use the source line at the bottom of the slide.
7. **Apply the Pyramid Principle**: answer first, then supporting arguments, then evidence.
8. **Apply MECE structure**: arguments should be mutually exclusive and collectively exhaustive.

---

## What to NEVER Do

- ❌ Decorative icons, emoji, or clip art
- ❌ Gradient backgrounds or fills
- ❌ Drop shadows or glow effects
- ❌ 3D effects on charts
- ❌ Rounded corners greater than 4px
- ❌ Multiple fonts on the same slide
- ❌ More than 2 colors per chart (exception: when comparing 3+ categories, use the gray scale)
- ❌ Centered body text (always left-align)
- ❌ Topic titles instead of action titles
- ❌ Dense paragraphs of text on slides
- ❌ Decorative accent lines under titles
- ❌ Pie charts or donut charts
- ❌ Animated transitions or builds

---

## Format-Specific Notes

### PPTX
- Layout: `LAYOUT_16x9` (10" x 5.625")
- Red accent bar: `addShape(RECTANGLE, { x:0, y:0, w:10, h:0.05, fill:{color:"CC2229"} })`
- Page number: bottom-right, 8pt, `7F7F7F`
- Source line: bottom-left, 7-8pt, `7F7F7F`
- No slide transitions

### HTML
- Include a thin `4px solid #CC2229` bar at the top of each section
- Use `<hr>` styled as `1px solid #D9D9D9` between sections
- Tables: use `border-collapse: collapse` with only horizontal borders
- Charts: use simple CSS or inline SVG. No heavy charting libraries unless data demands it.
- Print stylesheet: ensure clean A4/Letter printing

### PDF
- Red accent line at top of page: `c.setFillColor(HexColor("#CC2229")); c.rect(margin, H-margin, content_w, 3, fill=1)`
- Use Helvetica/Arial for Latin text, Noto Sans CJK for Chinese/Japanese/Korean
- Generous margins: 25-30mm
- Source lines in 7pt gray at bottom of content blocks
