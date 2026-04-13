---
name: brain-to-deck
description: "One-click transformation of raw thoughts, uploads, and brain dumps into polished visual deliverables (HTML, PDF, PPTX, images). Use this skill whenever a user wants to turn messy inputs — text notes, images, docx, PDFs, scattered ideas, brainstorms, meeting notes, or any combination — into a cohesive presentation or visual material. Trigger on: 'organize my thoughts', 'turn this into a presentation', 'make this presentable', 'visualize my ideas', 'structure this', 'create a deck from these', 'one-pager', 'summarize and visualize', mixed file uploads needing unified output, or raw unstructured text that needs to become something showable. Also trigger for '展示材料', '整合', '结构化', '可视化'. The bridge between 'stuff in my head' and 'something I can show people'."
---

# Brain-to-Deck: Raw Thoughts → Polished Deliverables

Turn any combination of messy inputs into structured, beautiful output materials. This skill is the bridge between chaotic brain dumps and presentation-ready deliverables.

## Philosophy

People's ideas don't arrive in neat outlines. They come as scattered notes, half-formed thoughts, random screenshots, meeting recordings, PDFs from research, docs from collaborators — a mess of signal mixed with noise. This skill's job is to **extract the signal, find the structure, and produce something beautiful** — all in one shot.

## Workflow Overview

```
[Raw Inputs] → [Ingest & Parse] → [Analyze & Structure] → [Design & Render] → [Output]
```

### Step 1: Ingest Everything

Read ALL uploaded files and user-provided text. Use the right tool for each:

| Input Type | How to Read |
|-----------|-------------|
| Plain text / markdown in chat | Already in context — just use it |
| `.docx` | Read the docx skill: `/mnt/skills/public/docx/SKILL.md` |
| `.pdf` | Read the pdf-reading skill: `/mnt/skills/public/pdf-reading/SKILL.md` |
| `.pptx` | `python -m markitdown file.pptx` (install markitdown if needed) |
| `.xlsx` / `.csv` | Read with pandas or openpyxl |
| Images (`.png`, `.jpg`, etc.) | View them directly — describe what you see, extract any text |
| `.txt` / `.md` | `cat` or `view` tool |

Check `/mnt/user-data/uploads/` for all uploaded files. Don't skip anything — every piece of input matters.

### Step 2: Analyze & Structure

This is the most critical intellectual step. You are not just reformatting — you are **thinking**.

1. **Extract key themes**: What are the 3-7 core ideas across all inputs?
2. **Find the narrative**: What story connects these ideas? What's the logical flow?
3. **Identify hierarchy**: What's the big picture vs. supporting details?
4. **Spot data**: Are there numbers, comparisons, timelines, or relationships that deserve visualization?
5. **Catalog visuals**: Which uploaded images should be included? What new visuals should be created?

Present your proposed structure to the user before rendering, unless they've said "just do it" or the intent is obvious. A quick outline like:

```
Here's how I'd structure this:
1. [Title/Hook] — ...
2. [Core Problem/Context] — ...
3. [Key Insight A] — ...
4. [Key Insight B] — ...
5. [Data/Evidence] — ...
6. [Conclusion/Call to Action] — ...

Should I proceed, or would you like to adjust?
```

### Step 3: Choose Output Format

If the user specified a format, use it. Otherwise, recommend based on context:

| Use Case | Recommended Format | Why |
|----------|-------------------|-----|
| Sharing in meetings, presenting to team | **PPTX** | Standard presentation format, editable |
| Sharing via link, interactive content | **HTML** | Rich interactivity, works everywhere |
| Printing, formal documents, reports | **PDF** | Print-ready, professional |
| Social media, quick visual summary | **Image (HTML→screenshot)** | Single visual, easy to share |
| User says "展示材料" / "展示" without specifics | **HTML** first (richest output) | Most visual impact |

### Step 4: Design & Render

Read the appropriate downstream skill BEFORE generating:

- **PPTX** → Read `/mnt/skills/public/pptx/SKILL.md` (and its sub-references like `pptxgenjs.md`)
- **PDF** → Read `/mnt/skills/public/pdf/SKILL.md`
- **HTML** → Read `/mnt/skills/public/frontend-design/SKILL.md`
- **DOCX** → Read `/mnt/skills/public/docx/SKILL.md`

Then **ALWAYS read the design reference file** before generating ANY output:

→ **Read `references/bain-style.md`** (relative to this skill's directory)

This file defines the complete visual system: colors, typography, slide anatomy, chart formatting, layout patterns, content writing rules, and format-specific implementation notes. It is the single source of truth for all visual decisions. Do not deviate from it.

Key principles from the design reference (read the full file for specifics):

- **Action titles, not topic titles** — every slide title is a complete sentence stating the conclusion
- **One message per slide/section** — if you have two insights, make two slides
- **Pyramid Principle** — lead with the answer, then supporting arguments, then evidence
- **Bain Red (`CC2229`) as accent only** — white background, grayscale dominant, red to spotlight
- **Arial everywhere** — no decorative fonts, no emoji, no icons
- **Charts: column/bar, waterfall, line** — never pie charts, never 3D, never donut
- **Callout boxes on charts** — annotate the key insight directly on the chart
- **Source lines on every data slide** — bottom-left, 7-8pt gray
- **No decoration** — no gradients, shadows, glow, rounded corners, or animated transitions

### Step 5: Quality Check

After generating:

1. **Content check**: Does the output capture ALL the key ideas from the inputs? Did anything get lost?
2. **Flow check**: Does the narrative make sense from start to finish?
3. **Visual check**: Convert to images and inspect. Fix any overlaps, cut-off text, or alignment issues.
4. **Fidelity check**: Does it sound like the user, or did it drift into generic corporate speak?

## Special Handling

### Mixed-Language Content
If inputs contain multiple languages, produce the output in the dominant language unless the user specifies otherwise. Preserve key terms in their original language when they're domain-specific.

### Data-Heavy Inputs
When inputs contain lots of numbers/data (spreadsheets, reports with stats):
- Create dedicated data visualization slides/sections
- Use chart types appropriate to the data (comparison → bar chart, trend → line chart, composition → pie/donut, relationship → scatter)
- Always include the actual numbers alongside visualizations

### Image-Heavy Inputs
When the user uploads many images:
- Create a visual-forward layout (gallery, grid, or full-bleed images)
- Use images as section backgrounds or hero elements
- Add captions extracted from context

### Minimal Input ("Just some ideas")
When the user gives you very little to work with:
- Ask 1-2 clarifying questions about audience and purpose
- Expand on their ideas with reasonable assumptions
- Note your assumptions in the output so they can correct

## Output Delivery

Always save final output to `/mnt/user-data/outputs/` and present via `present_files`.

If you generated an HTML file, also mention they can open it in a browser for the full interactive experience.

If the user might want multiple formats, offer: "I've created this as [format]. Want me to also export it as [other format]?"

## Example Scenarios

**Scenario 1**: User pastes meeting notes + uploads a spreadsheet
→ Extract action items and data → Create a PPTX deck with meeting summary, key decisions, data charts, and next steps.

**Scenario 2**: User uploads 5 research PDFs + types "make me a literature review presentation"
→ Read all PDFs → Identify themes, agreements, contradictions → Create an HTML presentation with sections per theme, citation cards, and a synthesis conclusion.

**Scenario 3**: User uploads photos from a trip + types "make a travel memory book"
→ View all photos → Organize chronologically or thematically → Create a beautiful HTML page with photo gallery, captions, and story flow.

**Scenario 4**: User types a stream-of-consciousness brain dump about a startup idea
→ Extract the core value proposition, target market, competitive advantage → Create a one-pager PDF or pitch deck PPTX.
