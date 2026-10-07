---
name: origin-word-figure-layout
description: Preflight and create compact, publication-ready multi-panel figures in Microsoft Word from Origin-exported EMF, PNG, JPEG, or TIFF panels. Use when Codex must check source-panel consistency (subfigure labels, axis titles, tick labels, legends, fonts, and line widths), crop white margins, preserve built-in a/b/c labels, align plot frames, keep one common scale, minimize gaps without content collisions, follow a manuscript/sample DOCX layout, or deliver and visually verify a DOCX figure sheet.
---

# Origin Word Figure Layout

Build the figure from the exported panels themselves. Preserve vectors when the inputs are EMF, and treat every visible mark—including built-in `(a)`, `(b)`, `(c)` labels—as protected content.

## Required companion workflow

For an AI assistant with a separate document-authoring skill installed, also follow that skill. This repository does not bundle external document skills. A direct Python CLI user must still render and inspect the final DOCX; prefer Microsoft Word rendering when floating EMF anchors make LibreOffice unreliable.

## Workflow

1. Inspect the input folder, the manuscript, and any user-edited sample DOCX. Determine panel order, rows, columns, page size, text-block width, and whether labels are already embedded.
2. Never add, remove, or relocate embedded subfigure labels unless explicitly requested. Do not assume a fixed left crop.
3. Before any DOCX authoring, run `scripts/preflight_origin_panels.py` and open its contact sheet. Compare subfigure labels, axis titles, tick labels, legends, fonts, curve widths, border widths, units, and canvas/frame proportions.
4. If any source-level inconsistency is material, stop before composing. Name the exact panel and element, ask the user to correct the Origin export, and rerun preflight after replacement. Never hide a source mismatch with independent Word scaling.
5. Render analysis previews at a common source scale. For EMF, use `scripts/render_emf_previews.ps1`; the composer calls it automatically on Windows.
6. Detect each panel's complete nonwhite-content boundary and plot frame. Include separated or outward-projecting a/b/c labels in the complete boundary.
7. Align plot-frame left edges within columns and plot-frame top edges within rows. Use collision masks to find the smallest safe horizontal and vertical separations.
8. Scale all panels by one common factor. Never resize panels independently merely to equalize their outer boxes; that would change font and line-width consistency.
9. Create a new DOCX. Keep EMF inputs as EMF, use zero-wrap floating anchors, permit geometric image-box overlap, and prevent visible-content overlap.
10. Render the DOCX with Microsoft Word, inspect every panel at high resolution, and iterate until frames align and no label, axis title, tick text, legend, or curve is clipped or collides.
11. Preserve source exports, manuscript, and sample DOCX files. Write a clearly named new output.

## Mandatory source-panel preflight

Run this before the compose command:

```powershell
python scripts/preflight_origin_panels.py `
  --input-dir "C:\path\to\origin-exports" `
  --cols 2 `
  --files "panel-a.emf" "panel-b.emf" "panel-c.emf" "panel-d.emf" `
  --qa-dir "C:\path\to\figure-preflight"
```

Open `preflight-contact-sheet.png` at full resolution. Treat the preflight as a human visual gate, not as a fully automatic typography detector. Check:

- a/b/c/d font family, weight, height, and offset from the plot frame;
- x/y-axis-title font family, size, weight, capitalization, units, and symbol formatting;
- tick-label font size and weight;
- legend font, sample-line length/thickness, marker size, spacing, and placement;
- curve, marker, axis, and border line widths;
- matching canvas size and plot-frame proportions where the scientific panel design should match.

If uncertain, report the suspected mismatch and ask for a corrected export. Only proceed after the contact sheet is visually consistent. Read [references/layout-principles.md](references/layout-principles.md) for the pass/fail boundary.

## Compose command

Use the Python environment described in the repository README, then run:

```powershell
python scripts/compose_origin_word_figure.py `
  --input-dir "C:\path\to\origin-exports" `
  --output "C:\path\to\figure-compact.docx" `
  --cols 3 `
  --files "panel-a.emf" "panel-b.emf" "panel-c.emf" "panel-d.emf" "panel-e.emf" "panel-f.emf" `
  --max-width-in 5.75 `
  --safety-pt 1.0 `
  --qa-dir "C:\path\to\figure-qa"
```

Omit `--files` only when natural filename order is correct. The panel count must fill the rectangular grid. Use `--frame-overrides-json` only when automatic plot-frame detection reports uncertainty.

Inputs in one run should all be EMF or all be raster images. Prefer EMF. Export all panels from the same Origin canvas/page size; the analyzer normalizes them to one preview width so differing EMF internal pixel densities do not change font scale. Raster inputs must likewise share one export canvas and DPI.

## Spacing and a/b/c rules

- Define crop edges from the union of all nonwhite content, not from the plot frame or axis labels alone.
- Keep a small white padding outside the detected union.
- Let a label farther left than the y-axis title enlarge that panel's protected left extent.
- Let a label aligned with the y-axis title add no artificial margin.
- Allow cropped image rectangles to overlap in Word when their visible content does not overlap.
- Treat raster-image crop rectangles as opaque unless proven otherwise; do not overlap their rectangles even when their visible masks would fit. EMF panels may use negative rectangle spacing because Word preserves their transparent background.
- Default to `1.0 pt` visible-content clearance. If the user provides a manually adjusted sample, measure and match its apparent clearance while preserving collision safety.
- If a label touches the source canvas edge, do not guess that it is complete. Flag the source as already clipped and ask for a re-export with more canvas margin.

For the layout model and failure checks, read [references/layout-principles.md](references/layout-principles.md).

## Validation checklist

- Confirm the DOCX contains exactly one anchored image per panel and no unintended inline duplicates.
- Confirm the preflight contact sheet was inspected before authoring and no unresolved source-style mismatch remains.
- Confirm vector inputs remain `.emf` inside `word/media`.
- Confirm all panel extents share one scale factor.
- Confirm plot frames form straight rows and columns.
- Inspect outward a/b/c labels, y-axis titles, x-axis titles, legends, and edge tick labels at 200–300% zoom.
- Confirm the source and reference hashes are unchanged.
- Deliver the DOCX and a rendered preview.
