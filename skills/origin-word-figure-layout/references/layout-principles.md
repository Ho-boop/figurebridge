# Layout principles

## Source-style preflight gate

Do not start Word composition until the normalized preflight contact sheet has been inspected. Grouping can correct crop, position, and spacing; it cannot safely correct inconsistent typography or line styling embedded inside exported panels.

Block composition and ask for a new Origin export when a difference is visible at the same normalized canvas scale in any of these roles: subfigure label, axis title, tick label, legend text, legend sample line/marker, curve width, marker size, plot border, unit notation, or font weight/family. Name the exact panel and role. A small difference caused only by different glyph shapes or different text strings is not a failure; compare cap height, stroke weight, and role-to-frame proportions.

Do not independently resize a mismatched panel to make its text look equal. Independent resizing also changes curves, borders, markers, and all other typography. Re-export the source instead.

## Protected content

Build a binary mask from every pixel that differs materially from white. The protected mask includes plot borders, curves, markers, text, legends, and subfigure labels. Compute the crop box from the union of this mask plus padding.

An embedded a/b/c label is ordinary protected content:

- If it lies left of the y-axis title, its left edge controls the crop and collision boundary.
- If it shares the y-axis-title column, it usually does not change the crop.
- If it is detached from the frame, keep it; do not discard small components merely because they are isolated.
- If any protected pixel touches the source edge, report possible source clipping.

## Alignment coordinate system

Detect the rectangular plot frame independently of the outer content box. Express each panel's mask coordinates relative to the frame's left and top edges. This allows different label extents while keeping the scientific plotting regions aligned.

For horizontal placement, align frame tops within a row and find the minimum frame-origin displacement for which mask envelopes do not intersect, then add the requested visible clearance. For vertical placement, align frame left edges within a column and solve the analogous constraint. Solve constraints for all panel pairs, not only adjacent pairs, so a detached label cannot cross a skipped whitespace band and collide with a nonadjacent panel.

## Common scale

Use one inches-per-preview-pixel factor for the full figure. Compute positions in preview pixels first, then scale the entire system to the manuscript text-block width. Independent panel scaling is forbidden unless the user explicitly confirms the Origin exports were created at different physical scales and requests normalization.

## Safety distance

Expand the protected-mask separation constraints by the requested clearance. A final clearance near 1 pt is usually sufficient. A user-edited reference DOCX may justify a smaller value, but never allow protected masks to overlap.

Raster exports commonly contain opaque white backgrounds. For raster panels, also enforce non-overlap of the cropped image rectangles; otherwise a later panel can paint white over an earlier panel's text. EMF panels can use mask-safe negative rectangle spacing when Word renders their background transparently.

## Automatic frame detection

Look for the outermost long dark horizontal and vertical strokes. Reject a detection when the inferred frame is implausibly small or only one edge is found. When rejected, render the preview and either provide a JSON frame override in preview-pixel coordinates or request a cleaner Origin export.

Frame override format:

```json
{
  "panel-a.emf": [left, top, right, bottom],
  "panel-b.emf": [left, top, right, bottom]
}
```

## Final visual checks

Inspect the Word-rendered output, not only a synthetic montage. Check:

1. all plot-frame tops and bottoms in each row;
2. all plot-frame left and right edges in each column;
3. a/b/c labels and y-axis titles at the left boundary;
4. x-axis titles against the next row;
5. legends and tick labels near neighboring panels;
6. source-edge clipping that layout code cannot repair.
