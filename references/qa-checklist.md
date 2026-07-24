# QA Checklist

Run this checklist before calling a DOCX final.

## Source QA

- No unclosed fenced code blocks.
- No missing local image files.
- No empty headings.
- No accidental duplicated first-level chapters.
- Mermaid blocks rendered successfully or intentionally left as placeholders with a clear reason.
- `待补充`, `TODO`, and bracketed missing facts are listed in the preflight report.

## Layout QA

- Cover page has the correct project title.
- Table of contents is present. If it is a Word field, remind the user to update fields after opening when needed.
- Heading hierarchy is consistent.
- First-level chapters start cleanly.
- Body text is readable and not over-compressed.
- Tables have header shading, readable text, and adequate cell padding.
- Wide tables are split, reduced, or marked for manual landscape treatment.
- Diagrams are crisp and captions are visible.
- Flowchart text stays inside boxes; nodes and connectors do not overlap. Dense Mermaid diagrams have been split, widened, or rerendered when needed.
- Images do not exceed page width.
- Authorization-letter ID copies stay with the authorization letter and are not promoted to standalone chapter pages.
- Contract titles appear with the first contract scan, not on a separate explanatory page.
- The final DOCX contains no production notes such as `可用扫描件`, `本合同附件共收录`, `以下每张扫描件独立成页`, image-correction explanations, or QA caveats.
- AAA credit and award-proof sections begin with concise claim/summary pages before evidence images.
- Header/footer and page numbers are present.
- Rendered PDF has no blank or nearly blank pages. If `pdf-pages-fullscan/` exists, sparse page count must be zero.
- Rendered PDF has no content touching page edges. If a visual scan report exists, edge-risk page count must be zero.
- Wide Mermaid diagrams are either placed on landscape pages or split into readable sub-diagrams.

## Final Evidence

Acceptable evidence includes:

- `preflight-report.json` with no fatal issues.
- `diagram-manifest.json` with all required diagrams rendered.
- final DOCX exists and has nonzero size.
- `docx-audit.json` has no fatal issues and no Mermaid leftovers.
- rendered page PNGs or PDF inspection when the renderer is available.
- `qa-report.json` / `qa-report.md` shows `pdf_sparse_page_review` passed when PDF rendering is available.
- full-page visual scan evidence, when generated, shows zero sparse pages and zero edge-risk pages.
- if rendering is unavailable, a structural audit plus explicit disclosure.
