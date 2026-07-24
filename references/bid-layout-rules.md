# Bid Layout Rules

Use these rules when a bid manuscript contains authorization letters, representative ID copies, contract scans, AAA credit/honor materials, award proof, Mermaid flowcharts, or dense organization charts.

## Never Expose Production Notes

Do not put process explanations in the final DOCX, including text like:

- `可用扫描件 N 张`
- `本合同附件共收录可用扫描件 N 张`
- `以下每张扫描件独立成页`
- `保持原件比例，不拼版、不裁切`
- image-mapping decisions, missing-page explanations, or QA caveats

Record such information only in sidecar files such as `图片映射清单.md`, `preflight-report.md`, or `pdf-qa-report.md`.

## Authorization Pages

- Treat `全权代表身份证复印件` as part of the authorization letter, not as a standalone chapter.
- Do not render `全权代表身份证复印件` as a large heading.
- Keep the representative ID image on the same page as the authorization letter whenever the source identifies it as an attachment to the letter.
- If space is tight, reduce vertical whitespace and image size before allowing a page split.

## Contract Scans

- A contract section title such as `罗西十二组团项目合同` must appear on the same page as the first contract scan, above the image.
- Do not create a separate contract cover/section page unless the source contains a real cover page image.
- Place each contract scan on its own page after the first page. Do not let the next contract title or next chapter appear below a contract scan.
- Use only concise in-document labels. If available scan count differs from a declared page count, record that in a sidecar mapping/QA file, not in the DOCX.
- Do not invent total page counts. When the scan itself contains page numbering, avoid conflicting external labels.

## AAA Credit And Honors

- `AAA级荣誉证书` should begin with a short summary page stating the claim, for example `连续三年获得 AAA 级信用等级证书`.
- Put supporting red-head notices, lists, and certificate images after the summary, starting on the next page.
- Do not let the first evidence image substitute for the summary.

## Award Certificates

- `投标人获奖证书` should begin with a summary page saying the bidder has municipal and provincial honors.
- Include a compact table on that summary page with columns such as `类别`, `荣誉/文件`, `年度`, `证明材料`.
- Start red-head documents and other proof images on the next page.
- Keep municipal and provincial evidence grouped and ordered.

## Mermaid And Flowcharts

- Rendered flowcharts must be legible at PDF size: text must stay inside boxes, boxes must not overlap, and connector labels must not collide with nodes.
- If a Mermaid diagram has dense Chinese labels, long node text, or overlapping boxes, do not accept the default render.
- Improve it by shortening labels, wrapping node text, increasing node spacing/rank spacing, switching to landscape, or splitting into multiple smaller diagrams.
- If the built-in Mermaid fallback cannot avoid overlap, use Mermaid CLI, draw.io, or a custom diagram rendering path before delivery.
- Visual QA must inspect flowchart pages specifically; a structurally successful Mermaid render is not enough.
