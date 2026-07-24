# Style Presets

Use one preset for the whole document. Do not mix heading colors, table fills, or spacing systems from different presets.

All presets use SimSun for Chinese and Western text, 12 pt (Chinese small-four) body text,
and 10.5 pt (Chinese five) table text. Heading sizes are two points larger than the
original preset values.

## formal-blue

Default business bid style.

## hejia-yahei

Hejia bid style aligned to the approved Yahei small-four PDF layout.

- Accent: dark blue-gray `243B53`
- Secondary accent: gray-blue `E6EEF6`
- Heading 1: 20 pt, bold, SimSun, restrained dark blue-gray
- Heading 2: 17 pt, bold, SimSun
- Body: 12 pt (small-four), SimSun, 1.3 line spacing
- Tables: compact blue header, light blue alternating rows, unified single line spacing
- Page furniture: upper-shifted header rule, quiet footer rule, company name plus project title
- Supports: part pages, HTML paragraphs, underlines, HTML tables with merged cells, contract scans, image captions, landscape wide tables
- Use for: Chinese property bid Markdown that should match the current approved “微软雅黑小四” output style

- Accent: deep blue `1F4E79`
- Secondary accent: slate `D9EAF7`
- Heading 1: 20 pt, bold, blue, page break before after front matter
- Heading 2: 17 pt, bold, dark gray
- Body: 12 pt, 1.35 line spacing, first-line indent for Chinese prose
- Lists: real Word numbering, explicit restart per Markdown list block, 0.74 cm text indent with 0.74 cm hanging indent
- Page furniture: branded Hejia header/footer with quiet blue rules; cover first page remains clean
- Tables: blue header, light blue banding, compact but not cramped
- Use for: service bids, technical proposals, long formal reports

## government-red

Formal government procurement style.

- Accent: red `9E1B1B`
- Secondary accent: warm gray `F3E8E8`
- Heading 1: 20 pt, bold, red, strong spacing
- Heading 2: 17 pt, bold, black
- Body: 12 pt, 1.35 line spacing
- Tables: red header, pale red banding, conservative borders
- Use for: government procurement, state-owned enterprise tenders

## elegant-black

Minimal serious manuscript style.

- Accent: black `111111`
- Secondary accent: light gray `EDEDED`
- Heading 1: 20 pt, bold, black
- Heading 2: 16.5 pt, bold, black
- Body: 12 pt, 1.35 line spacing
- Tables: grayscale header, thin borders
- Use for: strict formal documents where color should be restrained

## compact-bid

High-density long bid style.

- Accent: dark blue-gray `243B53`
- Secondary accent: gray-blue `E6EEF6`
- Heading 1: 18 pt, bold, page break before
- Heading 2: 15.5 pt, bold
- Body: 12 pt, 1.25 line spacing
- Tables: compact padding, repeated headers, narrower body font
- Use for: manuscripts over 150,000 Chinese characters or documents where page count matters

## Diagram Themes

Use diagram colors that match the document preset:

- `bid-blue`: formal-blue, compact-bid
- `bid-red`: government-red
- `bid-mono`: elegant-black

Mermaid diagrams should use white backgrounds for Word and print.
