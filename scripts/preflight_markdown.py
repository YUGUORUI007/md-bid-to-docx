from __future__ import annotations

import argparse
from collections import Counter
from pathlib import Path

from common import (
    FENCE_RE,
    HEADING_RE,
    IMAGE_RE,
    TODO_RE,
    is_local_resource,
    is_mermaid_info,
    is_table_separator,
    read_text,
    split_table_row,
    strip_inline_markup,
    write_json,
    write_text,
)


def analyze_markdown(input_path: Path) -> dict:
    text = read_text(input_path)
    lines = text.splitlines()
    issues: list[dict] = []
    headings: list[dict] = []
    images: list[dict] = []
    mermaid_blocks: list[dict] = []
    tables: list[dict] = []
    placeholders: list[dict] = []

    in_fence = False
    fence_marker = ""
    fence_start = 0
    fence_info = ""
    fence_body: list[str] = []

    for index, line in enumerate(lines, start=1):
        fence_match = FENCE_RE.match(line)
        if fence_match:
            marker = fence_match.group(2)
            info = fence_match.group(3).strip()
            if not in_fence:
                in_fence = True
                fence_marker = marker
                fence_start = index
                fence_info = info
                fence_body = []
                continue
            if marker[0] == fence_marker[0] and len(marker) >= len(fence_marker):
                if is_mermaid_info(fence_info):
                    mermaid_blocks.append(
                        {
                            "start_line": fence_start,
                            "end_line": index,
                            "line_count": len(fence_body),
                            "info": fence_info,
                        }
                    )
                    if not "".join(fence_body).strip():
                        issues.append(
                            {
                                "severity": "fatal",
                                "line": fence_start,
                                "kind": "empty-mermaid",
                                "message": "Mermaid block is empty.",
                            }
                        )
                in_fence = False
                fence_marker = ""
                fence_info = ""
                fence_body = []
                continue
        if in_fence:
            fence_body.append(line)
            continue

        heading_match = HEADING_RE.match(line)
        if heading_match:
            level = len(heading_match.group(1))
            title = strip_inline_markup(heading_match.group(2))
            headings.append({"line": index, "level": level, "title": title})
            if not title:
                issues.append(
                    {
                        "severity": "fatal",
                        "line": index,
                        "kind": "empty-heading",
                        "message": "Heading has no visible text.",
                    }
                )

        for image_match in IMAGE_RE.finditer(line):
            alt, uri, title = image_match.groups()
            item = {"line": index, "alt": alt, "uri": uri, "title": title or ""}
            if is_local_resource(uri):
                resource = (input_path.parent / uri).resolve()
                item["resolved"] = str(resource)
                item["exists"] = resource.is_file()
                if not resource.is_file():
                    issues.append(
                        {
                            "severity": "fatal",
                            "line": index,
                            "kind": "missing-image",
                            "message": f"Local image not found: {uri}",
                        }
                    )
            images.append(item)

        if TODO_RE.search(line):
            placeholders.append({"line": index, "text": line.strip()[:240]})

    if in_fence:
        issues.append(
            {
                "severity": "fatal",
                "line": fence_start,
                "kind": "unclosed-fence",
                "message": "Fenced code block is not closed.",
            }
        )

    for previous, current in zip(headings, headings[1:]):
        if current["level"] > previous["level"] + 1:
            issues.append(
                {
                    "severity": "warning",
                    "line": current["line"],
                    "kind": "heading-level-jump",
                    "message": f"Heading jumps from H{previous['level']} to H{current['level']}.",
                }
            )

    heading_counter = Counter((h["level"], h["title"]) for h in headings if h["title"])
    for heading in headings:
        if heading_counter[(heading["level"], heading["title"])] > 1 and heading["level"] <= 2:
            issues.append(
                {
                    "severity": "warning",
                    "line": heading["line"],
                    "kind": "duplicate-heading",
                    "message": f"Repeated H{heading['level']} heading: {heading['title']}",
                }
            )

    index = 0
    while index < len(lines) - 1:
        if "|" in lines[index] and is_table_separator(lines[index + 1]):
            start = index + 1
            header = split_table_row(lines[index])
            row_count = 0
            max_cols = len(header)
            cursor = index + 2
            while cursor < len(lines) and "|" in lines[cursor].strip():
                cells = split_table_row(lines[cursor])
                max_cols = max(max_cols, len(cells))
                row_count += 1
                cursor += 1
            table = {
                "start_line": start,
                "columns": max_cols,
                "body_rows": row_count,
                "header": header,
            }
            tables.append(table)
            if max_cols > 6:
                issues.append(
                    {
                        "severity": "warning",
                        "line": start,
                        "kind": "wide-table",
                        "message": f"Table has {max_cols} columns; consider splitting or using compact/landscape treatment.",
                    }
                )
            index = cursor
        else:
            index += 1

    fatal_count = sum(1 for issue in issues if issue["severity"] == "fatal")
    warning_count = sum(1 for issue in issues if issue["severity"] == "warning")
    chinese_chars = sum(1 for ch in text if "\u4e00" <= ch <= "\u9fff")
    report = {
        "input": str(input_path),
        "line_count": len(lines),
        "character_count": len(text),
        "chinese_character_count": chinese_chars,
        "heading_count": len(headings),
        "image_count": len(images),
        "mermaid_block_count": len(mermaid_blocks),
        "table_count": len(tables),
        "placeholder_count": len(placeholders),
        "fatal_count": fatal_count,
        "warning_count": warning_count,
        "headings": headings,
        "images": images,
        "mermaid_blocks": mermaid_blocks,
        "tables": tables,
        "placeholders": placeholders,
        "issues": issues,
    }
    return report


def report_to_markdown(report: dict) -> str:
    lines = [
        "# Markdown Preflight Report",
        "",
        f"- Input: `{report['input']}`",
        f"- Lines: {report['line_count']}",
        f"- Characters: {report['character_count']}",
        f"- Chinese characters: {report['chinese_character_count']}",
        f"- Headings: {report['heading_count']}",
        f"- Tables: {report['table_count']}",
        f"- Images: {report['image_count']}",
        f"- Mermaid blocks: {report['mermaid_block_count']}",
        f"- Placeholders: {report['placeholder_count']}",
        f"- Fatal issues: {report['fatal_count']}",
        f"- Warnings: {report['warning_count']}",
        "",
        "## Issues",
        "",
    ]
    if not report["issues"]:
        lines.append("No issues detected.")
    else:
        lines.append("| Severity | Line | Kind | Message |")
        lines.append("| --- | ---: | --- | --- |")
        for issue in report["issues"]:
            lines.append(
                f"| {issue['severity']} | {issue.get('line', '')} | {issue['kind']} | {issue['message']} |"
            )

    if report["placeholders"]:
        lines.append("")
        lines.append(f"> Warning: {report['placeholder_count']} placeholder(s) remain in the manuscript.")
        lines.extend(["", "## Placeholders", "", "| Line | Text |", "| ---: | --- |"])
        for item in report["placeholders"][:200]:
            text = item["text"].replace("|", "\\|")
            lines.append(f"| {item['line']} | {text} |")
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description="Preflight a Markdown bid manuscript.")
    parser.add_argument("input", type=Path)
    parser.add_argument("--out-dir", type=Path, default=None)
    parser.add_argument("--fail-on-fatal", action="store_true")
    args = parser.parse_args()

    input_path = args.input.resolve()
    if not input_path.is_file():
        parser.error(f"Input Markdown not found: {input_path}")
    out_dir = (args.out_dir or input_path.with_suffix("")).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    report = analyze_markdown(input_path)
    write_json(out_dir / "preflight-report.json", report)
    write_text(out_dir / "preflight-report.md", report_to_markdown(report))
    print(f"Wrote {out_dir / 'preflight-report.md'}")
    if args.fail_on_fatal and report["fatal_count"]:
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
