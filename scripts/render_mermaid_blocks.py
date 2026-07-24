from __future__ import annotations

import argparse
import json
import math
import re
import textwrap
import urllib.error
import urllib.request
from pathlib import Path

from common import (
    FENCE_RE,
    HEADING_RE,
    find_executable,
    is_mermaid_info,
    parse_fence_info,
    read_text,
    relative_posix,
    run_command,
    slugify,
    write_json,
    write_text,
    strip_inline_markup,
)


THEMES = {
    "bid-blue": {
        "primaryColor": "#E6EEF6",
        "primaryBorderColor": "#1F4E79",
        "primaryTextColor": "#1F2933",
        "lineColor": "#486581",
        "secondaryColor": "#F5F8FB",
        "tertiaryColor": "#FFFFFF",
        "fontFamily": "Microsoft YaHei, SimSun, Arial",
    },
    "bid-red": {
        "primaryColor": "#F3E8E8",
        "primaryBorderColor": "#9E1B1B",
        "primaryTextColor": "#1F1F1F",
        "lineColor": "#7A1E1E",
        "secondaryColor": "#FAF5F5",
        "tertiaryColor": "#FFFFFF",
        "fontFamily": "Microsoft YaHei, SimSun, Arial",
    },
    "bid-mono": {
        "primaryColor": "#EDEDED",
        "primaryBorderColor": "#111111",
        "primaryTextColor": "#111111",
        "lineColor": "#555555",
        "secondaryColor": "#F7F7F7",
        "tertiaryColor": "#FFFFFF",
        "fontFamily": "Microsoft YaHei, SimSun, Arial",
    },
}


FLOWCHART_HEADER_RE = re.compile(r"^\s*(flowchart|graph)\s+(TD|TB|BT|LR|RL)\b", re.IGNORECASE)
NODE_TOKEN_RE = re.compile(r"^\s*([A-Za-z0-9_\-\u4e00-\u9fff]+)\s*(.*)\s*$")
MERMAID_DIRECTIVE_RE = re.compile(r"^\s*(style|classDef|class|linkStyle)\b", re.IGNORECASE)


def with_theme(source: str, theme: str) -> str:
    if "%%{init:" in source:
        return source
    config = {"theme": "base", "themeVariables": THEMES.get(theme, THEMES["bid-blue"])}
    return f"%%{{init: {json.dumps(config, ensure_ascii=True)} }}%%\n{source.strip()}\n"


def strip_mermaid_init(source_text: str) -> str:
    return "\n".join(line for line in source_text.splitlines() if not line.strip().startswith("%%{init:"))


def render_with_mmdc(mmdc: str, source: Path, output: Path, timeout: int) -> tuple[bool, str]:
    cmd = [
        mmdc,
        "-i",
        str(source),
        "-o",
        str(output),
        "-w",
        "2400",
        "--backgroundColor",
        "white",
        "--theme",
        "base",
    ]
    result = run_command(cmd, timeout=timeout)
    if result.ok and output.is_file() and output.stat().st_size > 100:
        return True, ""
    return False, (result.stderr or result.stdout or "mmdc did not create an output image.").strip()


def render_with_kroki(source_text: str, output: Path, timeout: int) -> tuple[bool, str]:
    request = urllib.request.Request(
        "https://kroki.io/mermaid/png",
        data=source_text.encode("utf-8"),
        headers={"Content-Type": "text/plain", "Accept": "image/png"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            body = response.read()
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        return False, f"Kroki HTTP {exc.code}: {detail[:1000]}"
    except Exception as exc:  # noqa: BLE001
        return False, f"Kroki request failed: {exc}"
    if len(body) < 100:
        return False, "Kroki returned an empty or tiny response."
    output.write_bytes(body)
    return True, ""


def load_font(size: int, bold: bool = False):
    from PIL import ImageFont

    candidates = [
        Path("C:/Windows/Fonts/msyhbd.ttc" if bold else "C:/Windows/Fonts/msyh.ttc"),
        Path("C:/Windows/Fonts/simhei.ttf" if bold else "C:/Windows/Fonts/simsun.ttc"),
        Path("C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf"),
    ]
    for candidate in candidates:
        if candidate.is_file():
            try:
                return ImageFont.truetype(str(candidate), size)
            except Exception:
                continue
    return ImageFont.load_default()


def clean_node_label(raw: str) -> tuple[str, str]:
    raw = raw.strip().rstrip(";")
    shape = "rect"
    label = raw
    if raw.startswith("((") and raw.endswith("))"):
        shape = "ellipse"
        label = raw[2:-2]
    elif raw.startswith("{") and raw.endswith("}"):
        shape = "diamond"
        label = raw[1:-1]
    elif raw.startswith("[") and raw.endswith("]"):
        shape = "rect"
        label = raw[1:-1]
    elif raw.startswith("(") and raw.endswith(")"):
        shape = "round"
        label = raw[1:-1]
    elif raw.startswith(">") and raw.endswith("]"):
        shape = "rect"
        label = raw[2:-1]
    label = label.strip().strip("\"'")
    label = re.sub(r"<br\s*/?>", "\n", label, flags=re.IGNORECASE)
    label = re.sub(r"<[^>]+>", "", label)
    label = label.replace("&nbsp;", " ").replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
    return label or raw, shape


def parse_node_token(token: str) -> tuple[str, str, str]:
    token = token.strip().rstrip(";")
    match = NODE_TOKEN_RE.match(token)
    if not match:
        fallback = token.strip() or "node"
        return fallback, fallback, "rect"
    node_id, raw_label = match.groups()
    if raw_label.strip():
        label, shape = clean_node_label(raw_label)
    else:
        label, shape = node_id, "rect"
    return node_id, label, shape


def split_mermaid_side(side: str) -> list[str]:
    parts = [item.strip() for item in side.split("&")]
    return [item for item in parts if item]


def parse_flowchart(source_text: str) -> tuple[str, dict[str, dict], list[dict], str | None]:
    source = strip_mermaid_init(source_text)
    direction = "TD"
    nodes: dict[str, dict] = {}
    edges: list[dict] = []
    found_header = False

    for raw_line in source.splitlines():
        line = raw_line.strip()
        if not line or line.startswith("%%"):
            continue
        header = FLOWCHART_HEADER_RE.match(line)
        if header:
            direction = header.group(2).upper()
            found_header = True
            continue
        if line.lower() in {"end"} or line.lower().startswith("subgraph "):
            continue
        if MERMAID_DIRECTIVE_RE.match(line):
            continue
        line = line.rstrip(";")
        normalized = line.replace("-.->", "-->").replace("==>", "-->")
        if "-->" in normalized:
            left, right = normalized.split("-->", 1)
            label = ""
            if "--" in left:
                source_side, label = left.rsplit("--", 1)
            else:
                source_side = left
            right = right.strip()
            pipe_label = re.match(r"^\|([^|]+)\|\s*(.+)$", right)
            if pipe_label:
                label = pipe_label.group(1)
                target_side = pipe_label.group(2)
            else:
                target_side = right
            for source_token in split_mermaid_side(source_side):
                source_id, source_label, source_shape = parse_node_token(source_token)
                nodes.setdefault(source_id, {"id": source_id, "label": source_label, "shape": source_shape})
                for target_token in split_mermaid_side(target_side):
                    target_id, target_label, target_shape = parse_node_token(target_token)
                    nodes.setdefault(target_id, {"id": target_id, "label": target_label, "shape": target_shape})
                    edges.append({"source": source_id, "target": target_id, "label": label.strip()})
        else:
            node_id, label, shape = parse_node_token(normalized)
            if node_id:
                nodes.setdefault(node_id, {"id": node_id, "label": label, "shape": shape})

    if not found_header:
        return direction, nodes, edges, "Not a flowchart/graph block."
    if not nodes:
        return direction, nodes, edges, "No flowchart nodes parsed."
    return direction, nodes, edges, None


def wrapped_text(draw, text: str, font, max_width: int) -> list[str]:
    chunks = []
    for paragraph in text.splitlines() or [text]:
        current = ""
        for char in paragraph:
            trial = current + char
            width = draw.textbbox((0, 0), trial, font=font)[2]
            if current and width > max_width:
                chunks.append(current)
                current = char
            else:
                current = trial
        if current:
            chunks.append(current)
    return chunks or [text]


def draw_arrow(draw, points: list[tuple[float, float]], fill: str, width: int = 4) -> None:
    if len(points) < 2:
        return
    draw.line(points, fill=fill, width=width, joint="curve")
    x1, y1 = points[-2]
    x2, y2 = points[-1]
    angle = math.atan2(y2 - y1, x2 - x1)
    size = 16
    left = (x2 - size * math.cos(angle - math.pi / 6), y2 - size * math.sin(angle - math.pi / 6))
    right = (x2 - size * math.cos(angle + math.pi / 6), y2 - size * math.sin(angle + math.pi / 6))
    draw.polygon([(x2, y2), left, right], fill=fill)


def edge_key(edge: dict) -> tuple[str, str, str]:
    return (edge["source"], edge["target"], edge.get("label", ""))


def path_exists(adjacency: dict[str, list[str]], start: str, target: str, ignored: tuple[str, str]) -> bool:
    stack = [start]
    seen: set[str] = set()
    while stack:
        node = stack.pop()
        if node == target:
            return True
        if node in seen:
            continue
        seen.add(node)
        for nxt in adjacency.get(node, []):
            if (node, nxt) == ignored:
                continue
            stack.append(nxt)
    return False


def compute_layers(node_ids: list[str], edges: list[dict]) -> tuple[dict[str, int], set[tuple[str, str, str]]]:
    order = {node_id: index for index, node_id in enumerate(node_ids)}
    adjacency: dict[str, list[str]] = {node_id: [] for node_id in node_ids}
    for edge in edges:
        adjacency.setdefault(edge["source"], []).append(edge["target"])
        adjacency.setdefault(edge["target"], [])

    cycle_edges: set[tuple[str, str, str]] = set()
    for edge in edges:
        if order.get(edge["target"], 0) <= order.get(edge["source"], 0):
            cycle_edges.add(edge_key(edge))

    dag_edges = [edge for edge in edges if edge_key(edge) not in cycle_edges]
    indegree = {node_id: 0 for node_id in node_ids}
    dag_adjacency: dict[str, list[str]] = {node_id: [] for node_id in node_ids}
    for edge in dag_edges:
        dag_adjacency.setdefault(edge["source"], []).append(edge["target"])
        indegree[edge["target"]] = indegree.get(edge["target"], 0) + 1
        indegree.setdefault(edge["source"], 0)

    roots = [node_id for node_id in node_ids if indegree.get(node_id, 0) == 0]
    if not roots and node_ids:
        roots = [node_ids[0]]
    layers = {node_id: 0 for node_id in node_ids}
    queue = roots[:]
    processed: set[str] = set()
    while queue:
        node = queue.pop(0)
        if node in processed:
            continue
        processed.add(node)
        for target in dag_adjacency.get(node, []):
            layers[target] = max(layers.get(target, 0), layers.get(node, 0) + 1)
            indegree[target] -= 1
            if indegree[target] <= 0:
                queue.append(target)

    # Any remaining nodes belong to complicated cycles or disconnected islands.
    # Place them near their first incoming source instead of stretching the graph.
    for node_id in node_ids:
        if node_id not in processed:
            incoming = [edge for edge in edges if edge["target"] == node_id and edge["source"] in layers]
            if incoming:
                layers[node_id] = min(layers[incoming[0]["source"]] + 1, max(layers.values(), default=0) + 1)
            else:
                layers[node_id] = max(layers.values(), default=0) + 1
    return layers, cycle_edges


def render_builtin_flowchart(source_text: str, output: Path, theme: str) -> tuple[bool, str]:
    from PIL import Image, ImageDraw

    direction, nodes, edges, error = parse_flowchart(source_text)
    if error:
        return False, error

    horizontal = direction in {"LR", "RL"}
    node_ids = list(nodes)
    layers, cycle_edges = compute_layers(node_ids, edges)

    grouped: dict[int, list[str]] = {}
    for node_id in node_ids:
        grouped.setdefault(layers.get(node_id, 0), []).append(node_id)
    max_nodes_per_layer = 6
    visual_slots: dict[str, tuple[int, int, int, int]] = {}
    visual_grouped: dict[int, list[tuple[str, int, int]]] = {}
    for layer, ids in grouped.items():
        ordered = sorted(ids, key=node_ids.index)
        for idx, node_id in enumerate(ordered):
            subrow = idx // max_nodes_per_layer
            subcol = idx % max_nodes_per_layer
            subrows = (len(ordered) + max_nodes_per_layer - 1) // max_nodes_per_layer
            visual_slots[node_id] = (layer, subrow, subcol, subrows)
            visual_grouped.setdefault(layer, []).append((node_id, subrow, subcol))

    font = load_font(34)
    label_font = load_font(26)
    temp = Image.new("RGB", (10, 10), "white")
    temp_draw = ImageDraw.Draw(temp)

    node_w = 280
    node_h = 110
    diamond_w = 230
    margin_x = 120
    margin_y = 110
    gap_x = 190 if horizontal else 90
    gap_y = 120 if horizontal else 120
    max_layer = max(grouped) if grouped else 0
    max_count = min(max((len(items) for items in grouped.values()), default=1), max_nodes_per_layer)
    max_subrows = max((slot[3] for slot in visual_slots.values()), default=1)
    folded = (not horizontal) and max_layer >= 8
    fold_rows = 7
    if horizontal:
        width = margin_x * 2 + (max_layer + 1) * node_w + max_layer * gap_x
        height = margin_y * 2 + max_count * node_h + (max_count - 1) * gap_y
    elif folded:
        folded_gap_x = 150
        folded_gap_y = 90
        col_count = max_layer // fold_rows + 1
        col_widths: list[float] = []
        for col in range(col_count):
            layer_numbers = range(col * fold_rows, min((col + 1) * fold_rows, max_layer + 1))
            widest = max(
                (
                    min(len(grouped.get(layer, [])), max_nodes_per_layer) * node_w
                    + max(min(len(grouped.get(layer, [])), max_nodes_per_layer) - 1, 0) * 50
                    for layer in layer_numbers
                ),
                default=node_w,
            )
            col_widths.append(max(widest, node_w))
        width = margin_x * 2 + sum(col_widths) + folded_gap_x * max(col_count - 1, 0)
        used_rows = min(fold_rows, max_layer + 1)
        height = margin_y * 2 + used_rows * (node_h + (max_subrows - 1) * 85) + (used_rows - 1) * folded_gap_y
    else:
        width = margin_x * 2 + max_count * node_w + (max_count - 1) * gap_x
        height = margin_y * 2 + (max_layer + 1) * (node_h + (max_subrows - 1) * 85) + max_layer * gap_y
    width = max(width, 1200)
    height = max(height, 760)

    palette = THEMES.get(theme, THEMES["bid-blue"])
    fill = palette["primaryColor"]
    stroke = palette["primaryBorderColor"]
    text_color = palette["primaryTextColor"]
    line_color = palette["lineColor"]
    secondary = palette["secondaryColor"]

    image = Image.new("RGB", (int(width), int(height)), "white")
    draw = ImageDraw.Draw(image)
    positions: dict[str, tuple[float, float, float, float]] = {}

    for layer, ids in grouped.items():
        ids = sorted(ids, key=node_ids.index)
        for idx, node_id in enumerate(ids):
            _layer, subrow, subcol, subrows = visual_slots[node_id]
            visible_count = min(len(ids) - subrow * max_nodes_per_layer, max_nodes_per_layer)
            if horizontal:
                x = margin_x + layer * (node_w + gap_x)
                total_h = visible_count * node_h + (visible_count - 1) * gap_y
                y = (height - total_h) / 2 + subcol * (node_h + gap_y)
            elif folded:
                col = layer // fold_rows
                row = layer % fold_rows
                col_start = margin_x + sum(col_widths[:col]) + folded_gap_x * col
                group_width = visible_count * node_w + (visible_count - 1) * 50
                x = col_start + (col_widths[col] - group_width) / 2 + idx * (node_w + 50)
                x = col_start + (col_widths[col] - group_width) / 2 + subcol * (node_w + 50)
                y = margin_y + row * (node_h + (max_subrows - 1) * 85 + folded_gap_y) + subrow * 85
            else:
                total_w = visible_count * node_w + (visible_count - 1) * gap_x
                x = (width - total_w) / 2 + subcol * (node_w + gap_x)
                y = margin_y + layer * (node_h + (max_subrows - 1) * 85 + gap_y) + subrow * 85
            w = diamond_w if nodes[node_id]["shape"] == "diamond" else node_w
            h = node_h
            if nodes[node_id]["shape"] == "diamond":
                x = x + (node_w - diamond_w) / 2
            positions[node_id] = (x, y, w, h)

    for edge in edges:
        if edge["source"] not in positions or edge["target"] not in positions:
            continue
        sx, sy, sw, sh = positions[edge["source"]]
        tx, ty, tw, th = positions[edge["target"]]
        source_layer = layers.get(edge["source"], 0)
        target_layer = layers.get(edge["target"], 0)
        is_cycle = edge_key(edge) in cycle_edges or target_layer <= source_layer
        if horizontal:
            start = (sx + sw, sy + sh / 2)
            end = (tx, ty + th / 2)
            mid_x = (start[0] + end[0]) / 2
            points = [start, (mid_x, start[1]), (mid_x, end[1]), end]
            label_pos = (mid_x + 8, (start[1] + end[1]) / 2 - 18)
        elif folded and (tx > sx + sw + 30 or sx > tx + tw + 30):
            if tx > sx:
                start = (sx + sw, sy + sh / 2)
                end = (tx, ty + th / 2)
            else:
                start = (sx, sy + sh / 2)
                end = (tx + tw, ty + th / 2)
            mid_x = (start[0] + end[0]) / 2
            points = [start, (mid_x, start[1]), (mid_x, end[1]), end]
            label_pos = (mid_x + 8, (start[1] + end[1]) / 2 - 18)
        elif is_cycle:
            route_x = max(sx + sw, tx + tw) + 70
            start = (sx + sw, sy + sh / 2)
            end = (tx + tw, ty + th / 2)
            points = [start, (route_x, start[1]), (route_x, end[1]), end]
            label_pos = (route_x + 8, (start[1] + end[1]) / 2 - 18)
        else:
            start = (sx + sw / 2, sy + sh)
            end = (tx + tw / 2, ty)
            mid_y = (start[1] + end[1]) / 2
            points = [start, (start[0], mid_y), (end[0], mid_y), end]
            label_pos = ((start[0] + end[0]) / 2 + 8, mid_y - 30)
        draw_arrow(draw, points, line_color)
        if edge.get("label"):
            bbox = draw.textbbox(label_pos, edge["label"], font=label_font)
            pad = 8
            draw.rounded_rectangle(
                (bbox[0] - pad, bbox[1] - pad, bbox[2] + pad, bbox[3] + pad),
                radius=10,
                fill="white",
                outline=secondary,
            )
            draw.text(label_pos, edge["label"], font=label_font, fill=text_color)

    for node_id, (x, y, w, h) in positions.items():
        shape = nodes[node_id]["shape"]
        if shape == "diamond":
            points = [(x + w / 2, y), (x + w, y + h / 2), (x + w / 2, y + h), (x, y + h / 2)]
            draw.polygon(points, fill=fill, outline=stroke)
            draw.line(points + [points[0]], fill=stroke, width=4)
        elif shape == "ellipse":
            draw.ellipse((x, y, x + w, y + h), fill=fill, outline=stroke, width=4)
        else:
            draw.rounded_rectangle((x, y, x + w, y + h), radius=20, fill=fill, outline=stroke, width=4)
        label = nodes[node_id]["label"]
        lines = wrapped_text(temp_draw, label, font, int(w - 36))
        line_height = draw.textbbox((0, 0), "国", font=font)[3] + 8
        total_text_h = len(lines) * line_height
        ty = y + (h - total_text_h) / 2 - 2
        for item in lines:
            bbox = draw.textbbox((0, 0), item, font=font)
            draw.text((x + (w - (bbox[2] - bbox[0])) / 2, ty), item, font=font, fill=text_color)
            ty += line_height

    # Add a quiet frame so the diagram feels intentional on a white Word page.
    draw.rounded_rectangle((24, 24, width - 24, height - 24), radius=24, outline=secondary, width=2)
    output.parent.mkdir(parents=True, exist_ok=True)
    image.save(output)
    return True, ""


def write_placeholder_png(output: Path, caption: str, message: str) -> None:
    from PIL import Image, ImageDraw

    image = Image.new("RGB", (1600, 900), "white")
    draw = ImageDraw.Draw(image)
    y = 90
    draw.rectangle((60, 60, 1540, 840), outline="#9E1B1B", width=4)
    draw.text((100, y), "Mermaid render failed", fill="#9E1B1B")
    y += 70
    for line in textwrap.wrap(caption, width=70)[:3]:
        draw.text((100, y), line, fill="#111111")
        y += 42
    y += 20
    for line in textwrap.wrap(message, width=95)[:12]:
        draw.text((100, y), line, fill="#333333")
        y += 34
    output.parent.mkdir(parents=True, exist_ok=True)
    image.save(output)


def extract_caption(info: str, index: int, fallback_heading: str = "") -> str:
    _lang, attrs = parse_fence_info(info)
    if attrs.get("caption") or attrs.get("title"):
        return attrs.get("caption") or attrs.get("title") or ""
    if fallback_heading:
        clean = re.sub(r"^\d+(?:\.\d+)*\s*", "", fallback_heading).strip()
        return f"{clean}示意图" if "图" not in clean else clean
    return f"图 {index:03d}"


def render_blocks(input_path: Path, out_dir: Path, theme: str, use_kroki: bool, timeout: int) -> dict:
    text = read_text(input_path)
    lines = text.splitlines()
    figures_dir = out_dir / "figures"
    figures_dir.mkdir(parents=True, exist_ok=True)
    normalized: list[str] = []
    manifest: list[dict] = []
    mmdc = find_executable("mmdc")
    index = 0
    cursor = 0
    last_heading = ""

    while cursor < len(lines):
        line = lines[cursor]
        fence = FENCE_RE.match(line)
        if not fence or not is_mermaid_info(fence.group(3).strip()):
            heading = HEADING_RE.match(line)
            if heading:
                last_heading = strip_inline_markup(heading.group(2))
            normalized.append(line)
            cursor += 1
            continue

        opening_marker = fence.group(2)
        info = fence.group(3).strip()
        start_line = cursor + 1
        body: list[str] = []
        cursor += 1
        closed = False
        while cursor < len(lines):
            close = FENCE_RE.match(lines[cursor])
            if close and close.group(2)[0] == opening_marker[0] and len(close.group(2)) >= len(opening_marker):
                closed = True
                cursor += 1
                break
            body.append(lines[cursor])
            cursor += 1

        index += 1
        caption = extract_caption(info, index, last_heading)
        stem = f"fig-{index:03d}-{slugify(caption, 'diagram')}"
        source_path = figures_dir / f"{stem}.mmd"
        image_path = figures_dir / f"{stem}.png"
        source_text = with_theme("\n".join(body), theme)
        write_text(source_path, source_text)

        status = "failed"
        renderer = ""
        error = ""
        if mmdc:
            renderer = "mmdc"
            ok, error = render_with_mmdc(mmdc, source_path, image_path, timeout)
            if ok:
                status = "rendered"
        if status != "rendered":
            renderer = "builtin-flowchart"
            ok, error = render_builtin_flowchart(source_text, image_path, theme)
            if ok:
                status = "rendered"
        if status != "rendered" and use_kroki:
            renderer = "kroki"
            ok, error = render_with_kroki(source_text, image_path, timeout)
            if ok:
                status = "rendered"
        if status != "rendered":
            renderer = renderer or "none"
            write_placeholder_png(image_path, caption, error or "No Mermaid renderer is available.")

        rel_image = relative_posix(image_path, out_dir)
        normalized.append(f'![{caption}]({rel_image} "{caption}")')
        normalized.append("")
        manifest.append(
            {
                "index": index,
                "caption": caption,
                "start_line": start_line,
                "closed": closed,
                "renderer": renderer,
                "status": status,
                "source": str(source_path),
                "image": str(image_path),
                "error": error,
            }
        )

    normalized_path = out_dir / "normalized.md"
    write_text(normalized_path, "\n".join(normalized).rstrip() + "\n")
    result = {
        "input": str(input_path),
        "normalized_markdown": str(normalized_path),
        "theme": theme,
        "mmdc_available": bool(mmdc),
        "kroki_enabled": use_kroki,
        "diagram_count": len(manifest),
        "failed_count": sum(1 for item in manifest if item["status"] != "rendered"),
        "diagrams": manifest,
    }
    write_json(out_dir / "diagram-manifest.json", result)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description="Render Mermaid blocks and write normalized Markdown.")
    parser.add_argument("input", type=Path)
    parser.add_argument("--out-dir", type=Path, required=True)
    parser.add_argument("--theme", choices=sorted(THEMES), default="bid-blue")
    parser.add_argument("--no-kroki", action="store_true")
    parser.add_argument("--timeout", type=int, default=60)
    args = parser.parse_args()

    input_path = args.input.resolve()
    if not input_path.is_file():
        parser.error(f"Input Markdown not found: {input_path}")
    out_dir = args.out_dir.resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    result = render_blocks(input_path, out_dir, args.theme, not args.no_kroki, args.timeout)
    print(f"Wrote {result['normalized_markdown']}")
    print(f"Rendered diagrams: {result['diagram_count'] - result['failed_count']}/{result['diagram_count']}")
    return 0 if result["failed_count"] == 0 else 3


if __name__ == "__main__":
    raise SystemExit(main())
