from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
ASSET_DIR = SKILL_DIR / "assets"
PANDOC_ASSET_DIR = ASSET_DIR / "pandoc-docx-template"

DEFAULT_COMPANY_NAME = "浙江合家物业发展有限公司"
DEFAULT_COMPANY_WEBSITE = "www.hejiawuye.cn"
DEFAULT_COMPANY_SLOGAN = "主动 高效 专业 立体"

HEADING_RE = re.compile(r"^(#{1,6})\s+(.+?)(?:\s+#+)?\s*$")
FENCE_RE = re.compile(r"^(\s*)(`{3,}|~{3,})(.*)$")
IMAGE_RE = re.compile(r"!\[([^\]]*)\]\(([^)\s]+)(?:\s+\"([^\"]*)\")?\)")
ATTR_RE = re.compile(r"([A-Za-z_][\w-]*)=(\"[^\"]*\"|'[^']*'|[^\s}]+)")
TODO_RE = re.compile(r"(TODO|TBD|待补充|待完善|待确认|\[待[^\]]+\])", re.IGNORECASE)
GENERIC_CHAPTER_TITLE_RE = re.compile(r"^第[一二三四五六七八九十百千万\d]+[章节篇部分]\s*")
PROJECT_NAME_RE = re.compile(r"(?:项目名称|工程名称|采购项目|招标项目|选聘项目)\s*[:：]\s*(.+)")
QUOTED_BID_RE = re.compile(r"《([^》]{4,100}(?:项目|物业|服务|工程|采购|选聘|招标)[^》]{0,40})》")


STYLE_PRESETS = {
    "hejia-yahei": {
        "accent": "243B53",
        "accent_light": "E6EEF6",
        "accent_mid": "486581",
        "body_size": 12,
        "body_spacing": 1.3,
        "table_font": 10.5,
        "heading1": 20,
        "heading2": 17,
        "font_east_asia": "SimSun",
        "font_ascii": "SimSun",
        "heading_font_east_asia": "SimSun",
        "heading_font_ascii": "SimSun",
    },
    "formal-blue": {
        "accent": "1F4E79",
        "accent_light": "D9EAF7",
        "accent_mid": "5B9BD5",
        "body_size": 12,
        "body_spacing": 1.35,
        "table_font": 10.5,
        "heading1": 20,
        "heading2": 17,
        "font_east_asia": "SimSun",
        "font_ascii": "SimSun",
        "heading_font_east_asia": "SimSun",
        "heading_font_ascii": "SimSun",
    },
    "government-red": {
        "accent": "9E1B1B",
        "accent_light": "F3E8E8",
        "accent_mid": "B33A3A",
        "body_size": 12,
        "body_spacing": 1.35,
        "table_font": 10.5,
        "heading1": 20,
        "heading2": 17,
        "font_east_asia": "SimSun",
        "font_ascii": "SimSun",
        "heading_font_east_asia": "SimSun",
        "heading_font_ascii": "SimSun",
    },
    "elegant-black": {
        "accent": "111111",
        "accent_light": "EDEDED",
        "accent_mid": "666666",
        "body_size": 12,
        "body_spacing": 1.35,
        "table_font": 10.5,
        "heading1": 20,
        "heading2": 16.5,
        "font_east_asia": "SimSun",
        "font_ascii": "SimSun",
        "heading_font_east_asia": "SimSun",
        "heading_font_ascii": "SimSun",
    },
    "compact-bid": {
        "accent": "243B53",
        "accent_light": "E6EEF6",
        "accent_mid": "486581",
        "body_size": 12,
        "body_spacing": 1.25,
        "table_font": 10.5,
        "heading1": 18,
        "heading2": 15.5,
        "font_east_asia": "SimSun",
        "font_ascii": "SimSun",
        "heading_font_east_asia": "SimSun",
        "heading_font_ascii": "SimSun",
    },
}

# Override legacy mojibake defaults with stable visible text.
DEFAULT_COMPANY_NAME = "浙江合家物业发展有限公司"
DEFAULT_COMPANY_SLOGAN = "主动 高效 专业 立体"


@dataclass
class CommandResult:
    ok: bool
    returncode: int
    stdout: str
    stderr: str


def skill_python() -> str:
    return sys.executable


def read_text(path: Path) -> str:
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def write_json(path: Path, data: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def find_executable(name: str, extra_candidates: Iterable[Path | str] = ()) -> str | None:
    found = shutil.which(name)
    if found:
        return found
    if os.name == "nt" and not name.lower().endswith((".exe", ".cmd", ".bat")):
        for suffix in (".cmd", ".exe", ".bat"):
            found = shutil.which(name + suffix)
            if found:
                return found
    for candidate in extra_candidates:
        path = Path(candidate)
        if path.is_file():
            return str(path)
    return None


def existing_path(path: Path | str | None) -> Path | None:
    if not path:
        return None
    candidate = Path(path)
    return candidate if candidate.is_file() else None


def newest_existing(paths: Iterable[Path]) -> Path | None:
    existing = [path for path in paths if path.is_file()]
    if not existing:
        return None
    return sorted(existing, key=lambda p: p.stat().st_mtime, reverse=True)[0]


def find_document_renderer() -> Path | None:
    env_path = existing_path(os.environ.get("DOCX_RENDER_SCRIPT"))
    if env_path:
        return env_path
    home = Path.home()
    candidates = [
        home / ".codex/plugins/cache/openai-primary-runtime/documents/26.619.11828/skills/documents/render_docx.py",
    ]
    cache = home / ".codex/plugins/cache"
    if cache.is_dir():
        candidates.extend(cache.glob("**/skills/documents/render_docx.py"))
    return newest_existing(candidates)


def find_codex_runtime_tool(name: str) -> Path | None:
    executable = find_executable(name)
    if executable:
        return Path(executable)
    names = [name]
    if os.name == "nt" and not name.lower().endswith(".exe"):
        names.insert(0, f"{name}.exe")
    roots: list[Path] = []
    env_root = os.environ.get("CODEX_RUNTIME_DIR")
    if env_root:
        roots.append(Path(env_root))
    roots.append(Path.home() / ".cache/codex-runtimes")
    candidates: list[Path] = []
    for root in roots:
        if root.is_dir():
            for item_name in names:
                candidates.extend(root.glob(f"**/{item_name}"))
    return newest_existing(candidates)


def run_command(cmd: list[str], *, cwd: Path | None = None, timeout: int | None = None, stdin: bytes | None = None) -> CommandResult:
    try:
        completed = subprocess.run(
            cmd,
            cwd=str(cwd) if cwd else None,
            input=stdin,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout,
            check=False,
        )
    except FileNotFoundError as exc:
        return CommandResult(False, 127, "", str(exc))
    except subprocess.TimeoutExpired as exc:
        stdout = exc.stdout.decode("utf-8", errors="replace") if exc.stdout else ""
        stderr = exc.stderr.decode("utf-8", errors="replace") if exc.stderr else ""
        return CommandResult(False, 124, stdout, stderr or f"Timed out after {timeout}s")
    stdout = completed.stdout.decode("utf-8", errors="replace")
    stderr = completed.stderr.decode("utf-8", errors="replace")
    return CommandResult(completed.returncode == 0, completed.returncode, stdout, stderr)


def slugify(value: str, fallback: str = "item") -> str:
    value = value.strip().lower()
    value = re.sub(r"[^\w\-]+", "-", value, flags=re.ASCII)
    value = re.sub(r"-{2,}", "-", value).strip("-")
    return value or fallback


def parse_fence_info(info: str) -> tuple[str, dict[str, str]]:
    info = info.strip()
    attrs: dict[str, str] = {}
    for key, raw in ATTR_RE.findall(info):
        attrs[key.lower()] = raw.strip("\"'")
    tokens = [t for t in re.split(r"\s+", info) if t and not t.startswith("{")]
    lang = ""
    if tokens:
        first = tokens[0].strip("{}").strip()
        if first.startswith("."):
            first = first[1:]
        lang = first.lower()
    if not lang and ".mermaid" in info.lower():
        lang = "mermaid"
    if "mermaid" in info.lower() and not lang:
        lang = "mermaid"
    return lang, attrs


def is_mermaid_info(info: str) -> bool:
    lang, _attrs = parse_fence_info(info)
    return lang == "mermaid" or ".mermaid" in info.lower()


def is_generic_document_heading(title: str) -> bool:
    clean = strip_inline_markup(title).strip()
    if not clean:
        return True
    if GENERIC_CHAPTER_TITLE_RE.match(clean):
        return True
    generic_titles = {"目录", "技术标", "商务标", "资信标", "投标文件", "技术偏离表"}
    return clean in generic_titles


def clean_bid_title(value: str, fallback: str) -> str:
    title = strip_inline_markup(value).strip()
    title = re.sub(r"[《》“”\"']", "", title)
    title = re.sub(r"\s+", "", title)
    title = re.sub(r"[（(][^）)]*(编号|招标|选聘|采购)[^）)]*[）)]", "", title)
    suffixes = (
        "招标文件",
        "选聘公告",
        "采购公告",
        "招标公告",
        "比选公告",
        "磋商文件",
        "谈判文件",
        "竞争性磋商",
        "公开招标",
        "公告",
        "文件",
    )
    changed = True
    while changed:
        changed = False
        for suffix in suffixes:
            if title.endswith(suffix):
                title = title[: -len(suffix)]
                changed = True
    for boundary in ("街道", "镇", "乡"):
        if boundary in title and len(title.rsplit(boundary, 1)[-1]) >= 4:
            title = title.rsplit(boundary, 1)[-1]
            break
    title = title.strip(" ，,。；;：:")
    if not title:
        title = strip_inline_markup(fallback).strip() or "投标文件"
    if "物业" in title and "物业服务" not in title:
        title = title.replace("物业", "物业服务", 1)
    if not any(marker in title for marker in ("技术标", "商务标", "资信标", "投标文件")):
        if "项目" not in title and any(marker in title for marker in ("服务", "物业", "工程", "采购", "选聘")):
            title += "项目"
        title += "技术标"
    return title


def infer_bid_title(markdown_text: str, fallback: str) -> str:
    candidates: list[tuple[int, str]] = []
    for line in markdown_text.splitlines():
        match = PROJECT_NAME_RE.search(line)
        if match:
            candidates.append((100, match.group(1)))
        heading = HEADING_RE.match(line)
        if heading:
            title = strip_inline_markup(heading.group(2)).strip()
            if title and not is_generic_document_heading(title):
                score = 90 if len(heading.group(1)) == 1 else 55
                candidates.append((score, title))
    for match in QUOTED_BID_RE.finditer(markdown_text[:20000]):
        candidates.append((80, match.group(1)))
    if fallback:
        candidates.append((30, fallback))
    if not candidates:
        return "投标文件"
    best = sorted(candidates, key=lambda item: (item[0], len(item[1])), reverse=True)[0][1]
    return clean_bid_title(best, fallback)


def extract_title(markdown_text: str, fallback: str) -> str:
    for line in markdown_text.splitlines():
        match = HEADING_RE.match(line)
        if match and len(match.group(1)) == 1:
            title = strip_inline_markup(match.group(2)).strip()
            if title and not is_generic_document_heading(title):
                return title
    return infer_bid_title(markdown_text, fallback)


def strip_inline_markup(text: str) -> str:
    text = re.sub(r"`([^`]+)`", r"\1", text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)
    text = re.sub(r"__([^_]+)__", r"\1", text)
    text = re.sub(r"\*([^*]+)\*", r"\1", text)
    text = re.sub(r"_([^_]+)_", r"\1", text)
    text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)
    return text.strip()


def split_table_row(line: str) -> list[str]:
    stripped = line.strip()
    if stripped.startswith("|"):
        stripped = stripped[1:]
    if stripped.endswith("|"):
        stripped = stripped[:-1]
    cells: list[str] = []
    current: list[str] = []
    escaped = False
    for char in stripped:
        if escaped:
            current.append(char)
            escaped = False
        elif char == "\\":
            escaped = True
        elif char == "|":
            cells.append("".join(current).strip())
            current = []
        else:
            current.append(char)
    cells.append("".join(current).strip())
    return cells


def is_table_separator(line: str) -> bool:
    cells = split_table_row(line)
    if len(cells) < 2:
        return False
    return all(re.fullmatch(r":?-{3,}:?", cell.strip()) for cell in cells)


def is_local_resource(uri: str) -> bool:
    lowered = uri.lower()
    return not (
        lowered.startswith("http://")
        or lowered.startswith("https://")
        or lowered.startswith("data:")
        or lowered.startswith("#")
        or lowered.startswith("mailto:")
    )


def relative_posix(path: Path, base: Path) -> str:
    try:
        rel = path.relative_to(base)
    except ValueError:
        rel = path
    return rel.as_posix()


def choose_reference_doc(style: str) -> Path | None:
    template_dir = PANDOC_ASSET_DIR / "templates"
    templates = sorted(template_dir.glob("*.docx"))
    if not templates:
        return None
    non_sci = [p for p in templates if "sci" not in p.name.lower()]
    return non_sci[0] if non_sci else templates[0]


def style_preset(name: str) -> dict[str, object]:
    if name not in STYLE_PRESETS:
        return STYLE_PRESETS["formal-blue"]
    return STYLE_PRESETS[name]
