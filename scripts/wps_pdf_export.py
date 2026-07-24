from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path


def prepare_pywin32_paths(extra_root: Path | None = None) -> None:
    candidates = []
    if extra_root:
        candidates.append(extra_root)
    env_root = os.environ.get("PYWIN32_TARGET")
    if env_root:
        candidates.append(Path(env_root))
    candidates.append(Path.cwd() / ".tmp_pydeps")
    for root in candidates:
        if not root.is_dir():
            continue
        system32 = root / "pywin32_system32"
        if system32.is_dir():
            try:
                os.add_dll_directory(str(system32))
            except Exception:
                pass
            os.environ["PATH"] = str(system32) + os.pathsep + os.environ.get("PATH", "")
        for extra in (root, root / "win32", root / "win32" / "lib", root / "win32com"):
            if extra.is_dir():
                sys.path.insert(0, str(extra))


def export_pdf(input_path: Path, output_path: Path, pywin32_root: Path | None = None) -> dict:
    prepare_pywin32_paths(pywin32_root)
    try:
        import pythoncom
        import win32com.client
    except ModuleNotFoundError as exc:
        return {"status": "failed", "reason": f"pywin32 unavailable: {exc}"}

    pythoncom.CoInitialize()
    app = None
    doc = None
    prog_errors = []
    try:
        for prog_id in ("KWPS.Application", "WPS.Application", "Word.Application"):
            try:
                app = win32com.client.DispatchEx(prog_id)
                app.Visible = False
                active_prog = prog_id
                break
            except Exception as exc:  # noqa: BLE001
                prog_errors.append(f"{prog_id}: {exc}")
        else:
            return {"status": "failed", "reason": "; ".join(prog_errors)}

        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc = app.Documents.Open(str(input_path))
        try:
            doc.ExportAsFixedFormat(str(output_path), 17)
        except Exception:
            doc.SaveAs2(str(output_path), FileFormat=17)
        if output_path.is_file() and output_path.stat().st_size > 1000:
            return {"status": "exported", "program": active_prog, "pdf": str(output_path), "size_bytes": output_path.stat().st_size}
        return {"status": "failed", "program": active_prog, "reason": "PDF was not created or is too small."}
    except Exception as exc:  # noqa: BLE001
        return {"status": "failed", "reason": str(exc)}
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def main() -> int:
    parser = argparse.ArgumentParser(description="Export DOCX to PDF through local WPS/Word COM.")
    parser.add_argument("input", type=Path)
    parser.add_argument("--output", "-o", type=Path, required=True)
    parser.add_argument("--pywin32-root", type=Path, default=None)
    args = parser.parse_args()
    result = export_pdf(args.input.resolve(), args.output.resolve(), args.pywin32_root)
    print(result)
    return 0 if result.get("status") == "exported" else 2


if __name__ == "__main__":
    raise SystemExit(main())
