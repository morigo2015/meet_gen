from __future__ import annotations

from pathlib import Path
from typing import Any

from docxtpl import DocxTemplate

from .errors import RenderingError


def render_docx(
    template_path: str | Path,
    output_path: str | Path,
    context: dict[str, Any],
) -> Path:
    template = Path(template_path)
    target = Path(output_path)
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise RenderingError(f"Could not create output directory: {target.parent}") from exc

    try:
        document = DocxTemplate(template)
        document.render(context)
        document.save(target)
    except Exception as exc:
        raise RenderingError(f"Could not render document {target.name}: {exc}") from exc
    return target
