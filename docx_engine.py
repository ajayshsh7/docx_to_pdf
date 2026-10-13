
from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
import zipfile

from pathlib import Path
from typing import Any

import pandas as pd
from docx import Document


PLACEHOLDER_RE = re.compile(r"\{\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*\}\}")


def clean_value(value: Any) -> str:
    """Convert Excel values to printable text."""
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def safe_filename(value: str) -> str:
    """Remove unsafe filename characters."""
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value)
    value = re.sub(r"\s+", "_", value).strip("._ ")
    return value[:100] or "document"



def replace_paragraph(paragraph, replacements: dict[str, str]) -> None:
    """Replace placeholders while preserving their character formatting."""
    if not paragraph.runs:
        return

    runs = paragraph.runs
    original = "".join(run.text for run in runs)

    matches = list(PLACEHOLDER_RE.finditer(original))
    if not matches:
        return

    # Record each run's character range in the paragraph.
    offsets = []
    position = 0
    for run in runs:
        start = position
        position += len(run.text)
        offsets.append((start, position))

    # Process backwards so earlier character positions remain valid.
    for match in reversed(matches):
        key = match.group(1)
        if key not in replacements:
            continue

        start, end = match.span()
        start_idx = next(
            (i for i, (a, b) in enumerate(offsets)
             if a <= start < b or (start == b and a == b)),
            None,
        )
        end_idx = next(
            (i for i, (a, b) in enumerate(offsets)
             if a < end <= b),
            None,
        )

        if start_idx is None or end_idx is None:
            continue

        start_offset = start - offsets[start_idx][0]
        end_offset = end - offsets[end_idx][0]

        if start_idx == end_idx:
            text = runs[start_idx].text
            runs[start_idx].text = (
                text[:start_offset]
                + replacements[key]
                + text[end_offset:]
            )
        else:
            first_text = runs[start_idx].text
            last_text = runs[end_idx].text

            runs[start_idx].text = (
                first_text[:start_offset] + replacements[key]
            )
            for i in range(start_idx + 1, end_idx):
                runs[i].text = ""
            runs[end_idx].text = last_text[end_offset:]

    """Replace tokens even if Word split them across runs.

    Formatting outside the replaced text is retained. When a token
    spans multiple runs, the replacement inherits the first run's style.
    """
    if not paragraph.runs:
        return

    original = "".join(run.text for run in paragraph.runs)
    updated = PLACEHOLDER_RE.sub(
        lambda match: replacements.get(match.group(1), match.group(0)),
        original,
    )

    if updated == original:
        return

    # Keep the paragraph's first run and its character formatting.
    paragraph.runs[0].text = updated
    for run in paragraph.runs[1:]:
        run.text = ""


def replace_in_table(table, replacements: dict[str, str]) -> None:
    for row in table.rows:
        for cell in row.cells:
            replace_in_container(cell, replacements)


def replace_in_container(container, replacements: dict[str, str]) -> None:
    for paragraph in container.paragraphs:
        replace_paragraph(paragraph, replacements)

    for table in container.tables:
        replace_in_table(table, replacements)


def fill_docx(
    template_path: Path,
    output_path: Path,
    replacements: dict[str, str],
) -> None:
    doc = Document(str(template_path))

    replace_in_container(doc, replacements)

    for section in doc.sections:
        for part in (
            section.header,
            section.footer,
            section.first_page_header,
            section.first_page_footer,
            section.even_page_header,
            section.even_page_footer,
        ):
            replace_in_container(part, replacements)

    doc.save(str(output_path))


def find_placeholders(template_path: Path) -> set[str]:
    """Discover tokens in the supported document areas."""
    doc = Document(str(template_path))
    found: set[str] = set()

    def scan(container):
        for paragraph in container.paragraphs:
            found.update(PLACEHOLDER_RE.findall(paragraph.text))
        for table in container.tables:
            for row in table.rows:
                for cell in row.cells:
                    scan(cell)

    scan(doc)

    for section in doc.sections:
        for part in (
            section.header, section.footer,
            section.first_page_header, section.first_page_footer,
            section.even_page_header, section.even_page_footer,
        ):
            scan(part)

    return found


def find_soffice() -> str:
    """Locate LibreOffice's command-line executable."""
    executable = shutil.which("soffice") or shutil.which("libreoffice")
    if executable:
        return executable

    # Common Windows installation paths.
    candidates = [
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    ]
    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)

    raise RuntimeError(
        "LibreOffice was not found. Install it and ensure "
        "'soffice' is on PATH."
    )


def convert_to_pdf(docx_path: Path, pdf_path: Path) -> None:
    """Convert one DOCX to PDF in an isolated temporary profile."""
    soffice = find_soffice()
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="lo_profile_") as profile:
        profile_uri = Path(profile).resolve().as_uri()

        result = subprocess.run(
            [
                soffice,
                f"-env:UserInstallation={profile_uri}",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", str(pdf_path.parent.resolve()),
                str(docx_path.resolve()),
            ],
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )

        expected_pdf = pdf_path.parent / f"{docx_path.stem}.pdf"

        if result.returncode != 0 or not expected_pdf.is_file():
            raise RuntimeError(
                f"PDF conversion failed for {docx_path.name}: "
                f"{result.stderr or result.stdout}"
            )

        if expected_pdf.resolve() != pdf_path.resolve():
            expected_pdf.replace(pdf_path)


def generate_batch(
    template_path: Path,
    dataframe: pd.DataFrame,
    output_dir: Path,
    name_column: str,
    id_column: str,
    progress_callback=None,
) -> tuple[list[Path], list[dict[str, str]]]:
    """Generate one PDF per valid Excel row."""
    output_dir.mkdir(parents=True, exist_ok=True)
    generated: list[Path] = []
    errors: list[dict[str, str]] = []

    with tempfile.TemporaryDirectory(prefix="docx_batch_") as tmp:
        temp_dir = Path(tmp)

        for index, (_, row) in enumerate(dataframe.iterrows(), start=1):
            replacements = {
                str(column): clean_value(value)
                for column, value in row.items()
            }

            name = replacements.pop(name_column, "")
            person_id = replacements.pop(id_column, "")

            if not name or not person_id:
                errors.append({
                    "row": str(index + 1),
                    "name": name,
                    "error": "Name or ID is blank",
                })
                if progress_callback:
                    progress_callback(index, len(dataframe))
                continue

            stem = safe_filename(person_id)
            docx_path = temp_dir / f"record_{index}.docx"
            pdf_path = output_dir / f"{stem}.pdf"

            # Avoid overwriting PDFs if duplicate IDs exist.
            if pdf_path.exists():
                pdf_path = output_dir / f"{stem}_{index}.pdf"

            try:
                fill_docx(template_path, docx_path, replacements)

                remaining = find_placeholders(docx_path)
                if remaining:
                    raise ValueError(
                        "Unreplaced placeholders: "
                        + ", ".join(sorted(remaining))
                    )

                convert_to_pdf(docx_path, pdf_path)
                generated.append(pdf_path)

            except Exception as exc:
                errors.append({
                    "row": str(index + 1),
                    "name": name,
                    "error": str(exc),
                })

            if progress_callback:
                progress_callback(index, len(dataframe))

    return generated, errors


def create_zip(pdf_paths: list[Path], zip_path: Path) -> None:
    with zipfile.ZipFile(
        zip_path, "w", compression=zipfile.ZIP_DEFLATED
    ) as archive:
        for pdf_path in pdf_paths:
            archive.write(pdf_path, arcname=pdf_path.name)
