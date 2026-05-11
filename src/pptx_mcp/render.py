"""Render a .pptx to per-slide PNGs via LibreOffice headless."""
from __future__ import annotations
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

SOFFICE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]


def find_soffice() -> str:
    env = os.environ.get("SOFFICE")
    if env and Path(env).exists():
        return env
    which = shutil.which("soffice")
    if which:
        return which
    for p in SOFFICE_CANDIDATES:
        if Path(p).exists():
            return p
    raise FileNotFoundError(
        "LibreOffice 'soffice' not found. Install it or set SOFFICE env var."
    )


def render_deck_to_pngs(pptx_path: str | Path, out_dir: str | Path) -> list[Path]:
    """Render every slide of pptx_path to PNG. Returns list of PNG paths in slide order.

    Strategy: pptx -> pdf via soffice, then pdf -> per-page png via soffice again.
    LibreOffice's direct pptx->png only emits the first slide, so we go through PDF.
    """
    pptx_path = Path(pptx_path).resolve()
    out_dir = Path(out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    soffice = find_soffice()

    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)
        # Step 1: pptx -> pdf
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf",
             "--outdir", str(td_path), str(pptx_path)],
            check=True, capture_output=True, timeout=120,
        )
        pdf_path = td_path / (pptx_path.stem + ".pdf")
        if not pdf_path.exists():
            raise RuntimeError(f"PDF render failed; no file at {pdf_path}")

        # Step 2: pdf -> per-page png. soffice supports --convert-to png on pdf,
        # but it also emits one page only. So we use pdftoppm if present,
        # otherwise fall back to pypdfium2.
        png_dir = td_path / "pngs"
        png_dir.mkdir()
        pngs = _pdf_to_pngs(pdf_path, png_dir)

        # Move into out_dir with stable names
        # First clear old slide_*.png so deletions are reflected
        for old in out_dir.glob("slide_*.png"):
            old.unlink()
        results: list[Path] = []
        for i, src in enumerate(pngs, start=1):
            dst = out_dir / f"slide_{i}.png"
            shutil.copyfile(src, dst)
            results.append(dst)
        return results


def _pdf_to_pngs(pdf: Path, out: Path) -> list[Path]:
    # Try pdftoppm (poppler) first
    if shutil.which("pdftoppm"):
        subprocess.run(
            ["pdftoppm", "-png", "-r", "144", str(pdf), str(out / "slide")],
            check=True, capture_output=True, timeout=120,
        )
        return sorted(out.glob("slide-*.png"))

    # Fallback: pypdfium2
    try:
        import pypdfium2 as pdfium
    except ImportError as e:
        raise RuntimeError(
            "Need pdftoppm (poppler) or pypdfium2 to convert PDF to PNGs. "
            "Install with: pip install pypdfium2"
        ) from e

    doc = pdfium.PdfDocument(str(pdf))
    try:
        pngs: list[Path] = []
        for i in range(len(doc)):
            page = doc[i]
            bitmap = page.render(scale=2.0)
            pil_image = bitmap.to_pil()
            p = out / f"slide-{i+1:03d}.png"
            pil_image.save(p)
            pngs.append(p)
        return pngs
    finally:
        doc.close()
