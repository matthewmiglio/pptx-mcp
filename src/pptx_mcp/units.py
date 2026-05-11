"""Unit conversion. All public APIs accept inches; internally pptx uses EMU."""
from pptx.util import Inches, Pt, Emu


def inches(v):
    return Inches(float(v))


def pt(v):
    return Pt(float(v))


def emu_to_inches(emu) -> float:
    return float(emu) / 914400.0


def parse_color(c):
    """Accept '#RRGGBB' or 'RRGGBB' or (r,g,b) tuple/list. Returns RGBColor."""
    from pptx.dml.color import RGBColor
    if c is None:
        return None
    if isinstance(c, (list, tuple)) and len(c) == 3:
        return RGBColor(int(c[0]), int(c[1]), int(c[2]))
    if isinstance(c, str):
        s = c.lstrip("#")
        if len(s) != 6:
            raise ValueError(f"bad color: {c!r}")
        return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
    raise ValueError(f"bad color: {c!r}")
