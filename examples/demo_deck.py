"""Smoke test: builds a 2-slide deck and renders it. No MCP, just the lib."""
from pathlib import Path
from pptx_mcp.deck import DeckManager
from pptx_mcp.render import render_deck_to_pngs

ROOT = Path(__file__).parent.parent / "decks"
M = DeckManager(ROOT)

DECK = "demo"
# Clean if re-running
import shutil
if (ROOT / DECK).exists():
    shutil.rmtree(ROOT / DECK)

M.create(DECK)

# Slide 0 — title
M.add_slide(DECK)
M.set_slide_background(DECK, 0, color="#0A0A23")
M.add_text_box(DECK, 0, x=1, y=2.8, w=11.333, h=1.5,
               text="pptx-mcp works", size_pt=60, bold=True,
               color="#FFFFFF", align="center")
M.add_text_box(DECK, 0, x=1, y=4.5, w=11.333, h=0.6,
               text="author -> render -> see -> iterate",
               size_pt=22, color="#A0A0FF", align="center", italic=True)

# Slide 1 — shapes
M.add_slide(DECK)
M.set_slide_background(DECK, 1, color="#FFFFFF")
M.add_text_box(DECK, 1, x=0.5, y=0.4, w=12.5, h=0.8,
               text="Shapes & layout", size_pt=32, bold=True, color="#0A0A23")
for i, (shape, fill) in enumerate([
    ("rectangle", "#FF6B6B"), ("ellipse", "#4ECDC4"),
    ("diamond", "#FFD166"), ("star", "#A78BFA"),
]):
    M.add_shape(DECK, 1, shape, x=0.7 + i * 3.1, y=2.0, w=2.5, h=2.5,
                fill=fill, text=shape)

M.save(DECK)
pngs = render_deck_to_pngs(M.get(DECK).path, M.get(DECK).render_dir)
print("Rendered:")
for p in pngs:
    print(" ", p)
