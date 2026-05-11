"""MCP server entry point. Registers tools that wrap DeckManager + render."""
from __future__ import annotations
import asyncio
import json
import os
from pathlib import Path

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

from .deck import DeckManager, SHAPE_TYPES
from .render import render_deck_to_pngs, find_soffice

ROOT = Path(os.environ.get("PPTX_MCP_ROOT", Path.home() / "pptx-mcp-decks")).resolve()
ROOT.mkdir(parents=True, exist_ok=True)
MANAGER = DeckManager(ROOT)

server = Server("pptx-mcp")


def _ok(data) -> list[TextContent]:
    return [TextContent(type="text", text=json.dumps(data, default=str, indent=2))]


def _err(msg: str) -> list[TextContent]:
    return [TextContent(type="text", text=json.dumps({"error": msg}, indent=2))]


# ---- Tool definitions (schemas) ----
TOOLS: list[Tool] = [
    Tool(name="list_decks", description="List all decks under the root directory.",
         inputSchema={"type": "object", "properties": {}}),
    Tool(name="create_deck", description="Create a new deck. Slide size in inches; default 13.333x7.5 (16:9).",
         inputSchema={"type": "object",
                      "properties": {"deck_id": {"type": "string"},
                                     "width_in": {"type": "number", "default": 13.333},
                                     "height_in": {"type": "number", "default": 7.5}},
                      "required": ["deck_id"]}),
    Tool(name="open_deck", description="Open an existing deck by id.",
         inputSchema={"type": "object", "properties": {"deck_id": {"type": "string"}}, "required": ["deck_id"]}),
    Tool(name="save_deck", description="Save deck to disk. Required before render.",
         inputSchema={"type": "object", "properties": {"deck_id": {"type": "string"}}, "required": ["deck_id"]}),
    Tool(name="close_deck", description="Drop the in-memory handle (does not delete files).",
         inputSchema={"type": "object", "properties": {"deck_id": {"type": "string"}}, "required": ["deck_id"]}),
    Tool(name="get_deck_info", description="Return slide count, size, and path for a deck.",
         inputSchema={"type": "object", "properties": {"deck_id": {"type": "string"}}, "required": ["deck_id"]}),
    Tool(name="list_slides", description="Return [{index, element_count}] for the deck.",
         inputSchema={"type": "object", "properties": {"deck_id": {"type": "string"}}, "required": ["deck_id"]}),
    Tool(name="get_slide_info", description="Detailed dump of every element on a slide (id, type, position, text).",
         inputSchema={"type": "object",
                      "properties": {"deck_id": {"type": "string"}, "slide_index": {"type": "integer"}},
                      "required": ["deck_id", "slide_index"]}),
    Tool(name="get_text_style",
         description=("Return text styling (font, size, color, bold/italic, position) for shapes "
                      "on a slide. Resolves placeholder inheritance from the slide layout and "
                      "surfaces layout label text like 'Presenter:'. Pass element_id to target one "
                      "shape, or search_text (e.g. 'Presenter: James Bruin') to match by content."),
         inputSchema={"type": "object",
                      "properties": {"deck_id": {"type": "string"},
                                     "slide_index": {"type": "integer"},
                                     "element_id": {"type": "integer"},
                                     "search_text": {"type": "string"}},
                      "required": ["deck_id", "slide_index"]}),
    Tool(name="add_slide", description="Append a slide. layout=6 is blank.",
         inputSchema={"type": "object",
                      "properties": {"deck_id": {"type": "string"}, "layout": {"type": "integer", "default": 6}},
                      "required": ["deck_id"]}),
    Tool(name="delete_slide", description="Delete a slide by index.",
         inputSchema={"type": "object",
                      "properties": {"deck_id": {"type": "string"}, "slide_index": {"type": "integer"}},
                      "required": ["deck_id", "slide_index"]}),
    Tool(name="reorder_slide", description="Move a slide to a new index.",
         inputSchema={"type": "object",
                      "properties": {"deck_id": {"type": "string"}, "from_index": {"type": "integer"},
                                     "to_index": {"type": "integer"}},
                      "required": ["deck_id", "from_index", "to_index"]}),
    Tool(name="set_slide_background",
         description="Set slide background color (#RRGGBB) or full-bleed image (image_path).",
         inputSchema={"type": "object",
                      "properties": {"deck_id": {"type": "string"}, "slide_index": {"type": "integer"},
                                     "color": {"type": "string"}, "image_path": {"type": "string"}},
                      "required": ["deck_id", "slide_index"]}),
    Tool(name="add_text_box",
         description="Add a text box. All positions in inches. Returns element_id.",
         inputSchema={"type": "object",
                      "properties": {"deck_id": {"type": "string"}, "slide_index": {"type": "integer"},
                                     "x": {"type": "number"}, "y": {"type": "number"},
                                     "w": {"type": "number"}, "h": {"type": "number"},
                                     "text": {"type": "string"},
                                     "font": {"type": "string"}, "size_pt": {"type": "number"},
                                     "bold": {"type": "boolean"}, "italic": {"type": "boolean"},
                                     "color": {"type": "string"},
                                     "align": {"type": "string", "enum": ["left", "center", "right", "justify"]},
                                     "valign": {"type": "string", "enum": ["top", "middle", "bottom"]},
                                     "auto_size": {"type": "boolean"}},
                      "required": ["deck_id", "slide_index", "x", "y", "w", "h"]}),
    Tool(name="add_image",
         description="Add an image. Omit w/h to keep native aspect. Returns element_id.",
         inputSchema={"type": "object",
                      "properties": {"deck_id": {"type": "string"}, "slide_index": {"type": "integer"},
                                     "image_path": {"type": "string"},
                                     "x": {"type": "number"}, "y": {"type": "number"},
                                     "w": {"type": "number"}, "h": {"type": "number"}},
                      "required": ["deck_id", "slide_index", "image_path", "x", "y"]}),
    Tool(name="add_shape",
         description=f"Add a shape. shape_type one of: {sorted(SHAPE_TYPES)}. Returns element_id.",
         inputSchema={"type": "object",
                      "properties": {"deck_id": {"type": "string"}, "slide_index": {"type": "integer"},
                                     "shape_type": {"type": "string"},
                                     "x": {"type": "number"}, "y": {"type": "number"},
                                     "w": {"type": "number"}, "h": {"type": "number"},
                                     "fill": {"type": "string"}, "line_color": {"type": "string"},
                                     "line_width_pt": {"type": "number"}, "text": {"type": "string"}},
                      "required": ["deck_id", "slide_index", "shape_type", "x", "y", "w", "h"]}),
    Tool(name="set_text",
         description="Update text and/or styling of an existing element.",
         inputSchema={"type": "object",
                      "properties": {"deck_id": {"type": "string"}, "slide_index": {"type": "integer"},
                                     "element_id": {"type": "integer"},
                                     "text": {"type": "string"}, "font": {"type": "string"},
                                     "size_pt": {"type": "number"}, "bold": {"type": "boolean"},
                                     "italic": {"type": "boolean"}, "underline": {"type": "boolean"},
                                     "color": {"type": "string"},
                                     "align": {"type": "string", "enum": ["left", "center", "right", "justify"]},
                                     "valign": {"type": "string", "enum": ["top", "middle", "bottom"]}},
                      "required": ["deck_id", "slide_index", "element_id"]}),
    Tool(name="set_position", description="Move an element. x,y in inches.",
         inputSchema={"type": "object",
                      "properties": {"deck_id": {"type": "string"}, "slide_index": {"type": "integer"},
                                     "element_id": {"type": "integer"},
                                     "x": {"type": "number"}, "y": {"type": "number"}},
                      "required": ["deck_id", "slide_index", "element_id"]}),
    Tool(name="set_size", description="Resize an element. w,h in inches.",
         inputSchema={"type": "object",
                      "properties": {"deck_id": {"type": "string"}, "slide_index": {"type": "integer"},
                                     "element_id": {"type": "integer"},
                                     "w": {"type": "number"}, "h": {"type": "number"}},
                      "required": ["deck_id", "slide_index", "element_id"]}),
    Tool(name="set_rotation", description="Rotate element by degrees (clockwise).",
         inputSchema={"type": "object",
                      "properties": {"deck_id": {"type": "string"}, "slide_index": {"type": "integer"},
                                     "element_id": {"type": "integer"},
                                     "degrees": {"type": "number"}},
                      "required": ["deck_id", "slide_index", "element_id", "degrees"]}),
    Tool(name="set_z_order",
         description="Reorder element. action: front|back|forward|backward.",
         inputSchema={"type": "object",
                      "properties": {"deck_id": {"type": "string"}, "slide_index": {"type": "integer"},
                                     "element_id": {"type": "integer"},
                                     "action": {"type": "string",
                                                "enum": ["front", "back", "forward", "backward"]}},
                      "required": ["deck_id", "slide_index", "element_id", "action"]}),
    Tool(name="delete_element", description="Remove an element from a slide.",
         inputSchema={"type": "object",
                      "properties": {"deck_id": {"type": "string"}, "slide_index": {"type": "integer"},
                                     "element_id": {"type": "integer"}},
                      "required": ["deck_id", "slide_index", "element_id"]}),
    Tool(name="render_slide",
         description="Save deck, render it, return the PNG path for ONE slide. Open with the Read tool.",
         inputSchema={"type": "object",
                      "properties": {"deck_id": {"type": "string"}, "slide_index": {"type": "integer"}},
                      "required": ["deck_id", "slide_index"]}),
    Tool(name="render_all",
         description="Save deck, render all slides, return list of PNG paths.",
         inputSchema={"type": "object", "properties": {"deck_id": {"type": "string"}},
                      "required": ["deck_id"]}),
    Tool(name="check_environment",
         description="Diagnostic: where is soffice, where is the deck root, etc.",
         inputSchema={"type": "object", "properties": {}}),
]


@server.list_tools()
async def list_tools() -> list[Tool]:
    return TOOLS


@server.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    try:
        return await asyncio.to_thread(_dispatch, name, arguments or {})
    except Exception as e:
        return _err(f"{type(e).__name__}: {e}")


def _dispatch(name: str, a: dict) -> list[TextContent]:
    M = MANAGER
    if name == "list_decks":
        return _ok(M.list_decks())
    if name == "create_deck":
        h = M.create(a["deck_id"], a.get("width_in", 13.333), a.get("height_in", 7.5))
        return _ok({"deck_id": h.deck_id, "path": str(h.path)})
    if name == "open_deck":
        h = M.open(a["deck_id"])
        return _ok(M.get_deck_info(h.deck_id))
    if name == "save_deck":
        return _ok({"saved_to": M.save(a["deck_id"])})
    if name == "close_deck":
        M.close(a["deck_id"])
        return _ok({"closed": a["deck_id"]})
    if name == "get_deck_info":
        return _ok(M.get_deck_info(a["deck_id"]))
    if name == "list_slides":
        return _ok(M.list_slides(a["deck_id"]))
    if name == "get_slide_info":
        return _ok(M.get_slide_info(a["deck_id"], a["slide_index"]))
    if name == "get_text_style":
        return _ok(M.get_text_style(a["deck_id"], a["slide_index"],
                                    element_id=a.get("element_id"),
                                    search_text=a.get("search_text")))

    if name == "add_slide":
        idx = M.add_slide(a["deck_id"], a.get("layout", 6))
        return _ok({"slide_index": idx})
    if name == "delete_slide":
        M.delete_slide(a["deck_id"], a["slide_index"])
        return _ok({"deleted": a["slide_index"]})
    if name == "reorder_slide":
        M.reorder_slide(a["deck_id"], a["from_index"], a["to_index"])
        return _ok({"ok": True})
    if name == "set_slide_background":
        M.set_slide_background(a["deck_id"], a["slide_index"],
                               color=a.get("color"), image_path=a.get("image_path"))
        return _ok({"ok": True})

    if name == "add_text_box":
        eid = M.add_text_box(
            a["deck_id"], a["slide_index"], a["x"], a["y"], a["w"], a["h"],
            text=a.get("text", ""), font=a.get("font"), size_pt=a.get("size_pt"),
            bold=a.get("bold"), italic=a.get("italic"), color=a.get("color"),
            align=a.get("align"), valign=a.get("valign"), auto_size=a.get("auto_size", False),
        )
        return _ok({"element_id": eid})
    if name == "add_image":
        eid = M.add_image(a["deck_id"], a["slide_index"], a["image_path"],
                          a["x"], a["y"], a.get("w"), a.get("h"))
        return _ok({"element_id": eid})
    if name == "add_shape":
        eid = M.add_shape(a["deck_id"], a["slide_index"], a["shape_type"],
                          a["x"], a["y"], a["w"], a["h"],
                          fill=a.get("fill"), line_color=a.get("line_color"),
                          line_width_pt=a.get("line_width_pt"), text=a.get("text"))
        return _ok({"element_id": eid})

    if name == "set_text":
        M.set_text(a["deck_id"], a["slide_index"], a["element_id"],
                   text=a.get("text"), font=a.get("font"), size_pt=a.get("size_pt"),
                   bold=a.get("bold"), italic=a.get("italic"), underline=a.get("underline"),
                   color=a.get("color"), align=a.get("align"), valign=a.get("valign"))
        return _ok({"ok": True})
    if name == "set_position":
        M.set_position(a["deck_id"], a["slide_index"], a["element_id"],
                       x=a.get("x"), y=a.get("y"))
        return _ok({"ok": True})
    if name == "set_size":
        M.set_size(a["deck_id"], a["slide_index"], a["element_id"],
                   w=a.get("w"), h=a.get("h"))
        return _ok({"ok": True})
    if name == "set_rotation":
        M.set_rotation(a["deck_id"], a["slide_index"], a["element_id"], a["degrees"])
        return _ok({"ok": True})
    if name == "set_z_order":
        M.set_z_order(a["deck_id"], a["slide_index"], a["element_id"], a["action"])
        return _ok({"ok": True})
    if name == "delete_element":
        M.delete_element(a["deck_id"], a["slide_index"], a["element_id"])
        return _ok({"ok": True})

    if name == "render_slide":
        M.save(a["deck_id"])
        h = M.get(a["deck_id"])
        pngs = render_deck_to_pngs(h.path, h.render_dir)
        idx = a["slide_index"]
        if not (0 <= idx < len(pngs)):
            return _err(f"slide_index {idx} out of range; deck has {len(pngs)} slides")
        return _ok({"png_path": str(pngs[idx]), "slide_index": idx})
    if name == "render_all":
        M.save(a["deck_id"])
        h = M.get(a["deck_id"])
        pngs = render_deck_to_pngs(h.path, h.render_dir)
        return _ok({"png_paths": [str(p) for p in pngs]})
    if name == "check_environment":
        try:
            soffice = find_soffice()
        except Exception as e:
            soffice = f"NOT FOUND: {e}"
        return _ok({"soffice": soffice, "deck_root": str(ROOT),
                    "open_decks": list(MANAGER._open.keys())})

    return _err(f"unknown tool: {name}")


async def _amain():
    async with stdio_server() as (read, write):
        await server.run(read, write, server.create_initialization_options())


def main():
    asyncio.run(_amain())


if __name__ == "__main__":
    main()
