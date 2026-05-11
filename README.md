# pptx-mcp

An MCP server that lets an LLM author, render, and iterate on PowerPoint decks. The LLM creates slides, renders them to PNG, **sees** the result, and adjusts — closing the visual feedback loop.

## Tech stack

```
   ┌──────────┐   MCP    ┌─────────────┐   edit    ┌─────────────┐
   │  Claude  │ ───────▶ │  pptx-mcp   │ ────────▶ │ python-pptx │ ──▶ .pptx
   └──────────┘          │  (server)   │           └─────────────┘
        ▲                │             │   render  ┌─────────────┐
        │   PNG path     │             │ ────────▶ │ LibreOffice │ ──▶ .png
        └────────────────┤             │           │  + pypdfium │
                         └─────────────┘           └─────────────┘
```

## Usage

Install:

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m pip install -e .
```

Register with Claude Code:

```powershell
claude mcp add pptx -s user -- "C:\My_Files\my_programs\pptx-mcp\.venv\Scripts\python.exe" -m pptx_mcp.server
```

Optional env: `PPTX_MCP_ROOT` (deck directory), `SOFFICE` (path to `soffice.exe`).

Typical loop:

```
create_deck → add_slide → add_text_box / add_image / add_shape
            → render_slide → Read the PNG → set_position / set_text → re-render
```

Full tool reference: see `src/` or call `check_environment` from the MCP client.
