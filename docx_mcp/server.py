from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from docx_mcp.tools.read_tools import register_read_tools
from docx_mcp.tools.style_tools import register_style_tools
from docx_mcp.tools.write_tools import register_write_tools


def create_server() -> FastMCP:
    mcp = FastMCP("docx-mcp")
    register_read_tools(mcp)
    register_write_tools(mcp)
    register_style_tools(mcp)
    return mcp


def run_server(transport: str = "stdio") -> None:
    mcp = create_server()
    mcp.run(transport=transport)
