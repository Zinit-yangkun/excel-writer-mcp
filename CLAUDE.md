# excel-writer-mcp

MCP server for writing Excel files (.xlsx and .xlsm) using openpyxl.

## Key Design Decisions

- **xlsm support**: All xlsm operations use `keep_vba=True` when loading to preserve VBA macros
- **Framework**: FastMCP for automatic tool registration and stdio transport
- **No new xlsm creation**: Cannot create .xlsm from scratch (openpyxl limitation). Must copy an existing template.

## Project Structure

```
src/excel_writer_mcp/
  __init__.py
  __main__.py      # Entry point (stdio transport)
  server.py        # FastMCP server with all tools
```

## Running

```bash
# stdio mode (for MCP clients)
python -m excel_writer_mcp

# or via entry point
excel-writer-mcp
```

## MCP Client Configuration

```json
{
  "mcpServers": {
    "excel-writer-mcp": {
      "command": "uv",
      "args": [
        "--directory",
        "/path/to/excel-writer-mcp",
        "run",
        "excel-writer-mcp"
      ]
    }
  }
}
```
