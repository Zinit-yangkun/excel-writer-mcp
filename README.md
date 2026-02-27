# excel-writer-mcp

MCP server for writing Excel files (.xlsx / .xlsm) based on [openpyxl](https://openpyxl.readthedocs.io/).

- `.xlsm` files are loaded with `keep_vba=True`, ensuring VBA macros are preserved after saving.
- `.xls` (legacy binary format) is **not supported**.

## Requirements

- Python >= 3.10
- [uv](https://docs.astral.sh/uv/) (recommended)

## Installation

```bash
git clone <repo-url>
cd excel-writer-mcp
uv venv && uv pip install -e .
```

## Usage

### stdio (for MCP clients)

```bash
uv run excel-writer-mcp
```

### MCP client configuration

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

## Tools

| Tool | Description |
|------|-------------|
| `create_workbook` | Create a new empty .xlsx workbook |
| `copy_file` | Safely copy any file (will not overwrite existing) |
| `get_workbook_info` | Get workbook metadata: sheets, dimensions, VBA status |
| `manage_sheets` | Create, delete, or rename a worksheet |
| `read_data` | Read data from a worksheet range (optionally include merged cell info) |
| `write_data` | Write a 2D array of data to a contiguous range |
| `write_cells` | Write to multiple specific cells by address (ideal for merged cell layouts) |
| `modify_rows_columns` | Insert or delete rows/columns |
| `merge_cells` | Merge or unmerge a range of cells |
| `format_cells` | Apply formatting: font, color, alignment, borders, number format, column width, row height |
| `create_chart` | Create a chart (bar, line, or pie) |

## .xlsm handling

- Cannot create `.xlsm` from scratch. Use `copy_file` to copy an existing `.xlsm` template.
- All read/write tools automatically detect `.xlsm` and load with `keep_vba=True`.
- No separate xlsm-specific tools needed — the same tools work for both `.xlsx` and `.xlsm`.

## License

MIT
