## What is Spire.XLS MCP Server?

The Spire.XLS MCP Server is a robust solution that empowers AI agents to work with Excel files using the Model Context Protocol (MCP). It is totally independent and doesn't require Microsoft Office to be installed on system. This tool enables AI agents to [create](https://www.e-iceblue.com/Tutorials/Python/Spire.XLS-for-Python/Program-Guide/Document-Operation/Python-Create-Read-or-Update-Excel-Documents.html), read, [edit](https://www.e-iceblue.com/Tutorials/Python/Spire.XLS-for-Python/Program-Guide/Document-Operation/Python-Edit-Excel-Documents.html), and [convert Excel workbooks](https://www.e-iceblue.com/Tutorials/Python/Spire.XLS-for-Python/Program-Guide/Conversion/Python-Convert-Excel-to-PDF.html) seamlessly

## Main Features: 

- Convert [Excel to PDF](https://www.e-iceblue.com/Tutorials/Python/Spire.XLS-for-Python/Program-Guide/Conversion/Python-Convert-Excel-to-PDF.html), [Excel to HTML](https://www.e-iceblue.com/Tutorials/Python/Spire.XLS-for-Python/Program-Guide/Conversion/Python-Convert-Excel-to-HTML-and-Vice-Versa.html), [Excel to CSV](https://www.e-iceblue.com/Tutorials/Python/Spire.XLS-for-Python/Program-Guide/Conversion/Python-Convert-Excel-to-CSV-and-Vice-Versa.html), Excel to image, Excel to XML, and more with high fidelity.
- Create, modify, and manage Excel workbooks
- Manage and control worksheets: [rename](https://www.e-iceblue.com/Tutorials/Python/Spire.XLS-for-Python/Program-Guide/Worksheet/Python-Change-Worksheet-Names-and-Set-Tab-Colors-in-Excel.html), move, hide, [freeze panes](https://www.e-iceblue.com/Tutorials/Python/Spire.XLS-for-Python/Program-Guide/Cells/Python-Freeze-or-Unfreeze-Panes-in-Excel.html), and more.
- Manage worksheets and cell ranges
- [Read and write data](https://www.e-iceblue.com/Tutorials/Python/Spire.XLS-for-Python/Program-Guide/Document-Operation/Python-Create-Read-or-Update-Excel-Documents.html)
- Analyze Excel data
- Add various chart types to create visual Excel dashboards from data
- [Create and manipulate pivot tables](https://www.e-iceblue.com/Tutorials/Python/Spire.XLS-for-Python/Program-Guide/Pivot-Table/Python-Create-or-Operate-Pivot-Tables-in-Excel.html) to summarize, analyze, explore, and present Excel data.

## How to use Spire.XLS MCP Server?

### Prerequisites

- Python 3.10 or higher

### Installation

1. Clone the repository:
```bash
git clone https://github.com/eiceblue/spire-xls-mcp-server.git
cd spire-xls-mcp-server
```

2. Install using uv:
```bash
uv pip install -e .
```
### Running the Server

Start the server (default port 8000):
```bash
uv run spire-xls-mcp-server
```

Custom port (e.g., 8080):

```bash
# Bash/Linux/macOS
export FASTMCP_PORT=8080 && uv run spire-xls-mcp-server

# Windows PowerShell
$env:FASTMCP_PORT = "8080"; uv run spire-xls-mcp-server
```

## Integration with AI Tools

### Cursor IDE

1. Add this configuration to Cursor:
```json
{
  "mcpServers": {
    "excel": {
      "url": "http://localhost:8000/sse",
      "env": {
        "EXCEL_FILES_PATH": "/path/to/excel/files"
      }
    }
  }
}
```
2. The Excel tools will be available through your AI assistant.

### Remote Hosting & Transport Protocols

This server uses Server-Sent Events (SSE) transport protocol. For different use cases:

1. **Using with Claude Desktop (requires stdio):**
   - Use [Supergateway](https://github.com/supercorp-ai/supergateway) to convert SSE to stdio

2. **Hosting Your MCP Server:**
   - [Remote MCP Server Guide](https://developers.cloudflare.com/agents/guides/remote-mcp-server/)

## Environment Variables

| Variable | Description | Default |
|--------|------|--------|
| `FASTMCP_PORT` | Server port | `8000` |
| `EXCEL_FILES_PATH` | Directory for Excel files | `./excel_files` |

## Available Tools

The server provides a comprehensive set of Excel manipulation tools. Here are the main categories:

- **Basic Operations**: Create, read, write, and [delete Excel worksheets](https://www.e-iceblue.com/Tutorials/Python/Spire.XLS-for-Python/Program-Guide/Worksheet/Python-Move-or-Delete-Worksheets-in-Excel.html) or workbooks.
- **Data Processing**: Read and write cell data, [apply formulas](https://www.e-iceblue.com/Tutorials/Python/Spire.XLS-for-Python/Program-Guide/Formula/Python-Add-or-Read-Formulas-in-Excel.html), sort and filter
- **Formatting**: Apply styles, [merge cells](https://www.e-iceblue.com/Tutorials/Python/Spire.XLS-for-Python/Program-Guide/Cells/Python-Merge-or-Unmerge-Cells-in-Excel.html), set fonts and colors
- **Advanced Features**: [Create charts](https://www.e-iceblue.com/Tutorials/Python/Spire.XLS-for-Python/Program-Guide/Chart/Python-Create-Column-Charts-in-Excel.html), pivot tables, [conditional formatting](https://www.e-iceblue.com/Tutorials/Python/Spire.XLS-for-Python/Program-Guide/Conditional-Formatting/Python-Apply-Conditional-Formatting-in-Excel.html)
- **Conversion**: Convert Excel to PDF, HTML, CSV, image, XML, and more with high fidelity.

See [TOOLS.md](https://github.com/eiceblue/spire-xls-mcp-server/blob/main/TOOLS.md) for complete documentation of all available tools.

## FAQ from Spire.XLS MCP Server?

Q1. Can I use Spire.XLS MCP Server for any directory?

Yes, Spire.XLS MCP Serer works for any directory.

Q2. Is Spire.XLS MCP Server free to use?

Yes, it is licensed under the MIT License, allowing free use and modification.

Q3. What programming languages does Spire.XLS MCP Server support?

It is built with Python.

## License
MIT
