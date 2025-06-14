# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository Overview

This repository contains Model Context Protocol (MCP) servers for Microsoft Office automation through AI assistance. The servers enable Claude to interact with PowerPoint and Excel applications via COM automation using `pywin32`.

**Key Components:**
- `mcp_powerpoint_server_win32.py`: PowerPoint automation server with presentation management, slide operations, and financial chart generation
- `mcp_excel_server_win32.py`: Excel automation server for workbook, worksheet, and cell operations
- `mcp_powerpoint_server.py`: Alternative PowerPoint server implementation

## Architecture

**COM Automation Pattern:**
- All servers use `win32com.client` for Office application interaction
- COM initialization with `pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)`
- Connection strategy: attempt to connect to running instance, launch new if needed
- Error handling for `pywintypes.com_error` exceptions

**MCP Integration:**
- Built on `mcp.server.fastmcp.FastMCP` framework
- Tools expose Office functionality as structured API endpoints
- Type annotations using `typing` module for parameter validation

**Data Flow:**
1. Claude sends commands through MCP protocol
2. Server translates to COM automation calls
3. Office applications execute operations
4. Results returned through MCP response

## Development Commands

**Installation:**
```bash
# Install Python dependencies
uv pip install pywin32
pip install -r requirements.txt

# Required post-install for COM automation (run as administrator)
python C:\path\to\your\env\Scripts\pywin32_postinstall.py -install
```

**Running Servers:**
```bash
# PowerPoint server
uv run mcp_powerpoint_server_win32.py

# Excel server  
uv run mcp_excel_server_win32.py
```

## Platform Requirements

- **Windows only** - COM automation requires Windows OS
- **Microsoft Office installed** - PowerPoint and/or Excel applications
- **Python 3.7+** with pywin32 package
- **Running Office instances** - servers connect to active applications

## Key Implementation Details

**PowerPoint Server:**
- Manages presentations through `PowerPoint.Application` COM object
- Shape operations use MSO constants (msoShapeRectangle, etc.)
- Placeholder management with ppPlaceholder constants
- Financial chart generation with dummy data (extensible for real APIs)
- Template system for slide creation

**Excel Server:**
- Workbook management through `Excel.Application` COM object
- Range operations with Excel constants (xlUp, xlDown, etc.)
- Type conversion for dates, currency, and numeric values
- Multiple save formats supported (.xlsx, .xlsm, .xlsb, .xls)
- **DCF Analysis Features:**
  - `find_used_ranges()`: Detect contiguous data blocks for structure analysis
  - `extract_string_cells()`: Extract all text cells with metadata (address, sheet, row/col)
  - `get_cell_types_in_range()`: Identify cell types ('text', 'number', 'formula', 'empty')
  - `analyze_label_value_patterns()`: Detect DCF patterns (horizontal pairs, table headers, vertical lists)
  - Label-to-value matching for understanding spreadsheet relationships

**Error Handling:**
- COM-specific exception handling with pywintypes
- Connection retry logic for application instances
- Graceful degradation when Office apps unavailable

## MCP Configuration

Add to Claude Desktop settings:
```json
{
    "mcpServers": {
        "powerpoint_mcp_win32": {
            "command": "uv",
            "args": ["run", "mcp_powerpoint_server_win32.py"],
            "cwd": "C:\\path\\to\\workspace"
        },
        "excel_mcp_win32": {
            "command": "uv", 
            "args": ["run", "mcp_excel_server_win32.py"],
            "cwd": "C:\\path\\to\\workspace"
        }
    }
}
```

## DCF Model Analysis Workflow

The Excel server now supports comprehensive DCF model analysis through structured data extraction:

### 1. Structure Discovery
```python
# Find all data regions in spreadsheet
find_used_ranges(workbook, sheet) 
# Returns: range boundaries, dimensions, coordinates
```

### 2. Label Extraction  
```python
# Extract all text cells for label identification
extract_string_cells(workbook, sheet)
# Returns: [{value: "Revenue", address: "A5", sheet: "DCF", row: 5, col: 1}, ...]
```

### 3. Pattern Recognition
```python
# Detect common DCF layouts
analyze_label_value_patterns(workbook, sheet)
# Returns: horizontal_pairs, table_headers, vertical_lists with relationships
```

### 4. Data Validation
```python
# Verify cell types in ranges  
get_cell_types_in_range(workbook, sheet, "A1:D20")
# Returns: 2D array of types ('text', 'number', 'formula', 'empty')
```

**Use Case**: Agents can now intelligently parse DCF models by first understanding the spreadsheet structure, then matching labels (like "EBITDA", "Free Cash Flow") with their corresponding value ranges before performing financial analysis.