# Microsoft Office MCP Servers

This repository contains Model Context Protocol (MCP) servers for interacting with Microsoft Office applications through AI assistance. Currently supported applications:
- PowerPoint: Create and manipulate presentations
- Excel: Interact with workbooks and spreadsheets

Both servers use `pywin32` for COM automation, allowing direct interaction with running Office applications.

## Prerequisites

- Windows operating system
- Microsoft Office installed (PowerPoint and/or Excel)
- Python 3.7+
- `pywin32` package

## Installation

1. Clone the repository:
```bash
git clone https://github.com/jenstangen1/mcp-pptx.git
cd mcp-pptx
```

2. Install dependencies using `uv`:
```bash
uv pip install pywin32
```

3. Run the `pywin32` post-install script with administrator privileges:
```bash
python C:\path\to\your\env\Scripts\pywin32_postinstall.py -install
```

## Setting up with Claude

To integrate these MCP servers with Claude, add the following configuration to your Claude Desktop app settings:

```json
{
    "mcpServers": {
        "powerpoint_mcp_win32": {
            "command": "uv",
            "args": [
                "run",
                "mcp_powerpoint_server_win32.py"
            ],
            "cwd": "C:\\path\\to\\your\\workspace"
        },
        "excel_mcp_win32": {
            "command": "uv",
            "args": [
                "run",
                "mcp_excel_server_win32.py"
            ],
            "cwd": "C:\\path\\to\\your\\workspace"
        }
    }
}
```

Note: Replace `C:\\path\\to\\your\\workspace` with your actual workspace path.

# PowerPoint MCP Server

The PowerPoint server provides a comprehensive API for AI models to interact with PowerPoint presentations, supporting advanced formatting, financial charts, and data integration.

## Features

### Presentation Management
- Create and modify PowerPoint presentations
- Add, delete, and modify slides
- Save and load presentations from workspace
- Template management system

### Element Operations
- Fine-grained control over slide elements (text, shapes, images, charts)
- Advanced shape creation and styling
- Element positioning and grouping
- Connector lines between shapes

### Financial Integration
- Create financial charts (line, bar, column, pie, waterfall, etc.)
- Generate comparison tables
- Support for various financial metrics:
  - Revenue
  - EBITDA
  - Profit
  - Assets
  - Equity
  - Growth rates
  - Margins
- Currently uses dummy data, with plans to integrate Proff API for Norwegian company data
- Adaptable to other financial data providers through API customization

### Styling and Formatting
- Rich text formatting
- Shape styling (fills, gradients, outlines)
- Chart customization
- Background colors and effects

## Installation

1. Clone the repository:
```bash
git clone https://github.com/jenstangen1/pptx-mcp.git
cd pptx-mcp
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Setting up with Claude

### Installing the MCP into Claude's Interface

To integrate this PowerPoint MCP with Claude, add the following JSON configuration to your Claude MCP file:

```json
{
    "mcpServers": {
      "powerpoint_mcp": {
        "command": "uv",
        "args": [
          "--directory",
          "your directory here",
          "run",
          "mcp_powerpoint_server.py"
        ]
      }
    }
}
```

Note: You may need to modify the directory path to match your installation location.

## Available MCP Tools

### Presentation Management
- `list_presentations`: List all PowerPoint files in the workspace
- `upload_presentation`: Upload a new presentation to the workspace
- `save_presentation`: Save the current presentation

### Slide Operations
- `add_slide`: Add a new slide to the presentation
- `delete_slide`: Delete a slide from the presentation
- `get_slide_count`: Get the total number of slides in the presentation
- `analyze_slide`: Analyze the content of a slide
- `set_background_color`: Set the background color of a slide

### Element Operations
- `add_text`: Add text to a slide
- `add_shape`: Add a shape to a slide
- `edit_element`: Edit an element's properties
- `style_element`: Apply styling to an element
- `connect_shapes`: Connect two shapes with a connector
- `find_element`: Find elements on a slide based on criteria

### Financial Tools
- `get_company_financials`: Get financial data for a company (currently returns dummy data)
- `create_financial_chart`: Create a financial chart on a slide
- `create_comparison_table`: Create a comparison table for companies

**Note:** The financial tools currently use dummy data. Future versions plan to integrate with the Proff API for automatic fetching of Norwegian company data. Users can modify the code to connect with local or preferred financial data providers.

### Template Operations
- `list_templates`: List all available templates
- `apply_template`: Apply a template to a presentation
- `create_slide_from_template`: Create a new slide from a template
- `save_as_template`: Save a slide as a template

### Debug Tools
- `debug_element_mappings`: Debug tool to inspect element mappings for a slide

## Usage

### Starting the Server

Run the server:
```bash
python mcp_powerpoint_server.py
```

The server will create a workspace directory for presentations and templates if they don't exist.

### Basic Operations

```python
# List presentations
presentations = mcp.list_presentations()

# Create a new slide
slide_index = mcp.add_slide(presentation_path, layout_name="Title and Content")

# Add text to a slide
element_id = mcp.add_text(
    presentation_path=presentation_path,
    slide_index=slide_index,
    text="Hello World",
    position=[1.0, 1.0],
    font_size=24
)

# Add a shape
shape_id = mcp.add_shape(
    presentation_path=presentation_path,
    slide_index=slide_index,
    shape_type="rectangle",
    position={"x": 2.0, "y": 2.0},
    size={"width": 2.0, "height": 1.0}
)
```

### Financial Charts

```python
# Create a financial chart
chart_id = mcp.create_financial_chart(
    presentation_path=presentation_path,
    slide_index=slide_index,
    chart_type="column",
    data={
        "categories": ["2020", "2021", "2022"],
        "series": [{
            "name": "Revenue",
            "values": [1000000, 1200000, 1500000]
        }]
    },
    position={"x": 1.0, "y": 1.0},
    size={"width": 6.0, "height": 4.0},
    title="Revenue Growth"
)

# Create a comparison table
table_id = mcp.create_comparison_table(
    presentation_path=presentation_path,
    slide_index=slide_index,
    companies=["Company A", "Company B"],
    metrics=["revenue", "ebitda", "margin"],
    position={"x": 1.0, "y": 1.0},
    title="Company Comparison"
)
```

### Template Management

```python
# List available templates
templates = mcp.list_templates()

# Apply a template
mcp.apply_template(
    presentation_path=presentation_path,
    template_name="financial_report",
    options={
        "apply_master": True,
        "apply_theme": True,
        "apply_layouts": True
    }
)

# Create a slide from template
mcp.create_slide_from_template(
    presentation_path=presentation_path,
    template_name="comparison_slide",
    content={
        "title": "Market Analysis",
        "subtitle": "Q3 2023"
    }
)
```

## Directory Structure

```
pptx-mcp/
├── mcp_powerpoint_server.py  # Main server implementation
├── requirements.txt          # Python dependencies
├── presentations/           # Workspace for presentations
│   └── templates/          # Template storage
└── README.md               # This file
```

## Dependencies

- python-pptx: PowerPoint file manipulation
- Pillow: Image processing
- numpy: Numerical operations
- MCP SDK: Model Context Protocol implementation

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

# Excel MCP Server

The Excel server provides tools for interacting with Excel workbooks, worksheets, and cell data through AI assistance.

## Features

### Workbook Management
- Connect to running Excel instances
- List open workbooks
- Save workbooks with various formats (.xlsx, .xlsm, .xlsb, .xls)

### Worksheet Operations
- List worksheets in a workbook
- Add new worksheets
- Access worksheets by name or index

### Cell and Range Operations
- Read and write individual cell values
- Get and set values for cell ranges
- Handle various data types (text, numbers, dates, currency)
- Automatic type conversion for dates and currency values

## Available MCP Tools

### Workbook Management
- `list_open_workbooks`: List all currently open Excel workbooks
- `save_workbook`: Save a workbook to disk with optional format selection

### Worksheet Operations
- `list_worksheets`: List all worksheets in a workbook
- `add_worksheet`: Add a new worksheet to a workbook
- `get_worksheet`: Get a worksheet by name or index

### Cell and Range Operations
- `get_cell_value`: Read a single cell's value
- `set_cell_value`: Set a single cell's value
- `get_range_values`: Read values from a range of cells
- `set_range_values`: Set values for a range of cells

## Usage Examples

### Basic Operations

```python
# List open workbooks
workbooks = mcp.list_open_workbooks()

# List worksheets in a workbook
sheets = mcp.list_worksheets(workbook_identifier="Book1.xlsx")

# Read a cell value
value = mcp.get_cell_value(
    identifier="Book1.xlsx",
    sheet_identifier="Sheet1",
    cell_address="A1"
)

# Write to a cell
mcp.set_cell_value(
    identifier="Book1.xlsx",
    sheet_identifier="Sheet1",
    cell_address="B1",
    value="Hello World"
)
```

### Working with Ranges

```python
# Read a range of cells
values = mcp.get_range_values(
    identifier="Book1.xlsx",
    sheet_identifier="Sheet1",
    range_address="A1:C5"
)

# Write to a range of cells
data = [
    ["Name", "Age", "City"],
    ["John", 30, "New York"],
    ["Jane", 25, "London"]
]
mcp.set_range_values(
    identifier="Book1.xlsx",
    sheet_identifier="Sheet1",
    start_cell="A1",
    values=data
)
```

### Saving Workbooks

```python
# Save with current name
mcp.save_workbook(identifier="Book1.xlsx")

# Save with a new name/location
mcp.save_workbook(
    identifier="Book1.xlsx",
    save_path="C:\\Documents\\NewBook.xlsx"
)
```

## Notes

- Both servers require Windows and their respective Microsoft Office applications installed
- The servers interact with *running* instances of the applications
- COM automation requires proper initialization; run the post-install script if you encounter COM-related errors
- For better constant handling, consider using `makepy` to generate Office constants