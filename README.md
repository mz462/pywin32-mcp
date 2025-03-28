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

### Interacting with Claude

Once you've configured the MCP servers in your Claude Desktop app, you can interact with PowerPoint and Excel through natural language commands. Here are some examples:

#### PowerPoint Examples

```
You: Create a new slide with a title "Market Analysis" and add a bar chart showing revenue growth.

Claude: I'll help you create that slide with the title and chart. I'll:
1. Add a new slide
2. Add the title text
3. Create a revenue chart

[Claude will then use the MCP tools in sequence:
- add_slide
- add_text
- create_financial_chart]

You: Make the title bigger and change its color to blue.

Claude: I'll modify the title's formatting.
[Claude will use:
- find_element (to locate the title)
- edit_element (to update the formatting)]

You: Add a comparison table below the chart comparing three companies.

Claude: I'll add a comparison table below the existing chart.
[Claude will use:
- create_comparison_table]
```

#### Excel Examples

```
You: Open the Q4 report and show me the revenue numbers from cells B2 to B5.

Claude: I'll help you retrieve those revenue figures.
[Claude will use:
- list_open_workbooks (to find the workbook)
- get_range_values (to read the specified cells)]

You: Calculate the sum of these numbers and put it in cell B6.

Claude: I'll calculate the sum and write it to B6.
[Claude will use:
- get_range_values (to get the numbers)
- set_cell_value (to write the sum)]

You: Create a new sheet called "Summary" and copy these values there.

Claude: I'll create a new sheet and copy the data.
[Claude will use:
- add_worksheet
- get_range_values (from source)
- set_range_values (to destination)]
```

### How It Works

1. **Natural Language Understanding**
   - Claude interprets your requests and breaks them down into specific actions
   - It understands context from previous interactions
   - It can handle complex, multi-step operations

2. **Tool Selection**
   - Claude automatically selects the appropriate MCP tools for each task
   - It can chain multiple tools together for complex operations
   - It handles error cases and provides feedback

3. **Context Management**
   - Claude maintains context about:
     - Currently open files
     - Recent operations
     - Selected elements
     - User preferences

4. **Error Handling**
   - If an operation fails, Claude will:
     - Explain what went wrong
     - Suggest alternatives
     - Help troubleshoot common issues

### Best Practices

1. **Be Specific**
   - Mention slide numbers when relevant
   - Specify exact cell ranges in Excel
   - Describe desired formatting clearly

2. **Complex Operations**
   - Break down complex requests into steps
   - Confirm intermediate results
   - Ask for adjustments as needed

3. **Troubleshooting**
   - Ensure PowerPoint/Excel is running
   - Check file permissions
   - Verify COM automation is working
   - Run pywin32_postinstall.py if needed

### Example Workflows

#### Creating a Financial Presentation

```
You: Create a new presentation about Q4 financial results.
Claude: I'll create a new presentation with a title slide.

You: Add revenue charts for the last 4 quarters.
Claude: I'll create a new slide with a chart showing quarterly revenue.

You: Now add a comparison with our competitors.
Claude: I'll add a comparison table with key metrics for you and competitors.
```

#### Analyzing Excel Data

```
You: Show me all sheets in the Q4 analysis workbook.
Claude: I'll list all worksheets in that workbook.

You: Find the highest revenue value in column B.
Claude: I'll scan column B and find the maximum value.

You: Create a summary of the top 5 values.
Claude: I'll create a new sheet with the top 5 revenue figures.
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


## Notes

- Both servers require Windows and their respective Microsoft Office applications installed
- The servers interact with *running* instances of the applications
- COM automation requires proper initialization; run the post-install script if you encounter COM-related errors
- For better constant handling, consider using `makepy` to generate Office constants