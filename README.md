# Enhanced PowerPoint Model Context Protocol (MCP)

This implementation provides an Enhanced Model Context Protocol (MCP) server and client for PowerPoint operations, following Anthropic's MCP standard. It enables AI systems to interact with PowerPoint presentations through a standardized interface with granular control over elements, shapes, charts, and templates.

## Features

### Core Capabilities
- **Modular Slide Editing**: Edit any element on a slide with precise control (text, position, color, size)
- **Shape & Visual Element Support**: Create and modify PowerPoint shapes with custom formatting
- **Chart & Graph Integration**: Create financial charts with Proff API data for Norwegian companies
- **Content Templates**: Template support for common slide types and industry-specific templates

### Advanced Operations
- Fine-grained element operations (find, edit, style, group)
- Advanced shape handling and connectors
- Financial data integration and visualization
- Template system for creating and using slide templates

## Installation

1. Install the required dependencies:
```bash
pip install -r requirements.txt
```

2. Create presentation workspace directory (optional, will be created automatically):
```bash
mkdir -p presentations/templates
```

## Usage

### Starting the MCP Server

Run the server:
```bash
python mcp_powerpoint_server.py
```

The server will start on `http://localhost:8000` by default.

### Using the MCP Client

```python
from mcp_powerpoint_client import PowerPointMCPClient

# Initialize the client
client = PowerPointMCPClient()

# Create a new presentation
presentation_path = "example.pptx"

# Add a slide
client.add_slide(presentation_path)

# Add text to the first slide
client.add_text(presentation_path, 0, "Hello, World!", position=(1, 1), font_size=24)

# Style the text element
element_id = client.find_element(presentation_path, 0, "text", "Hello, World!")["data"]["elements"][0]["id"]
client.style_element(presentation_path, 0, element_id, {
    "font": {
        "size": 32,
        "bold": True,
        "color": "#0000FF"
    },
    "fill": {
        "type": "solid",
        "color": "#FFFF00"
    }
})

# Add a shape
client.add_shape(presentation_path, 0, "rectangle", 
                position={"x": 1, "y": 3}, 
                size={"width": 3, "height": 2},
                style_properties={
                    "fill": {
                        "type": "solid",
                        "color": "#FF0000"
                    }
                })

# Get financial data
financials = client.get_company_financials(presentation_path, "Sample Company AS", 
                                        metrics=["revenue", "ebitda", "profit"])

# Create a financial chart
client.create_financial_chart(
    presentation_path, 0, "column",
    data={
        "categories": list(financials["data"]["financials"].keys()),
        "series": [
            {
                "name": "Revenue",
                "values": [financials["data"]["financials"][year]["revenue"] 
                          for year in financials["data"]["financials"]]
            }
        ]
    },
    position={"x": 1, "y": 6},
    size={"width": 6, "height": 4},
    title="Revenue Over Time"
)

# Save the presentation
client.save(presentation_path)
```

### Creating Company Overview Presentations

The client includes a convenience method for creating complete company overview presentations:

```python
# Create a complete company overview presentation
client.create_company_overview_presentation("Acme Inc")
```

## Available Operations

### Element Operations
- `find_element`: Find elements on a slide based on type, text, or position
- `edit_element`: Edit properties of a specific element
- `style_element`: Apply styling to a specific element
- `group_elements`: Group multiple elements together

### Shape Operations
- `add_shape`: Add a new shape to a slide
- `connect_shapes`: Create a connector between two shapes

### Financial Data Operations
- `get_company_financials`: Fetch financial data for a Norwegian company
- `create_financial_chart`: Create a financial chart on a slide
- `create_comparison_table`: Create a table comparing multiple companies

### Template Operations
- `apply_template`: Apply a template to a presentation
- `create_slide_from_template`: Create a new slide based on a template
- `save_as_template`: Save a slide as a template

### Basic Operations
- `add_slide`: Add a new slide to the presentation
- `add_text`: Add text to a specific slide
- `add_image`: Add an image to a specific slide
- `get_slide_count`: Get the total number of slides
- `get_slide_text`: Get all text from a specific slide
- `set_background_color`: Set the background color of a slide
- `delete_slide`: Delete a specific slide
- `save`: Save the presentation

## API Documentation

The server provides automatic API documentation at:
- Swagger UI: `http://localhost:8000/docs`
- ReDoc: `http://localhost:8000/redoc`

## Implementation Status

### Completed
- ✅ Extended PowerPoint object model with unique ID system for elements
- ✅ Built comprehensive element finding, editing, and styling capabilities
- ✅ Implemented advanced shape creation and connector functionality
- ✅ Created financial data processing with simulated Proff API data
- ✅ Built financial chart generation with various chart types
- ✅ Implemented the template system with storage and management
- ✅ Created robust error handling and validation
- ✅ Built RESTful API server with standardized endpoints
- ✅ Implemented client library for easy interaction

### Future Work
1. Create a comprehensive library of business templates
2. Replace simulation with actual Proff API connection
3. Extend python-pptx to better support element grouping
4. Improve the slide preview rendering quality
5. Optimize operations for faster execution
6. Implement more sophisticated caching for API data
7. Add version control for templates and presentations

## Error Handling

All operations return a response with the following structure:
```json
{
    "success": true/false,
    "data": {...} or null,
    "error": "error message" or null
}
```

## Security Considerations

- The server currently allows CORS from all origins. In production, you should restrict this to specific origins.
- File paths should be validated and sanitized in production.
- Consider adding authentication and authorization mechanisms for production use.
- API keys for external services should be stored securely using environment variables or a secrets manager.