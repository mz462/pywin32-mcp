import win32com.client
import pythoncom
from typing import List, Optional, Dict, Any, Union
from mcp.server.fastmcp import FastMCP
import os
import pywintypes # Import for specific exception types

# Constants from PowerPoint VBA Object Library (usually obtained via makepy)
# Using magic numbers for simplicity here, but using makepy is recommended practice
# Example: from win32com.client import constants
# ppSaveAsDefault = 1  (This might vary, better to use format-specific saves)
ppSaveAsOpenXMLPresentation = 24 # .pptx
msoShapeRectangle = 1
msoTextOrientationHorizontal = 1
msoPlaceholder = 14 # Shape type for placeholders

# Placeholder type constants (might vary slightly by PP version)
ppPlaceholderTitle = 1
ppPlaceholderBody = 2
ppPlaceholderCenterTitle = 13 # Often used for Title Only layout
ppPlaceholderSubtitle = 14
ppPlaceholderDate = 16
ppPlaceholderSlideNumber = 15
ppPlaceholderFooter = 17
ppPlaceholderHeader = 18
ppPlaceholderObject = 7 # Generic object/content placeholder


# Mapping for user-friendly placeholder names to constants
PLACEHOLDER_NAME_MAP = {
    "title": ppPlaceholderTitle,
    "body": ppPlaceholderBody,
    "centertitle": ppPlaceholderCenterTitle,
    "subtitle": ppPlaceholderSubtitle,
    "date": ppPlaceholderDate,
    "slidenumber": ppPlaceholderSlideNumber,
    "footer": ppPlaceholderFooter,
    "header": ppPlaceholderHeader,
    "object": ppPlaceholderObject,
    "content": ppPlaceholderObject, # Common synonym
}

# Mapping for user-friendly shape type names (reverse of _get_shape_type_name)
SHAPE_TYPE_NAME_MAP = {
    "rectangle": 1,
    "textbox": 17, # MSO_SHAPE_TYPE.TEXT_BOX (Note: might also be AutoShape with text)
    "oval": 9,     # MSO_SHAPE_TYPE.OVAL (Note: Check MSO AutoShape constants for more specific ovals)
    "table": 19,   # MSO_SHAPE_TYPE.TABLE
    "chart": 3,    # MSO_SHAPE_TYPE.CHART
    "picture": 13, # MSO_SHAPE_TYPE.PICTURE
    "line": 20,    # MSO_SHAPE_TYPE.LINE
    "connector": 10, # MSO_SHAPE_TYPE.CONNECTOR (check AutoShape types)
    "placeholder": msoPlaceholder,
    # Add more basic types as needed
}


class PowerPointEditorWin32:
    def __init__(self):
        self.app = None
        self._connect_or_launch_powerpoint()

    def _connect_or_launch_powerpoint(self):
        """Connects to a running instance of PowerPoint or launches a new one."""
        try:
            # Use the Pywin32 CoInitializeEx to avoid threading issues with COM
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            self.app = win32com.client.GetActiveObject("PowerPoint.Application")
            print("Connected to running PowerPoint application.")
        except pywintypes.com_error:
            try:
                self.app = win32com.client.Dispatch("PowerPoint.Application")
                self.app.Visible = True  # Make the application visible
                print("Launched new PowerPoint application.")
            except Exception as e:
                print(f"Error launching PowerPoint: {e}")
                self.app = None
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            self.app = None

    def _ensure_connection(self):
        """Ensures the PowerPoint application object is valid."""
        if self.app is None:
            self._connect_or_launch_powerpoint()
        if self.app is None:
            raise ConnectionError("Could not connect to or launch PowerPoint.")
        # Basic check if the app object seems responsive
        try:
            _ = self.app.Version
        except Exception as e:
            print(f"PowerPoint connection lost or unresponsive: {e}")
            self._connect_or_launch_powerpoint() # Try reconnecting
            if self.app is None:
                 raise ConnectionError("Could not reconnect to PowerPoint.")

    def list_open_presentations(self) -> List[Dict[str, Any]]:
        """Lists all currently open presentations."""
        self._ensure_connection()
        presentations_info = []
        try:
            for i in range(1, self.app.Presentations.Count + 1):
                pres = self.app.Presentations(i)
                presentations_info.append({
                    "name": pres.Name,
                    "path": pres.FullName if pres.Path else None,
                    "slides": pres.Slides.Count,
                    "saved": pres.Saved,
                    "index": i # Provide the 1-based index for reference
                })
        except Exception as e:
            print(f"Error listing presentations: {e}")
            # Maybe attempt reconnect if it's a COM error
            if "RPC server is unavailable" in str(e):
                 self._connect_or_launch_powerpoint()
                 # Retry might be needed here depending on strategy
            raise
        return presentations_info

    def get_presentation(self, identifier: str) -> Optional[Any]:
        """
        Gets a presentation object by its name, path, or 1-based index.

        Args:
            identifier (str): The name (e.g., "Presentation1.pptx"),
                              full path, or 1-based index (as string or int).
        """
        self._ensure_connection()
        try:
            # Try by index first if it's an integer
            if isinstance(identifier, int) or identifier.isdigit():
                idx = int(identifier)
                if 1 <= idx <= self.app.Presentations.Count:
                    return self.app.Presentations(idx)
                else:
                    print(f"Index {identifier} out of range.")
                    return None

            # Try by name or path
            for i in range(1, self.app.Presentations.Count + 1):
                pres = self.app.Presentations(i)
                if pres.Name == identifier or pres.FullName == identifier:
                    return pres
            print(f"Presentation '{identifier}' not found.")
            return None
        except Exception as e:
            print(f"Error getting presentation '{identifier}': {e}")
            raise

    def save_presentation(self, identifier: str, save_path: Optional[str] = None):
        """
        Saves the specified presentation.

        Args:
            identifier (str): Name, path, or index of the presentation.
            save_path (Optional[str]): Path to save to. If None, saves to its current path.
                                      If the presentation is new, save_path is required.
        """
        self._ensure_connection()
        pres = self.get_presentation(identifier)
        if not pres:
            raise ValueError(f"Presentation '{identifier}' not found.")

        try:
            if save_path:
                # Ensure the directory exists
                abs_path = os.path.abspath(save_path)
                os.makedirs(os.path.dirname(abs_path), exist_ok=True)
                pres.SaveAs(abs_path, ppSaveAsOpenXMLPresentation)
                print(f"Presentation saved as '{abs_path}'.")
            elif pres.Path: # Can only save if it has a path already
                pres.Save()
                print(f"Presentation '{pres.Name}' saved.")
            else:
                raise ValueError("save_path is required for a new presentation.")
        except Exception as e:
            print(f"Error saving presentation '{identifier}': {e}")
            raise

    def add_slide(self, identifier: str, layout_index: int = 1) -> int:
        """
        Adds a new slide to the presentation.

        Args:
            identifier (str): Name, path, or index of the presentation.
            layout_index (int): 1-based index of the slide layout to use.

        Returns:
            int: The 1-based index of the newly added slide.
        """
        self._ensure_connection()
        pres = self.get_presentation(identifier)
        if not pres:
            raise ValueError(f"Presentation '{identifier}' not found.")

        try:
            # Ensure layout_index is valid
            if not (1 <= layout_index <= pres.SlideMaster.CustomLayouts.Count):
                 print(f"Warning: Layout index {layout_index} invalid. Using layout 1.")
                 layout_index = 1
            layout = pres.SlideMaster.CustomLayouts(layout_index)

            # Add the slide (returns the new Slide object)
            new_slide = pres.Slides.AddSlide(pres.Slides.Count + 1, layout)
            print(f"Added slide {new_slide.SlideIndex} to '{pres.Name}'.")
            return new_slide.SlideIndex
        except Exception as e:
            print(f"Error adding slide to '{identifier}': {e}")
            raise

    def get_slide(self, identifier: str, slide_index: int) -> Optional[Any]:
        """
        Gets a slide object from a presentation.

        Args:
            identifier (str): Name, path, or index of the presentation.
            slide_index (int): 1-based index of the slide.

        Returns:
            Optional[Any]: The slide object or None if not found.
        """
        self._ensure_connection()
        pres = self.get_presentation(identifier)
        if not pres:
            return None

        try:
            if 1 <= slide_index <= pres.Slides.Count:
                return pres.Slides(slide_index)
            else:
                print(f"Slide index {slide_index} out of range for '{pres.Name}'.")
                return None
        except Exception as e:
            print(f"Error getting slide {slide_index} from '{identifier}': {e}")
            raise

    def add_text_box(self, identifier: str, slide_index: int, text: str,
                     left: float, top: float, width: float, height: float) -> int:
        """
        Adds a text box to a slide.

        Args:
            identifier (str): Presentation identifier.
            slide_index (int): 1-based slide index.
            text (str): Text content for the box.
            left, top, width, height (float): Position and size in points.

        Returns:
            int: The unique ID of the newly added shape.
        """
        self._ensure_connection()
        slide = self.get_slide(identifier, slide_index)
        if not slide:
            raise ValueError(f"Slide {slide_index} not found in presentation '{identifier}'.")

        try:
            shape = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, left, top, width, height)
            shape.TextFrame.TextRange.Text = text
            shape.Name = f"TextBox_{shape.Id}" # Assign a default name
            print(f"Added text box (ID: {shape.Id}) to slide {slide_index}.")
            return shape.Id
        except Exception as e:
            print(f"Error adding text box to slide {slide_index}: {e}")
            raise

    def add_shape(self, identifier: str, slide_index: int, shape_type: int,
                  left: float, top: float, width: float, height: float) -> int:
        """
        Adds a basic auto shape to a slide.

        Args:
            identifier (str): Presentation identifier.
            slide_index (int): 1-based slide index.
            shape_type (int): MSO AutoShapeType constant (e.g., msoShapeRectangle).
            left, top, width, height (float): Position and size in points.

        Returns:
            int: The unique ID of the newly added shape.
        """
        self._ensure_connection()
        slide = self.get_slide(identifier, slide_index)
        if not slide:
            raise ValueError(f"Slide {slide_index} not found in presentation '{identifier}'.")

        try:
            # Use AddShape for AutoShapes
            shape = slide.Shapes.AddShape(shape_type, left, top, width, height)
            shape.Name = f"Shape_{shape.Id}" # Assign a default name
            print(f"Added shape (ID: {shape.Id}, Type: {shape_type}) to slide {slide_index}.")
            return shape.Id
        except Exception as e:
            print(f"Error adding shape to slide {slide_index}: {e}")
            raise

    def get_shape_by_id(self, identifier: str, slide_index: int, shape_id: int) -> Optional[Any]:
        """Gets a shape object by its unique ID."""
        self._ensure_connection()
        slide = self.get_slide(identifier, slide_index)
        if not slide:
            return None

        try:
            # Iterate through shapes to find by ID
            for i in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(i)
                if shape.Id == shape_id:
                    return shape
            print(f"Shape with ID {shape_id} not found on slide {slide_index}.")
            return None
        except Exception as e:
            # Handle potential errors if shape is deleted during iteration etc.
            print(f"Error finding shape ID {shape_id} on slide {slide_index}: {e}")
            return None

    def get_shape_by_name(self, identifier: str, slide_index: int, shape_name: str) -> Optional[Any]:
        """Gets a shape object by its name."""
        self._ensure_connection()
        slide = self.get_slide(identifier, slide_index)
        if not slide:
            return None

        try:
            # Accessing by name directly might fail if name is not unique or contains odd chars
            return slide.Shapes(shape_name)
        except pywintypes.com_error as e:
            # Handle common error for item not found
            if e.hresult == -2147024809: # 0x80070057 (E_INVALIDARG often means not found by name)
                print(f"Shape with name '{shape_name}' not found on slide {slide_index}.")
            else:
                print(f"Error finding shape name '{shape_name}' on slide {slide_index}: {e}")
            return None
        except Exception as e:
            print(f"Error finding shape name '{shape_name}' on slide {slide_index}: {e}")
            return None

    def find_shape_by_text(self, identifier: Union[str, int], slide_index: int, search_text: str, partial_match: bool = True) -> List[Dict[str, Any]]:
        """
        Finds shapes on a slide containing specific text.

        Args:
            identifier (Union[str, int]): Presentation identifier.
            slide_index (int): 1-based slide index.
            search_text (str): The text to search for (case-insensitive).
            partial_match (bool): If True, finds shapes where search_text is a substring.
                                  If False, requires an exact match (ignoring case).

        Returns:
            List[Dict[str, Any]]: List of matching shapes with their info.
        """
        self._ensure_connection()
        matches = []
        slide = self.get_slide(identifier, slide_index)
        if not slide:
            print(f"Cannot find shapes by text: Slide {slide_index} not found in presentation '{identifier}'.")
            return []

        search_lower = search_text.lower()
        try:
            for i in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(i)
                shape_text = ""
                has_text_frame = False
                try:
                    # Check text frame exists and has text
                    if shape.HasTextFrame and shape.TextFrame.HasText:
                        has_text_frame = True
                        shape_text = shape.TextFrame.TextRange.Text
                except Exception:
                    continue # Skip shapes that error on text access

                if has_text_frame:
                    shape_text_lower = shape_text.lower()
                    found = False
                    if partial_match:
                        if search_lower in shape_text_lower:
                            found = True
                    else: # Exact match (case-insensitive)
                        if search_lower == shape_text_lower:
                            found = True

                    if found:
                        shape_data = self._get_shape_basic_info(shape)
                        shape_data["text_preview"] = shape_text[:100] + "..." if len(shape_text) > 100 else shape_text
                        matches.append(shape_data)

        except Exception as e:
            print(f"Error searching for text '{search_text}' on slide {slide_index}: {e}")
        return matches

    def find_shapes_by_type(self, identifier: Union[str, int], slide_index: int, shape_type_name: str) -> List[Dict[str, Any]]:
        """
        Finds shapes on a slide matching a specific type name.

        Args:
            identifier (Union[str, int]): Presentation identifier.
            slide_index (int): 1-based slide index.
            shape_type_name (str): The user-friendly name of the shape type (e.g., "rectangle", "textbox"). Case-insensitive.

        Returns:
            List[Dict[str, Any]]: List of matching shapes with their info.
        """
        self._ensure_connection()
        matches = []
        slide = self.get_slide(identifier, slide_index)
        if not slide:
            print(f"Cannot find shapes by type: Slide {slide_index} not found in presentation '{identifier}'.")
            return []

        type_name_lower = shape_type_name.lower()
        target_type_id = SHAPE_TYPE_NAME_MAP.get(type_name_lower)

        if target_type_id is None:
            print(f"Warning: Unknown shape type name '{shape_type_name}'. Supported types: {list(SHAPE_TYPE_NAME_MAP.keys())}")
            return []

        try:
            for i in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(i)
                # Special handling for textbox which might be an AutoShape
                is_match = False
                if target_type_id == SHAPE_TYPE_NAME_MAP["textbox"]:
                     # Check MSO_SHAPE_TYPE.TEXT_BOX or if it's an AutoShape with text
                     if shape.Type == SHAPE_TYPE_NAME_MAP["textbox"]:
                         is_match = True
                     elif shape.Type == 1: # msoAutoShape
                          try:
                              if shape.HasTextFrame and shape.TextFrame.HasText:
                                 # Could refine this - maybe check shape.AutoShapeType?
                                 # For now, consider any autoshape with text as potential textbox match
                                 is_match = True
                          except: pass # Ignore errors accessing text frame
                elif shape.Type == target_type_id:
                    is_match = True

                if is_match:
                    matches.append(self._get_shape_basic_info(shape))

        except Exception as e:
            print(f"Error searching for shape type '{shape_type_name}' on slide {slide_index}: {e}")
        return matches

    def get_placeholder_shape(self, identifier: Union[str, int], slide_index: int, placeholder_name: str) -> Optional[Dict[str, Any]]:
        """
        Finds a specific placeholder shape on a slide by its common name.

        Args:
            identifier (Union[str, int]): Presentation identifier.
            slide_index (int): 1-based slide index.
            placeholder_name (str): Common name of the placeholder (e.g., "title", "body", "footer"). Case-insensitive.

        Returns:
            Optional[Dict[str, Any]]: Info of the first matching placeholder shape, or None if not found.
        """
        self._ensure_connection()
        slide = self.get_slide(identifier, slide_index)
        if not slide:
            print(f"Cannot find placeholder: Slide {slide_index} not found in presentation '{identifier}'.")
            return None

        name_lower = placeholder_name.lower()
        target_ph_type = PLACEHOLDER_NAME_MAP.get(name_lower)

        if target_ph_type is None:
            print(f"Warning: Unknown placeholder name '{placeholder_name}'. Supported names: {list(PLACEHOLDER_NAME_MAP.keys())}")
            return None

        try:
            for i in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(i)
                try:
                    # Check if it's a placeholder and its type matches
                    if shape.Type == msoPlaceholder and shape.PlaceholderFormat.Type == target_ph_type:
                        print(f"Found placeholder '{placeholder_name}' (ID: {shape.Id}) on slide {slide_index}.")
                        return self._get_shape_basic_info(shape)
                except Exception:
                    # Some shapes might error on PlaceholderFormat access if not placeholders
                    continue
        except Exception as e:
            print(f"Error searching for placeholder '{placeholder_name}' on slide {slide_index}: {e}")

        print(f"Placeholder '{placeholder_name}' not found on slide {slide_index}.")
        return None

    def edit_element(self, identifier: str, slide_index: int, shape_identifier: Union[int, str],
                     properties: Dict[str, Any]) -> bool:
        """
        Edits properties of a shape identified by ID or name.

        Args:
            identifier (str): Presentation identifier.
            slide_index (int): 1-based slide index.
            shape_identifier (Union[int, str]): The shape's ID (int) or Name (str).
            properties (Dict[str, Any]): Dictionary of properties to change.
                                         Supported: 'text', 'left', 'top', 'width', 'height', 'name'.

        Returns:
            bool: True if successful, False otherwise.
        """
        self._ensure_connection()
        if isinstance(shape_identifier, int):
            shape = self.get_shape_by_id(identifier, slide_index, shape_identifier)
        elif isinstance(shape_identifier, str):
            shape = self.get_shape_by_name(identifier, slide_index, shape_identifier)
        else:
             print("Invalid shape_identifier type. Use int (ID) or str (Name).")
             return False

        if not shape:
            return False

        try:
            if 'text' in properties and shape.HasTextFrame and shape.TextFrame.HasText:
                 shape.TextFrame.TextRange.Text = str(properties['text'])
            if 'left' in properties:
                 shape.Left = float(properties['left'])
            if 'top' in properties:
                 shape.Top = float(properties['top'])
            if 'width' in properties:
                 shape.Width = float(properties['width'])
            if 'height' in properties:
                 shape.Height = float(properties['height'])
            if 'name' in properties:
                 shape.Name = str(properties['name']) # Allow renaming

            print(f"Edited properties for shape '{shape_identifier}' on slide {slide_index}.")
            return True
        except Exception as e:
            print(f"Error editing shape '{shape_identifier}' on slide {slide_index}: {e}")
            return False

    def list_shapes(self, identifier: Union[str, int], slide_index: int) -> List[Dict[str, Any]]:
        """Lists shapes on a slide with their ID, Name, Type, and basic geometry."""
        self._ensure_connection()
        shapes_info = []
        slide = self.get_slide(identifier, slide_index)
        if not slide:
            print(f"Cannot list shapes: Slide {slide_index} not found in presentation '{identifier}'.")
            return []

        try:
            for i in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(i)
                shapes_info.append(self._get_shape_basic_info(shape)) # Use helper
        except Exception as e:
            print(f"Error listing shapes on slide {slide_index}: {e}")
            # Don't raise, just return what we have or empty list
        return shapes_info

    def _get_shape_basic_info(self, shape: Any) -> Dict[str, Any]:
        """Helper to get common info dictionary for a shape."""
        info = {
            "id": -1, "name": "Unknown", "type_id": -1, "type_name": "Unknown",
            "left": 0, "top": 0, "width": 0, "height": 0,
            "has_text": False, "is_placeholder": False
        }
        try:
            info["id"] = shape.Id
            info["name"] = shape.Name
            info["type_id"] = shape.Type
            info["type_name"] = self._get_shape_type_name(shape.Type)
            info["left"] = shape.Left
            info["top"] = shape.Top
            info["width"] = shape.Width
            info["height"] = shape.Height

            # Check for text
            try:
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    info["has_text"] = True
            except Exception: pass

            # Check if placeholder
            try:
                if shape.Type == msoPlaceholder:
                     info["is_placeholder"] = True
                     info["placeholder_type_id"] = shape.PlaceholderFormat.Type
                     # Add friendly name for placeholder type if possible
                     for name, ph_id in PLACEHOLDER_NAME_MAP.items():
                         if ph_id == info["placeholder_type_id"]:
                             info["placeholder_type_name"] = name
                             break
            except Exception: pass

        except Exception as e:
            print(f"Error getting basic info for a shape: {e}")
            # Return partial info if possible
        return info

    def _get_shape_type_name(self, type_id: int) -> str:
        """Returns a readable name for MSO_SHAPE_TYPE IDs (add more as needed)."""
        # This is a simplified mapping. A full mapping would be extensive.
        # Consider reversing SHAPE_TYPE_NAME_MAP for consistency
        mapping = {
            1: "Rectangle/AutoShape", # Generic AutoShape
            17: "TextBox",
            9: "Oval", # Generic Oval (likely AutoShape)
            19: "Table",
            3: "Chart",
            13: "Picture",
            20: "Line",
            10: "Connector",
            msoPlaceholder: "Placeholder",
            # Add more from MSO_SHAPE_TYPE if needed
            6: "Group",
            7: "EmbeddedObject",
            8: "FormControl",
            11: "Freeform",
            12: "Media",
            15: "OLEControlObject",
            16: "ScriptAnchor",
            18: "Canvas",
            21: "Ink",
            22: "InkComment",
            23: "Diagram", # SmartArt? Check MSO_SHAPE_TYPE constants
            24: "WebVideo",
            25: "ContentApp", # Office Add-in
            26: "GraphicFrame", # Holds Table, Chart, SmartArt, etc.
            27: "LinkedOLEObject",
            28: "LinkedPicture",
            29: "Model3D",
            30: "LinkedContentApp",
        }
        return mapping.get(type_id, f"Unknown ({type_id})")


# --- MCP Server Setup ---

# Create the PowerPoint editor instance
editor = PowerPointEditorWin32()

# Create MCP server
mcp = FastMCP("PowerPoint MCP (Win32)")

@mcp.tool()
def list_open_presentations():
    """Lists currently open PowerPoint presentations."""
    try:
        return {"presentations": editor.list_open_presentations()}
    except Exception as e:
        return {"error": f"Failed to list presentations: {str(e)}"}

@mcp.tool()
def save_presentation(identifier: str, save_path: str = None):
    """Saves the specified presentation. Use index, name, or full path as identifier."""
    try:
        editor.save_presentation(identifier, save_path)
        return {"message": f"Save command issued for presentation '{identifier}'."}
    except Exception as e:
        return {"error": f"Failed to save presentation '{identifier}': {str(e)}"}

@mcp.tool()
def add_slide(identifier: str, layout_index: int = 1):
    """Adds a slide to the specified presentation. Use index, name, or full path as identifier."""
    try:
        new_slide_index = editor.add_slide(identifier, layout_index)
        return {"message": "Slide added successfully.", "slide_index": new_slide_index}
    except Exception as e:
        return {"error": f"Failed to add slide to '{identifier}': {str(e)}"}

@mcp.tool()
def add_text_box(identifier: str, slide_index: int, text: str,
                 left: float = 72, top: float = 72, width: float = 288, height: float = 72):
    """Adds a text box to a specific slide (dimensions in points)."""
    try:
        shape_id = editor.add_text_box(identifier, slide_index, text, left, top, width, height)
        return {"message": "Text box added successfully.", "shape_id": shape_id}
    except Exception as e:
        return {"error": f"Failed to add text box to slide {slide_index} in '{identifier}': {str(e)}"}

@mcp.tool()
def add_rectangle(identifier: str, slide_index: int,
                  left: float = 72, top: float = 150, width: float = 144, height: float = 72):
    """Adds a rectangle shape to a specific slide (dimensions in points)."""
    try:
        # msoShapeRectangle = 1
        shape_id = editor.add_shape(identifier, slide_index, msoShapeRectangle, left, top, width, height)
        return {"message": "Rectangle added successfully.", "shape_id": shape_id}
    except Exception as e:
        return {"error": f"Failed to add rectangle to slide {slide_index} in '{identifier}': {str(e)}"}

@mcp.tool()
def edit_element(identifier: str, slide_index: int, shape_identifier: Union[int, str],
                 properties: Dict[str, Any]):
    """Edits a shape's properties (text, left, top, width, height, name). Identify shape by ID (int) or Name (str)."""
    try:
        success = editor.edit_element(identifier, slide_index, shape_identifier, properties)
        if success:
            return {"message": f"Element '{shape_identifier}' updated successfully."}
        else:
            # Editor method already prints errors, return a generic failure
            return {"error": f"Failed to update element '{shape_identifier}'. See server logs for details."}
    except Exception as e:
        return {"error": f"Failed to edit element '{shape_identifier}' on slide {slide_index}: {str(e)}"}

@mcp.tool()
def list_shapes(identifier: str, slide_index: int):
    """Lists all shapes on a given slide with their ID, Name, and Type."""
    try:
        shapes = editor.list_shapes(identifier, slide_index)
        return {"shapes": shapes}
    except Exception as e:
        return {"error": f"Failed to list shapes on slide {slide_index} in '{identifier}': {str(e)}"}

@mcp.tool()
def find_shape_by_text(identifier: Union[str, int], slide_index: int, search_text: str, partial_match: bool = True):
    """Finds shapes on a slide containing specific text (case-insensitive). Set partial_match=False for exact match."""
    if not editor: return {"error": "PowerPoint editor not initialized."}
    try:
        matches = editor.find_shape_by_text(identifier, slide_index, search_text, partial_match)
        return {"matches": matches}
    except Exception as e:
        return _handle_tool_error("find_shape_by_text", e)

@mcp.tool()
def find_shapes_by_type(identifier: Union[str, int], slide_index: int, shape_type_name: str):
    """Finds shapes on a slide by type name (e.g., 'rectangle', 'textbox', 'picture', 'placeholder')."""
    if not editor: return {"error": "PowerPoint editor not initialized."}
    supported_types = list(SHAPE_TYPE_NAME_MAP.keys())
    if shape_type_name.lower() not in supported_types:
         return {"error": f"Unsupported shape type name '{shape_type_name}'. Try one of: {supported_types}"}
    try:
        matches = editor.find_shapes_by_type(identifier, slide_index, shape_type_name)
        return {"matches": matches}
    except Exception as e:
        return _handle_tool_error("find_shapes_by_type", e)

@mcp.tool()
def get_placeholder_shape(identifier: Union[str, int], slide_index: int, placeholder_name: str):
    """Gets a specific placeholder shape by name (e.g., 'title', 'body', 'footer', 'slidenumber')."""
    if not editor: return {"error": "PowerPoint editor not initialized."}
    supported_placeholders = list(PLACEHOLDER_NAME_MAP.keys())
    if placeholder_name.lower() not in supported_placeholders:
        return {"error": f"Unsupported placeholder name '{placeholder_name}'. Try one of: {supported_placeholders}"}
    try:
        match = editor.get_placeholder_shape(identifier, slide_index, placeholder_name)
        if match:
            return {"placeholder_found": True, "shape_info": match}
        else:
            return {"placeholder_found": False, "message": f"Placeholder '{placeholder_name}' not found on slide {slide_index}."}
    except Exception as e:
        return _handle_tool_error("get_placeholder_shape", e)


# You might want to add cleanup for COM objects, although Python's garbage collection
# combined with pywin32's handling often manages this. Explicitly setting self.app = None
# and maybe calling pythoncom.CoUninitialize() on shutdown could be added for robustness.

if __name__ == "__main__":
    print("Starting PowerPoint MCP Server (Win32)...")
    # The editor instance is created globally, attempting connection immediately.
    if editor.app is None:
         print("Warning: Failed to connect to PowerPoint on startup.")
         # Server will still run, but tools will fail until PowerPoint is available
         # and a tool call triggers a reconnect attempt.

    # Run the MCP server
    mcp.run()

# Make sure pywin32 is installed: pip install pywin32
# You might need to run `python Scripts/pywin32_postinstall.py -install`
# from your Python environment's directory if COM doesn't work initially.

# To get constants like ppSaveAsOpenXMLPresentation, run makepy:
# import win32com.client
# win32com.client.gencache.EnsureModule('{91493440-5A91-11CF-8700-00AA0060263B}', 0, 2, 12) # For Office/PPT 16.0
# Then you can use: from win32com.client import constants