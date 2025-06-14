import win32com.client
import pythoncom
import pywintypes # Import for specific exception types
from typing import List, Optional, Dict, Any, Union, Tuple
from mcp.server.fastmcp import FastMCP
import os
from win32com.client import constants
from contextlib import contextmanager

# Constants from Excel VBA Object Library (obtained via makepy or documentation)
# Using magic numbers for simplicity here, but using makepy is recommended
xlOpenXMLWorkbook = 51 # .xlsx format
xlUp = -4162
xlDown = -4121
xlToLeft = -4159
xlToRight = -4161

class ExcelEditorWin32:
    def __init__(self):
        self.app = None
        self._connect_or_launch_excel()

    def _connect_or_launch_excel(self):
        """Connects to a running instance of Excel or launches a new one."""
        try:
            # Use the Pywin32 CoInitializeEx to avoid threading issues with COM
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            self.app = win32com.client.GetActiveObject("Excel.Application")
            print("Connected to running Excel application.")
        except pywintypes.com_error:
            try:
                self.app = win32com.client.Dispatch("Excel.Application")
                self.app.Visible = True  # Make the application visible
                print("Launched new Excel application.")
            except Exception as e:
                print(f"Error launching Excel: {e}")
                self.app = None
        except Exception as e:
            print(f"An unexpected error occurred connecting to Excel: {e}")
            self.app = None
        # Make sure interaction errors are visible to the user
        if self.app:
             self.app.DisplayAlerts = True # Show Excel's own alerts

    @contextmanager
    def _performance_mode(self):
        """Context manager to optimize Excel performance during batch operations."""
        if not self.app:
            yield
            return
        
        # Store original settings
        original_screen_updating = self.app.ScreenUpdating
        original_display_alerts = self.app.DisplayAlerts
        
        try:
            # Disable UI updates for performance
            self.app.ScreenUpdating = False
            self.app.DisplayAlerts = False
            yield
        finally:
            # Always restore original settings
            self.app.ScreenUpdating = original_screen_updating
            self.app.DisplayAlerts = original_display_alerts

    def _ensure_connection(self):
        """Ensures the Excel application object is valid."""
        if self.app is None:
            self._connect_or_launch_excel()
        if self.app is None:
            raise ConnectionError("Could not connect to or launch Excel.")
        try:
            _ = self.app.Version
        except Exception as e:
            print(f"Excel connection lost or unresponsive: {e}")
            self._connect_or_launch_excel() # Try reconnecting
            if self.app is None:
                 raise ConnectionError("Could not reconnect to Excel.")

    def list_open_workbooks(self) -> List[Dict[str, Any]]:
        """Lists all currently open workbooks."""
        self._ensure_connection()
        workbooks_info = []
        try:
            if self.app.Workbooks.Count == 0:
                return []
            for i in range(1, self.app.Workbooks.Count + 1):
                wb = self.app.Workbooks(i)
                workbooks_info.append({
                    "name": wb.Name,
                    "path": wb.FullName if not wb.ReadOnly else wb.FullName + " (ReadOnly)",
                    "sheets_count": wb.Worksheets.Count,
                    "saved": wb.Saved,
                    "index": i # Provide the 1-based index for reference
                })
        except Exception as e:
            print(f"Error listing workbooks: {e}")
            if "RPC server is unavailable" in str(e):
                 self._connect_or_launch_excel()
            raise
        return workbooks_info

    def get_workbook(self, identifier: Union[str, int]) -> Optional[Any]:
        """
        Gets a workbook object by its name, path, or 1-based index.

        Args:
            identifier (Union[str, int]): The name (e.g., "Book1.xlsx"),
                              full path, or 1-based index.
        """
        self._ensure_connection()
        try:
            if isinstance(identifier, int):
                idx = identifier
                if 1 <= idx <= self.app.Workbooks.Count:
                    return self.app.Workbooks(idx)
                else:
                    print(f"Workbook index {identifier} out of range.")
                    return None
            elif isinstance(identifier, str) and identifier.isdigit():
                idx = int(identifier)
                if 1 <= idx <= self.app.Workbooks.Count:
                    return self.app.Workbooks(idx)
                else:
                    print(f"Workbook index {identifier} out of range.")
                    return None
            elif isinstance(identifier, str):
                for i in range(1, self.app.Workbooks.Count + 1):
                    wb = self.app.Workbooks(i)
                    if wb.Name.lower() == identifier.lower() or \
                       (wb.Path and wb.FullName.lower() == identifier.lower()):
                        return wb
                print(f"Workbook '{identifier}' not found.")
                return None
            else:
                print(f"Invalid workbook identifier type: {type(identifier)}. Use str or int.")
                return None
        except Exception as e:
            print(f"Error getting workbook '{identifier}': {e}")
            raise

    def save_workbook(self, identifier: Union[str, int], save_path: Optional[str] = None):
        """
        Saves the specified workbook.

        Args:
            identifier (Union[str, int]): Name, path, or index of the workbook.
            save_path (Optional[str]): Path to save to (e.g., 'C:\\MyFolder\\NewName.xlsx').
                                      If None, saves to its current path.
                                      If the workbook is new, save_path is required.
        """
        self._ensure_connection()
        wb = self.get_workbook(identifier)
        if not wb:
            raise ValueError(f"Workbook '{identifier}' not found.")

        try:
            with self._performance_mode():
                if save_path:
                    abs_path = os.path.abspath(save_path)
                    os.makedirs(os.path.dirname(abs_path), exist_ok=True)
                    # Determine file format based on extension, default to xlsx
                    file_format = None
                    if abs_path.lower().endswith(".xlsx"):
                        file_format = xlOpenXMLWorkbook # 51
                    elif abs_path.lower().endswith(".xlsm"):
                        file_format = 52 # xlOpenXMLWorkbookMacroEnabled
                    elif abs_path.lower().endswith(".xlsb"):
                        file_format = 50 # xlExcel12 (Binary)
                    elif abs_path.lower().endswith(".xls"):
                        file_format = 56 # xlExcel8
                    # Add other formats if needed (e.g., CSV = 6)

                    print(f"Attempting to save as '{abs_path}' with format {file_format}")
                    wb.SaveAs(abs_path, FileFormat=file_format)
                    print(f"Workbook saved as '{abs_path}'.")
                elif wb.Path: # Can only save if it has a path already
                    wb.Save()
                    print(f"Workbook '{wb.Name}' saved.")
                else:
                    raise ValueError("save_path is required for a new workbook that hasn't been saved before.")
        except Exception as e:
            print(f"Error saving workbook '{identifier}': {e}")
            # Provide more COM error details if possible
            if isinstance(e, pywintypes.com_error):
                 print(f"COM Error Details: HRESULT={e.hresult}, Message={e.excepinfo}")
            raise

    def list_worksheets(self, identifier: Union[str, int]) -> List[Dict[str, Any]]:
        """Lists all worksheets in the specified workbook."""
        self._ensure_connection()
        sheets_info = []
        wb = self.get_workbook(identifier)
        if not wb:
            print(f"Cannot list worksheets: Workbook '{identifier}' not found.")
            return []

        try:
            for i in range(1, wb.Worksheets.Count + 1):
                 ws = wb.Worksheets(i)
                 sheets_info.append({
                     "name": ws.Name,
                     "index": ws.Index, # 1-based index
                     "visible": ws.Visible == -1 # -1=xlSheetVisible, 0=xlSheetHidden, 2=xlSheetVeryHidden
                 })
        except Exception as e:
            print(f"Error listing worksheets in workbook '{identifier}': {e}")
            raise
        return sheets_info

    def get_worksheet(self, identifier: Union[str, int], sheet_identifier: Union[str, int]) -> Optional[Any]:
        """
        Gets a worksheet object from a workbook.

        Args:
            identifier (Union[str, int]): Workbook identifier (name, path, index).
            sheet_identifier (Union[str, int]): Worksheet identifier (name or 1-based index).

        Returns:
            Optional[Any]: The worksheet object or None if not found.
        """
        self._ensure_connection()
        wb = self.get_workbook(identifier)
        if not wb:
            return None

        try:
            # Try by index first
            if isinstance(sheet_identifier, int):
                idx = sheet_identifier
                if 1 <= idx <= wb.Worksheets.Count:
                    return wb.Worksheets(idx)
                else:
                    print(f"Worksheet index {idx} out of range for workbook '{wb.Name}'.")
                    return None
            # Try by name
            elif isinstance(sheet_identifier, str):
                 # Direct access by name is usually reliable in Excel COM
                 return wb.Worksheets(sheet_identifier)
            else:
                 print(f"Invalid sheet identifier type: {type(sheet_identifier)}. Use str or int.")
                 return None
        except pywintypes.com_error as e:
             if e.hresult == -2147352565: # Often indicates item not found
                  print(f"Worksheet '{sheet_identifier}' not found in workbook '{wb.Name}'.")
             else:
                  print(f"COM Error getting worksheet '{sheet_identifier}' from '{wb.Name}': HRESULT={e.hresult}")
             return None
        except Exception as e:
            print(f"Error getting worksheet '{sheet_identifier}' from '{identifier}': {e}")
            return None # Don't raise here, allow tool to report error

    def add_worksheet(self, identifier: Union[str, int], sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """
        Adds a new worksheet to the workbook.

        Args:
            identifier (Union[str, int]): Workbook identifier.
            sheet_name (Optional[str]): Name for the new worksheet. If None, Excel assigns a default.

        Returns:
            Dict[str, Any]: Information about the added sheet (name and index).
        """
        self._ensure_connection()
        wb = self.get_workbook(identifier)
        if not wb:
            raise ValueError(f"Workbook '{identifier}' not found.")

        try:
            with self._performance_mode():
                # Add sheet at the end
                new_sheet = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
                if sheet_name:
                     try:
                          new_sheet.Name = sheet_name
                     except Exception as name_e:
                          print(f"Warning: Could not set sheet name to '{sheet_name}'. Using default '{new_sheet.Name}'. Error: {name_e}")

                print(f"Added worksheet '{new_sheet.Name}' (Index: {new_sheet.Index}) to workbook '{wb.Name}'.")
                return {"name": new_sheet.Name, "index": new_sheet.Index}
        except Exception as e:
            print(f"Error adding worksheet to workbook '{identifier}': {e}")
            raise

    def get_cell_value(self, identifier: Union[str, int], sheet_identifier: Union[str, int], cell_address: str) -> Any:
        """
        Gets the value of a single cell.

        Args:
            identifier: Workbook identifier.
            sheet_identifier: Worksheet identifier.
            cell_address (str): Cell address (e.g., "A1", "B5").

        Returns:
            Any: The value of the cell (can be str, float, int, None, datetime, etc.).
        """
        self._ensure_connection()
        ws = self.get_worksheet(identifier, sheet_identifier)
        if not ws:
            raise ValueError(f"Worksheet '{sheet_identifier}' not found.")

        try:
            cell = ws.Range(cell_address)
            value = cell.Value
            # Convert COM dates (often floats) to Python datetime if they look like dates
            # This is heuristic - might need adjustment
            if isinstance(value, float) and value > 1 and value < 300000: # Plausible Excel date serial numbers
                 try:
                     # Excel dates are days since 1900-01-01 (or 1899-12-30 depending on settings)
                     # Using a known COM date conversion
                     dt_val = pywintypes.Time(int(value))
                     return dt_val # Returns a pywintypes time object, convertable to datetime
                 except ValueError:
                     pass # Wasn't a valid date float
            # Handle potential currency type (VT_CY) coming back as Decimal
            if type(value).__name__ == 'Decimal':
                 return float(value)
            return value
        except Exception as e:
            print(f"Error getting value from cell '{cell_address}' on sheet '{sheet_identifier}': {e}")
            raise

    def set_cell_value(self, identifier: Union[str, int], sheet_identifier: Union[str, int], cell_address: str, value: Any):
        """
        Sets the value of a single cell.

        Args:
            identifier: Workbook identifier.
            sheet_identifier: Worksheet identifier.
            cell_address (str): Cell address (e.g., "A1", "C10").
            value (Any): The value to set in the cell.
        """
        self._ensure_connection()
        ws = self.get_worksheet(identifier, sheet_identifier)
        if not ws:
            raise ValueError(f"Worksheet '{sheet_identifier}' not found.")

        try:
            cell = ws.Range(cell_address)
            cell.Value = value
            print(f"Set value of cell '{cell_address}' on sheet '{ws.Name}' to: {value}")
        except Exception as e:
            print(f"Error setting value for cell '{cell_address}' on sheet '{sheet_identifier}': {e}")
            raise

    def get_range_values(self, identifier: Union[str, int], sheet_identifier: Union[str, int], range_address: str) -> Union[Tuple[Tuple[Any, ...], ...], None]:
        """
        Gets the values from a range of cells.

        Args:
            identifier: Workbook identifier.
            sheet_identifier: Worksheet identifier.
            range_address (str): Range address (e.g., "A1:C5", "D10:D20").

        Returns:
            Union[Tuple[Tuple[Any, ...], ...], None]: A tuple of tuples containing the cell values,
                                                     or None if the range is invalid or empty.
                                                     Returns a single value if range_address is a single cell.
        """
        self._ensure_connection()
        ws = self.get_worksheet(identifier, sheet_identifier)
        if not ws:
            raise ValueError(f"Worksheet '{sheet_identifier}' not found.")

        try:
            data_range = ws.Range(range_address)
            values = data_range.Value

            # Handle single cell case
            if not isinstance(values, tuple):
                # If it's a single cell, data_range.Value returns the value directly.
                # We'll wrap it to match the expected tuple-of-tuples structure for consistency.
                 return ((values,),) # Return as tuple containing a tuple with the single value

            # Convert pywintypes time objects in the tuple to datetime
            # TODO: Need a more robust way to detect and convert dates/times/currency
            # This basic version just returns the raw tuple from COM
            return values
        except pywintypes.com_error as e:
             if e.hresult == -2146827284: # Typically invalid range address
                  print(f"Error: Invalid range address '{range_address}' on sheet '{sheet_identifier}'.")
                  raise ValueError(f"Invalid range address '{range_address}'.") from e
             else:
                  print(f"COM Error getting values from range '{range_address}' on sheet '{sheet_identifier}': {e}")
                  raise
        except Exception as e:
            print(f"Error getting values from range '{range_address}' on sheet '{sheet_identifier}': {e}")
            raise

    def set_range_values(self, identifier: Union[str, int], sheet_identifier: Union[str, int], start_cell: str, values: List[List[Any]]):
        """
        Sets values in a range of cells, starting from the specified cell.

        Args:
            identifier: Workbook identifier.
            sheet_identifier: Worksheet identifier.
            start_cell (str): The top-left cell of the range to write to (e.g., "A1").
            values (List[List[Any]]): A list of lists representing rows and columns of values.
        """
        self._ensure_connection()
        ws = self.get_worksheet(identifier, sheet_identifier)
        if not ws:
            raise ValueError(f"Worksheet '{sheet_identifier}' not found.")

        if not values or not isinstance(values, list) or not isinstance(values[0], list):
            raise ValueError("Input 'values' must be a non-empty list of lists.")

        try:
            with self._performance_mode():
                num_rows = len(values)
                num_cols = len(values[0])

                # Determine the target range based on start_cell and dimensions
                start_range = ws.Range(start_cell)
                # Use Resize property to define the target range
                target_range = start_range.Resize(num_rows, num_cols)

                # Set the values
                target_range.Value = values
                print(f"Set values in range {target_range.Address} on sheet '{ws.Name}'.")
        except Exception as e:
            print(f"Error setting values starting at cell '{start_cell}' on sheet '{sheet_identifier}': {e}")
            raise

    def find_used_ranges(self, identifier: Union[str, int], sheet_identifier: Union[str, int]) -> List[Dict[str, Any]]:
        """
        Find all contiguous data blocks in a worksheet.
        
        Returns:
            List[Dict]: List of used ranges with metadata
        """
        self._ensure_connection()
        ws = self.get_worksheet(identifier, sheet_identifier)
        if not ws:
            raise ValueError(f"Worksheet '{sheet_identifier}' not found.")

        try:
            used_range = ws.UsedRange
            if not used_range:
                return []
            
            # Get the overall used range
            main_range = {
                "range": used_range.Address,
                "rows": used_range.Rows.Count,
                "cols": used_range.Columns.Count,
                "first_row": used_range.Row,
                "last_row": used_range.Row + used_range.Rows.Count - 1,
                "first_col": used_range.Column,
                "last_col": used_range.Column + used_range.Columns.Count - 1
            }
            
            # TODO: Add logic to detect separate contiguous blocks within used range
            return [main_range]
        except Exception as e:
            print(f"Error finding used ranges on sheet '{sheet_identifier}': {e}")
            raise

    def extract_string_cells(self, identifier: Union[str, int], sheet_identifier: Union[str, int]) -> List[Dict[str, Any]]:
        """
        Extract all text cells with their metadata.
        
        Returns:
            List[Dict]: List of string cells with location info
        """
        self._ensure_connection()
        ws = self.get_worksheet(identifier, sheet_identifier)
        if not ws:
            raise ValueError(f"Worksheet '{sheet_identifier}' not found.")

        try:
            used_range = ws.UsedRange
            if not used_range:
                return []
            
            string_cells = []
            values = used_range.Value
            
            # Handle single cell case
            if not isinstance(values, tuple):
                if isinstance(values, str):
                    string_cells.append({
                        "value": values,
                        "address": used_range.Address,
                        "sheet": ws.Name,
                        "row": used_range.Row,
                        "col": used_range.Column
                    })
                return string_cells
            
            # Handle multiple cells
            for row_idx, row in enumerate(values):
                if isinstance(row, tuple):
                    for col_idx, cell_value in enumerate(row):
                        if isinstance(cell_value, str) and cell_value.strip():
                            actual_row = used_range.Row + row_idx
                            actual_col = used_range.Column + col_idx
                            cell_address = ws.Cells(actual_row, actual_col).Address
                            string_cells.append({
                                "value": cell_value.strip(),
                                "address": cell_address,
                                "sheet": ws.Name,
                                "row": actual_row,
                                "col": actual_col
                            })
                elif isinstance(row, str) and row.strip():
                    actual_row = used_range.Row + row_idx
                    actual_col = used_range.Column
                    cell_address = ws.Cells(actual_row, actual_col).Address
                    string_cells.append({
                        "value": row.strip(),
                        "address": cell_address,
                        "sheet": ws.Name,
                        "row": actual_row,
                        "col": actual_col
                    })
            
            return string_cells
        except Exception as e:
            print(f"Error extracting string cells from sheet '{sheet_identifier}': {e}")
            raise

    def get_cell_types_in_range(self, identifier: Union[str, int], sheet_identifier: Union[str, int], range_address: str) -> List[List[str]]:
        """
        Get the type of each cell in a range.
        
        Returns:
            List[List[str]]: Grid of cell types ('text', 'number', 'formula', 'empty')
        """
        self._ensure_connection()
        ws = self.get_worksheet(identifier, sheet_identifier)
        if not ws:
            raise ValueError(f"Worksheet '{sheet_identifier}' not found.")

        try:
            data_range = ws.Range(range_address)
            values = data_range.Value
            formulas = data_range.Formula
            
            # Handle single cell case
            if not isinstance(values, tuple):
                cell_type = self._determine_cell_type(values, formulas)
                return [[cell_type]]
            
            # Handle multiple cells
            cell_types = []
            for row_idx, row in enumerate(values):
                row_types = []
                if isinstance(row, tuple):
                    for col_idx, cell_value in enumerate(row):
                        formula_value = formulas[row_idx][col_idx] if isinstance(formulas[0], tuple) else formulas
                        cell_type = self._determine_cell_type(cell_value, formula_value)
                        row_types.append(cell_type)
                else:
                    formula_value = formulas[row_idx] if isinstance(formulas, tuple) else formulas
                    cell_type = self._determine_cell_type(row, formula_value)
                    row_types.append(cell_type)
                cell_types.append(row_types)
            
            return cell_types
        except Exception as e:
            print(f"Error getting cell types in range '{range_address}' on sheet '{sheet_identifier}': {e}")
            raise

    def _determine_cell_type(self, value: Any, formula: Any) -> str:
        """Determine the type of a cell based on its value and formula."""
        if formula and isinstance(formula, str) and formula.startswith('='):
            return 'formula'
        elif value is None or value == '':
            return 'empty'
        elif isinstance(value, str):
            return 'text'
        elif isinstance(value, (int, float)):
            return 'number'
        else:
            return 'other'

    def analyze_label_value_patterns(self, identifier: Union[str, int], sheet_identifier: Union[str, int]) -> Dict[str, Any]:
        """
        Analyze common DCF patterns: labels with adjacent values.
        
        Returns:
            Dict: Analysis of label-value relationships
        """
        self._ensure_connection()
        ws = self.get_worksheet(identifier, sheet_identifier)
        if not ws:
            raise ValueError(f"Worksheet '{sheet_identifier}' not found.")

        try:
            with self._performance_mode():
                string_cells = self.extract_string_cells(identifier, sheet_identifier)
                used_ranges = self.find_used_ranges(identifier, sheet_identifier)
            
            patterns = {
                "horizontal_pairs": [],  # Text in col A, numbers in col B
                "vertical_lists": [],    # Text labels with values below
                "table_headers": [],     # Horizontal headers with data below
                "summary": {
                    "total_string_cells": len(string_cells),
                    "used_ranges": len(used_ranges)
                }
            }
            
            # Analyze horizontal pairs (label-value pairs side by side)
            for string_cell in string_cells:
                adjacent_col = string_cell["col"] + 1
                try:
                    adjacent_cell = ws.Cells(string_cell["row"], adjacent_col)
                    adjacent_value = adjacent_cell.Value
                    if isinstance(adjacent_value, (int, float)) and adjacent_value != 0:
                        patterns["horizontal_pairs"].append({
                            "label": string_cell["value"],
                            "label_address": string_cell["address"],
                            "value": adjacent_value,
                            "value_address": adjacent_cell.Address,
                            "row": string_cell["row"]
                        })
                except:
                    continue
            
            # Analyze potential table headers
            for string_cell in string_cells:
                # Check if there are numbers in the same column below this text
                numbers_below = []
                for check_row in range(string_cell["row"] + 1, min(string_cell["row"] + 20, ws.UsedRange.Rows.Count + ws.UsedRange.Row)):
                    try:
                        check_cell = ws.Cells(check_row, string_cell["col"])
                        check_value = check_cell.Value
                        if isinstance(check_value, (int, float)) and check_value != 0:
                            numbers_below.append({
                                "value": check_value,
                                "address": check_cell.Address,
                                "row": check_row
                            })
                    except:
                        break
                
                if len(numbers_below) >= 2:  # At least 2 numbers below = likely a column header
                    patterns["table_headers"].append({
                        "header": string_cell["value"],
                        "header_address": string_cell["address"],
                        "data_count": len(numbers_below),
                        "data_range": f"{string_cell['address']}:{ws.Cells(numbers_below[-1]['row'], string_cell['col']).Address}"
                    })
            
            return patterns
        except Exception as e:
            print(f"Error analyzing label-value patterns on sheet '{sheet_identifier}': {e}")
            raise

# --- MCP Server Setup --- #

# Create the Excel editor instance
try:
    editor = ExcelEditorWin32()
except Exception as start_exc:
    print(f"CRITICAL: Failed to initialize ExcelEditorWin32: {start_exc}")
    editor = None # Ensure editor is None if initialization fails

# Create MCP server
mcp = FastMCP("Excel MCP (Win32)")

def _handle_excel_tool_error(tool_name: str, error: Exception) -> Dict[str, str]:
    """Standardizes error reporting for Excel tools."""
    err_msg = f"Error in tool '{tool_name}': {str(error)}"
    print(err_msg) # Log the error server-side
    if isinstance(error, ConnectionError) or "RPC server is unavailable" in str(error):
        return {"error": "Could not connect to Excel. Please ensure it is running."}
    if isinstance(error, ValueError):
         # Often used for 'not found' or invalid input errors from the editor class
         return {"error": str(error)}
    # Add specific handling for COM errors if needed
    if isinstance(error, pywintypes.com_error):
        return {"error": f"Excel Communication Error in {tool_name}: {str(error)}"}
    return {"error": err_msg}

@mcp.tool()
def list_open_workbooks():
    """Lists currently open Excel workbooks."""
    if not editor: return {"error": "Excel editor not initialized."}
    try:
        return {"workbooks": editor.list_open_workbooks()}
    except Exception as e:
        return _handle_excel_tool_error("list_open_workbooks", e)

@mcp.tool()
def save_workbook(identifier: Union[str, int], save_path: str = None):
    """Saves the specified workbook. Use index, name, or full path as identifier."""
    if not editor: return {"error": "Excel editor not initialized."}
    try:
        editor.save_workbook(identifier, save_path)
        return {"message": f"Save command issued for workbook '{identifier}' successfully."}
    except Exception as e:
        return _handle_excel_tool_error("save_workbook", e)

@mcp.tool()
def list_worksheets(identifier: Union[str, int]):
    """Lists worksheets in the specified workbook."""
    if not editor: return {"error": "Excel editor not initialized."}
    try:
        sheets = editor.list_worksheets(identifier)
        return {"worksheets": sheets}
    except Exception as e:
        return _handle_excel_tool_error("list_worksheets", e)

@mcp.tool()
def add_worksheet(identifier: Union[str, int], sheet_name: str = None):
    """Adds a worksheet to the specified workbook. Optionally provide a sheet_name."""
    if not editor: return {"error": "Excel editor not initialized."}
    try:
        new_sheet_info = editor.add_worksheet(identifier, sheet_name)
        return {"message": "Worksheet added successfully.", "sheet_info": new_sheet_info}
    except Exception as e:
        return _handle_excel_tool_error("add_worksheet", e)

@mcp.tool()
def get_cell_value(identifier: Union[str, int], sheet_identifier: Union[str, int], cell_address: str):
    """Gets the value from a specific cell (e.g., 'A1')."""
    if not editor: return {"error": "Excel editor not initialized."}
    try:
        value = editor.get_cell_value(identifier, sheet_identifier, cell_address)
        # Attempt basic serialization for common types COM might return
        if type(value).__name__ == 'datetime': # Handle pywintypes time object
            value = str(value)
        return {"value": value}
    except Exception as e:
        return _handle_excel_tool_error("get_cell_value", e)

@mcp.tool()
def set_cell_value(identifier: Union[str, int], sheet_identifier: Union[str, int], cell_address: str, value: Any):
    """Sets the value of a specific cell (e.g., 'A1')."""
    if not editor: return {"error": "Excel editor not initialized."}
    try:
        editor.set_cell_value(identifier, sheet_identifier, cell_address, value)
        return {"message": f"Successfully set cell '{cell_address}' to {value}."}
    except Exception as e:
        return _handle_excel_tool_error("set_cell_value", e)

@mcp.tool()
def get_range_values(identifier: Union[str, int], sheet_identifier: Union[str, int], range_address: str):
    """Gets values from a range (e.g., 'A1:B5'). Returns a list of lists (rows)."""
    if not editor: return {"error": "Excel editor not initialized."}
    try:
        values_tuple = editor.get_range_values(identifier, sheet_identifier, range_address)
        # Convert tuple of tuples to list of lists for JSON compatibility
        values_list = [list(row) for row in values_tuple] if values_tuple else []
        return {"values": values_list}
    except Exception as e:
        return _handle_excel_tool_error("get_range_values", e)

@mcp.tool()
def set_range_values(identifier: Union[str, int], sheet_identifier: Union[str, int], start_cell: str, values: List[List[Any]]):
    """Sets values in a range starting at start_cell. Expects 'values' as a list of lists."""
    if not editor: return {"error": "Excel editor not initialized."}
    try:
        editor.set_range_values(identifier, sheet_identifier, start_cell, values)
        num_rows = len(values)
        num_cols = len(values[0]) if num_rows > 0 else 0
        return {"message": f"Successfully set {num_rows}x{num_cols} range starting at '{start_cell}'."}
    except Exception as e:
        return _handle_excel_tool_error("set_range_values", e)

@mcp.tool()  
def find_used_ranges(identifier: Union[str, int], sheet_identifier: Union[str, int]):
    """Find all contiguous data blocks in a worksheet. Returns range metadata for DCF analysis."""
    if not editor: return {"error": "Excel editor not initialized."}
    try:
        ranges = editor.find_used_ranges(identifier, sheet_identifier)
        return {"used_ranges": ranges}
    except Exception as e:
        return _handle_excel_tool_error("find_used_ranges", e)

@mcp.tool()
def extract_string_cells(identifier: Union[str, int], sheet_identifier: Union[str, int]):
    """Extract all text cells with their metadata (address, sheet, row, col). Essential for DCF label detection."""
    if not editor: return {"error": "Excel editor not initialized."}
    try:
        string_cells = editor.extract_string_cells(identifier, sheet_identifier)
        return {"string_cells": string_cells}
    except Exception as e:
        return _handle_excel_tool_error("extract_string_cells", e)

@mcp.tool()
def get_cell_types_in_range(identifier: Union[str, int], sheet_identifier: Union[str, int], range_address: str):
    """Get cell types in a range ('text', 'number', 'formula', 'empty'). Useful for understanding data structure."""
    if not editor: return {"error": "Excel editor not initialized."}
    try:
        cell_types = editor.get_cell_types_in_range(identifier, sheet_identifier, range_address)
        return {"cell_types": cell_types}
    except Exception as e:
        return _handle_excel_tool_error("get_cell_types_in_range", e)

@mcp.tool()
def analyze_label_value_patterns(identifier: Union[str, int], sheet_identifier: Union[str, int]):
    """Analyze DCF patterns: horizontal pairs, table headers, vertical lists. Key for understanding DCF structure."""
    if not editor: return {"error": "Excel editor not initialized."}
    try:
        patterns = editor.analyze_label_value_patterns(identifier, sheet_identifier)
        return {"patterns": patterns}
    except Exception as e:
        return _handle_excel_tool_error("analyze_label_value_patterns", e)

# --- Server Execution --- #

# Optional: Add cleanup for COM objects like in the PowerPoint script
# import atexit
# def cleanup_excel_com():
#     global editor
#     if editor and editor.app:
#         # editor.app.Quit() # Careful: This closes Excel! Only use if intended.
#         editor.app = None
#         print("Released Excel application object.")
#     pythoncom.CoUninitialize()
#     print("COM Uninitialized.")
# atexit.register(cleanup_excel_com)

if __name__ == "__main__":
    print("Starting Excel MCP Server (Win32)...")
    if editor is None:
         print("CRITICAL: Excel editor could not be initialized. Server may not function correctly.")
    elif editor.app is None:
         print("Warning: Failed to connect to or launch Excel on startup.")
    else:
        print(f"Successfully connected to Excel version: {editor.app.Version}")

    # Run the MCP server
    mcp.run()

# --- Installation Notes --- #
# Make sure pywin32 is installed: uv pip install pywin32 (or pip install pywin32)
# If COM interactions fail unexpectedly after installation, you might need to run
# the post-install script from an ADMINISTRATOR command prompt:
# python C:\path\to\your\env\Scripts\pywin32_postinstall.py -install
# (Adjust path to your environment's Scripts folder)

# To get constants like xlOpenXMLWorkbook correctly (instead of magic numbers):
# 1. Run from python prompt:
#    import win32com.client
#    # Use the correct CLSID for your Excel version (this is for Excel)
#    win32com.client.gencache.EnsureModule('{00020813-0000-0000-C000-000000000046}', 0, 1, 9) # Adjust version numbers if needed (1.9 for Office 365/Excel 2016+)
# 2. Then you can use:
#    from win32com.client import constants
#    save_format = constants.xlOpenXMLWorkbook

save_format = constants.xlOpenXMLWorkbook 