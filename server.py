from mcp.server.fastmcp import FastMCP, Context
from typing import Optional, List, Dict, Any
import win32com.client
import sqlite3
import json
import random
import re

# Create server
mcp = FastMCP("AutoCAD-DB-Server")

# Initialize SQLite database
def init_db():
    try:
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        # Create table for entities
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS cad_elements (
            id INTEGER PRIMARY KEY,
            handle TEXT UNIQUE,
            name TEXT NOT NULL,
            type TEXT NOT NULL,
            layer TEXT,
            properties TEXT
        )
        ''')
        # Create table for text pattern statistics
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS text_patterns (
            id INTEGER PRIMARY KEY,
            pattern TEXT UNIQUE,
            count INTEGER DEFAULT 0,
            drawing TEXT
        )
        ''')
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Database initialization failed: {str(e)}")
        return False

# Ensure database is initialized
init_db()

# ======= AutoCAD Basic Tools =======

@mcp.tool()
def create_new_drawing(ctx: Context, template: Optional[str] = None) -> str:
    """Create a new AutoCAD drawing"""
    try:
        # Attempt to connect to AutoCAD
        acad = win32com.client.Dispatch("AutoCAD.Application")
        acad.Visible = True

        # Create a new document
        if template:
            doc = acad.Documents.Add(template)
        else:
            doc = acad.Documents.Add()

        return "Successfully created new drawing"
    except Exception as e:
        return f"Failed to create drawing: {str(e)}"

@mcp.tool()
def draw_line(ctx: Context, start_x: float, start_y: float, end_x: float, end_y: float, layer: Optional[str] = None) -> str:
    """Draw a straight line in AutoCAD.

    Args:
        start_x: X coordinate of the start point
        start_y: Y coordinate of the start point
        end_x: X coordinate of the end point
        end_y: Y coordinate of the end point
        layer: Optional layer name
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        if acad.Documents.Count == 0:
            return "No open document. Please create or open a drawing first."

        doc = acad.ActiveDocument
        model_space = doc.ModelSpace

        # If a layer is specified, switch to it or create it
        if layer:
            try:
                # Try to get the layer
                doc.Layers.Item(layer)
            except:
                # Layer does not exist — create it
                doc.Layers.Add(layer)

            # Set the current layer
            doc.ActiveLayer = doc.Layers.Item(layer)

        # Create the line
        line = model_space.AddLine(
            win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [start_x, start_y, 0]),
            win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [end_x, end_y, 0])
        )

        # Store the line information in the database
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        props = {
            "start_point": [start_x, start_y, 0],
            "end_point": [end_x, end_y, 0]
        }
        cursor.execute(
            "INSERT INTO cad_elements (handle, name, type, layer, properties) VALUES (?, ?, ?, ?, ?)",
            (line.Handle, "Line", "AcDbLine", doc.ActiveLayer.Name, json.dumps(props))
        )
        conn.commit()
        conn.close()

        return f"Line created, Handle: {line.Handle}, Layer: {doc.ActiveLayer.Name}"
    except Exception as e:
        return f"Failed to create line: {str(e)}"

@mcp.tool()
def draw_circle(ctx: Context, center_x: float, center_y: float, radius: float, layer: Optional[str] = None) -> str:
    """Draw a circle in AutoCAD.

    Args:
        center_x: X coordinate of the center point
        center_y: Y coordinate of the center point
        radius: Radius of the circle
        layer: Optional layer name
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        if acad.Documents.Count == 0:
            return "No open document. Please create or open a drawing first."

        doc = acad.ActiveDocument
        model_space = doc.ModelSpace

        # If a layer is specified, switch to it or create it
        if layer:
            try:
                # Try to get the layer
                doc.Layers.Item(layer)
            except:
                # Layer does not exist — create it
                doc.Layers.Add(layer)

            # Set the current layer
            doc.ActiveLayer = doc.Layers.Item(layer)

        # Create the circle
        center_point = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [center_x, center_y, 0])
        circle = model_space.AddCircle(center_point, radius)

        # Store the circle information in the database
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        props = {
            "center_point": [center_x, center_y, 0],
            "radius": radius
        }
        cursor.execute(
            "INSERT INTO cad_elements (handle, name, type, layer, properties) VALUES (?, ?, ?, ?, ?)",
            (circle.Handle, "Circle", "AcDbCircle", doc.ActiveLayer.Name, json.dumps(props))
        )
        conn.commit()
        conn.close()

        return f"Circle created, Handle: {circle.Handle}, Radius: {radius}, Layer: {doc.ActiveLayer.Name}"
    except Exception as e:
        return f"Failed to create circle: {str(e)}"

# ======= Entity scanning and database interaction =======

@mcp.tool()
def scan_all_entities(ctx: Context) -> str:
    """Scan all entities in the current drawing and save them to the database"""
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        if acad.Documents.Count == 0:
            return "No open document. Please create or open a drawing first."

        doc = acad.ActiveDocument
        model_space = doc.ModelSpace

        # Connect to the database
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()

        # Clear existing records (optional)
        cursor.execute("DELETE FROM cad_elements")

        # Statistics
        count = 0
        entity_types = {}

        # Iterate over all entities
        for i in range(model_space.Count):
            try:
                entity = model_space.Item(i)
                entity_type = entity.ObjectName

                # Count types
                if entity_type in entity_types:
                    entity_types[entity_type] += 1
                else:
                    entity_types[entity_type] = 1

                # Extract basic properties
                properties = {}
                if entity_type == "AcDbLine":
                    properties = {
                        "start_point": [entity.StartPoint[0], entity.StartPoint[1], entity.StartPoint[2]],
                        "end_point": [entity.EndPoint[0], entity.EndPoint[1], entity.EndPoint[2]]
                    }
                elif entity_type == "AcDbCircle":
                    properties = {
                        "center": [entity.Center[0], entity.Center[1], entity.Center[2]],
                        "radius": entity.Radius
                    }
                elif entity_type == "AcDbText" or entity_type == "AcDbMText":
                    properties = {
                        "text": entity.TextString,
                        "position": [entity.InsertionPoint[0], entity.InsertionPoint[1], entity.InsertionPoint[2]] if hasattr(entity, "InsertionPoint") else None,
                        "height": entity.Height if hasattr(entity, "Height") else None
                    }

                # Store entity information in the database
                cursor.execute(
                    "INSERT OR REPLACE INTO cad_elements (handle, name, type, layer, properties) VALUES (?, ?, ?, ?, ?)",
                    (entity.Handle, entity.ObjectName.replace("AcDb", ""), entity.ObjectName, entity.Layer, json.dumps(properties))
                )

                count += 1
            except Exception as e:
                print(f"Error processing entity {i}: {str(e)}")

        conn.commit()
        conn.close()

        # Format type summary
        type_summary = "\n".join([f"{t}: {c}" for t, c in entity_types.items()])

        return f"Scanned and saved {count} entities to the database.\n\nEntity type summary:\n{type_summary}"
    except Exception as e:
        return f"Failed to scan entities: {str(e)}"

@mcp.tool()
def highlight_entity(ctx: Context, handle: str, color: int = 1) -> str:
    """Highlight a specified entity in AutoCAD by its Handle.
    
    Args:
        handle: The entity Handle value
        color: Highlight color code (1=red, 2=yellow, 3=green, 4=cyan, 5=blue, 6=magenta)
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        if acad.Documents.Count == 0:
            return "No open document. Please create or open a drawing first."

        doc = acad.ActiveDocument

        # Create a temporary selection set
        doc.SelectionSets.Add("TempSS")
        selection = doc.SelectionSets.Item("TempSS")

        # Select entity by Handle
        filter_type = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_I2, [0])
        filter_data = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_VARIANT, ["HANDLE", handle])
        selection.Select(2, 0, 0, filter_type, filter_data)

        if selection.Count == 0:
            selection.Delete()
            return f"No entity found with Handle {handle}"

        # Change entity color
        entity = selection.Item(0)
        original_color = entity.Color
        entity.Color = color

        selection.Delete()

        return f"Highlighted entity {handle}, color changed from {original_color} to {color}"
    except Exception as e:
        return f"Failed to highlight entity: {str(e)}"

# ======= Text analysis tools =======

@mcp.tool()
def count_text_patterns(ctx: Context, pattern: str = "PMC-3M") -> str:
    """Count occurrences of a specific text pattern among text entities in the drawing.

    Args:
        pattern: The text pattern to search for (default "PMC-3M")
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        if acad.Documents.Count == 0:
            return "No open document. Please create or open a drawing first."

        doc = acad.ActiveDocument
        model_space = doc.ModelSpace
        drawing_name = doc.Name

        # Connect to the database
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()

        # Counter
        count = 0
        matching_entities = []

        # Iterate over all entities
        for i in range(model_space.Count):
            try:
                entity = model_space.Item(i)

                # Check if it's a text entity
                if hasattr(entity, "TextString"):
                    text = entity.TextString

                    # Search for the pattern
                    if pattern in text:
                        count += 1
                        matching_entities.append({
                            "handle": entity.Handle,
                            "text": text,
                            "layer": entity.Layer,
                            "position": [entity.InsertionPoint[0], entity.InsertionPoint[1]] if hasattr(entity, "InsertionPoint") else None
                        })
            except Exception as e:
                print(f"Error processing text entity {i}: {str(e)}")

        # Save statistics to the database
        cursor.execute(
            "INSERT OR REPLACE INTO text_patterns (pattern, count, drawing) VALUES (?, ?, ?)",
            (pattern, count, drawing_name)
        )
        conn.commit()
        conn.close()

        result = f"Found {count} occurrences of pattern '{pattern}' in drawing '{drawing_name}'."

        # If matches exist, show details
        if count > 0:
            details = "\n\nMatch details:"
            for i, match in enumerate(matching_entities[:10]):  # limit to first 10
                details += f"\n{i+1}. Text: '{match['text']}', Layer: {match['layer']}, Handle: {match['handle']}"

            if len(matching_entities) > 10:
                details += f"\n... and {len(matching_entities) - 10} more matches"

            result += details

        return result
    except Exception as e:
        return f"Failed to count text patterns: {str(e)}"

@mcp.tool()
def highlight_text_matches(ctx: Context, pattern: str = "PMC-3M", color: int = 1) -> str:
    """Highlight all text entities that contain the specified text pattern.

    Args:
        pattern: The text pattern to search for (default "PMC-3M")
        color: Highlight color code (1=red, 2=yellow, 3=green, 4=cyan, 5=blue, 6=magenta)
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        if acad.Documents.Count == 0:
            return "No open document. Please create or open a drawing first."

        doc = acad.ActiveDocument
        model_space = doc.ModelSpace

        # Create selection set
        try:
            # Attempt to remove any existing selection set
            doc.SelectionSets.Item("TextMatches").Delete()
        except:
            pass

        selection = doc.SelectionSets.Add("TextMatches")

        # Counter
        count = 0

        # Iterate over all entities
        for i in range(model_space.Count):
            try:
                entity = model_space.Item(i)

                # Check if it's a text entity
                if hasattr(entity, "TextString"):
                    text = entity.TextString

                    # Search for the pattern
                    if pattern in text:
                        # Save original color
                        original_color = entity.Color

                        # Change color
                        entity.Color = color

                        # Add to selection set
                        selection.AddItems([entity])

                        count += 1
            except Exception as e:
                print(f"Error processing text entity {i}: {str(e)}")

        if count > 0:
            # Zoom to selection set
            doc.ActiveView.ZoomAll()
            return f"Highlighted {count} text entities containing '{pattern}'"
        else:
            selection.Delete()
            return f"No text entities containing '{pattern}' found"
    except Exception as e:
        return f"Failed to highlight text matches: {str(e)}"

# ======= Database query tools =======

@mcp.tool()
def get_all_tables(ctx: Context) -> str:
    """Get all tables in the database"""
    try:
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()
        conn.close()

        table_list = [table[0] for table in tables]
        return json.dumps(table_list, indent=2)
    except Exception as e:
        return f"Failed to get table list: {str(e)}"

@mcp.tool()
def get_table_schema(ctx: Context, table_name: str) -> str:
    """Get schema information for a specified table"""
    try:
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        cursor.execute(f"PRAGMA table_info({table_name});")
        columns = cursor.fetchall()
        conn.close()

        schema = []
        for col in columns:
            schema.append({
                "cid": col[0],
                "name": col[1],
                "type": col[2],
                "notnull": col[3],
                "default_value": col[4],
                "pk": col[5]
            })

        return json.dumps(schema, indent=2)
    except Exception as e:
        return f"Failed to get table schema: {str(e)}"

@mcp.tool()
def execute_query(ctx: Context, query: str) -> str:
    """Execute a custom database query"""
    try:
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        cursor.execute(query)

        # If it's a SELECT query, fetch results
        if query.strip().upper().startswith("SELECT"):
            columns = [desc[0] for desc in cursor.description]
            rows = cursor.fetchall()

            # Convert results to a list of dictionaries
            result = []
            for row in rows:
                result.append(dict(zip(columns, row)))

            conn.commit()
            conn.close()
            return json.dumps(result, indent=2)
        else:
            # Non-SELECT query — return affected row count
            conn.commit()
            affected = cursor.rowcount
            conn.close()
            return f"Query executed successfully, affected {affected} rows"
    except Exception as e:
        return f"Failed to execute query: {str(e)}"

@mcp.tool()
def query_and_highlight(ctx: Context, sql_query: str, highlight_color: int = 1) -> str:
    """Highlight AutoCAD entities based on SQL query results.
    
    Args:
        sql_query: SQL query that must return a 'handle' column
        highlight_color: Highlight color code (1-255)
    """
    try:
        # Execute query
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        cursor.execute(sql_query)
        rows = cursor.fetchall()
        conn.close()
        
        if not rows:
            return "Query returned no results"
        
        # Get column names
        column_names = [description[0] for description in cursor.description]
        
        # Find 'handle' column
        handle_index = -1
        for i, name in enumerate(column_names):
            if name.lower() == 'handle':
                handle_index = i
                break
        
        if handle_index == -1:
            return "No 'handle' column found in query results"
        
        # Extract all handles
        handles = [row[handle_index] for row in rows]
        
        # Highlight entities
        acad = win32com.client.Dispatch("AutoCAD.Application")
        if acad.Documents.Count == 0:
            return "No open document. Please create or open a drawing first."
        
        doc = acad.ActiveDocument
        
        # Create selection set
        try:
            # Attempt to remove any existing selection set
            doc.SelectionSets.Item("QueryResults").Delete()
        except:
            pass
        
        selection = doc.SelectionSets.Add("QueryResults")
        
        # Highlight found entities
        highlighted_count = 0
        for handle in handles:
            try:
                # Select entity by Handle
                filter_type = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_I2, [0])
                filter_data = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_VARIANT, ["HANDLE", handle])
                
                # Create temporary selection set
                temp_selection = doc.SelectionSets.Add(f"Temp_{random.randint(1000, 9999)}")
                temp_selection.Select(2, 0, 0, filter_type, filter_data)
                
                if temp_selection.Count > 0:
                    # Change color
                    entity = temp_selection.Item(0)
                    entity.Color = highlight_color
                    
                    # Add to main selection set
                    selection.AddItems([entity])
                    highlighted_count += 1
                
                # Delete temporary selection set
                temp_selection.Delete()
            except Exception as e:
                print(f"Error processing entity {handle}: {str(e)}")
        
        if highlighted_count > 0:
            # Zoom to selection set
            doc.ActiveView.ZoomAll()
            return f"Highlighted {highlighted_count} entities (out of {len(handles)} results)"
        else:
            selection.Delete()
            return f"Failed to highlight any entities"
    except Exception as e:
        return f"Failed to query and highlight: {str(e)}"
# Added to existing code

@mcp.tool()
@mcp.tool()
def draw_line(ctx: Context, start_x: float, start_y: float, end_x: float, end_y: float, layer: Optional[str] = None) -> str:
    """Draw a straight line in AutoCAD"""
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        if acad.Documents.Count == 0:
            return "No open document. Please create or open a drawing first."

        doc = acad.ActiveDocument
        model_space = doc.ModelSpace

        # Handle layer...
        
        # Create the line
        line = model_space.AddLine(
            win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [start_x, start_y, 0]),
            win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [end_x, end_y, 0])
        )

        # Store line info in the database (without using handle field)
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        props = {
            "start_point": [start_x, start_y, 0],
            "end_point": [end_x, end_y, 0]
        }
        cursor.execute(
            "INSERT INTO cad_elements (name, type, layer, properties) VALUES (?, ?, ?, ?)",
            ("Line", "AcDbLine", doc.ActiveLayer.Name, json.dumps(props))
        )
        conn.commit()
        conn.close()

        return f"Created line from ({start_x}, {start_y}) to ({end_x}, {end_y})"
    except Exception as e:
        return f"Failed to create line: {str(e)}"

@mcp.tool()
def draw_device_connection(ctx: Context, start_device: str, end_device: str, start_x: Optional[float] = None, start_y: Optional[float] = None, end_x: Optional[float] = None, end_y: Optional[float] = None, layer: Optional[str] = None) -> str:
    """Draw a connection line between devices.

    Args:
        start_device: Label of the start device, e.g. "P14"
        end_device: Label of the end device, e.g. "P02"
        start_x: Optional start X coordinate (auto-locate if not provided)
        start_y: Optional start Y coordinate
        end_x: Optional end X coordinate
        end_y: Optional end Y coordinate
        layer: Optional layer name
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        if acad.Documents.Count == 0:
            return "No open document. Please create or open a drawing first."
        
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace
        
        # If a layer is specified, switch to it or create it
        if layer:
            try:
                # Try to get the layer
                doc.Layers.Item(layer)
            except:
                # Layer does not exist — create it
                doc.Layers.Add(layer)
            
            # Set the current layer
            doc.ActiveLayer = doc.Layers.Item(layer)

        # If coordinates are not provided, try to find devices in the database
        if start_x is None or start_y is None or end_x is None or end_y is None:
            conn = sqlite3.connect("autocad_data.db")
            cursor = conn.cursor()
            
            # Find start device
            cursor.execute(
                "SELECT properties FROM cad_elements WHERE type = 'CustomDevice' AND json_extract(properties, '$.label') = ?",
                (start_device,)
            )
            start_result = cursor.fetchone()
            
            # Find end device
            cursor.execute(
                "SELECT properties FROM cad_elements WHERE type = 'CustomDevice' AND json_extract(properties, '$.label') = ?",
                (end_device,)
            )
            end_result = cursor.fetchone()
            
            conn.close()
            
            if not start_result:
                return f"No device found with label {start_device}"
                
            if not end_result:
                return f"No device found with label {end_device}"
            
            # Parse device position and size
            start_props = json.loads(start_result[0])
            end_props = json.loads(end_result[0])
            
            start_pos = start_props["position"]
            end_pos = end_props["position"]
            
            # Set connection line start and end points (device's left-side connection point)
            start_x = start_pos[0] - 5  # left side of device
            start_y = start_pos[1]
            end_x = end_pos[0] - 5
            end_y = end_pos[1]
            
        # Create connection line (horizontal + vertical + horizontal)
        # First create starting horizontal segment
        line1_start = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [start_x, start_y, 0])
        line1_end = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [start_x - 10, start_y, 0])
        line1 = model_space.AddLine(line1_start, line1_end)
        
        # Create vertical connection
        line2_start = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [start_x - 10, start_y, 0])
        line2_end = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [start_x - 10, end_y, 0])
        line2 = model_space.AddLine(line2_start, line2_end)
        
        # Create ending horizontal segment
        line3_start = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [start_x - 10, end_y, 0])
        line3_end = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [end_x, end_y, 0])
        line3 = model_space.AddLine(line3_start, line3_end)
        
        return f"Created connection line from {start_device} to {end_device}"
    except Exception as e:
        return f"Failed to create connection line: {str(e)}"
    
# Start server
if __name__ == "__main__":
    mcp.run()