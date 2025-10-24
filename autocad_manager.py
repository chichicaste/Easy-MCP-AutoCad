from contextlib import contextmanager
import os
import tempfile
import win32com.client
from pyautocad import APoint


class AutoCADManager:
    def __init__(self):
        """Initialize the AutoCAD connection manager"""
        self.acad = None
        self.doc = None

    def connect(self):
        """Connect to the AutoCAD application"""
        try:
            self.acad = win32com.client.Dispatch("AutoCAD.Application")
            # If AutoCAD is available, ensure its window is visible
            if self.acad:
                self.acad.Visible = True

            # Get the current document or create a new one
            if self.acad.Documents.Count == 0:
                self.doc = self.acad.Documents.Add()
            else:
                self.doc = self.acad.ActiveDocument

            return True
        except Exception as e:
            print(f"Failed to connect to AutoCAD: {str(e)}")
            return False

    def disconnect(self):
        """Disconnect from AutoCAD"""
        self.doc = None
        self.acad = None

    @contextmanager
    def autocad_session(self):
        """Context manager for an AutoCAD session"""
        success = self.connect()
        if not success:
            raise Exception("Unable to connect to AutoCAD")
        try:
            yield self
        finally:
            self.disconnect()

    def create_new_drawing(self, template=None):
        """Create a new AutoCAD drawing"""
        try:
            if template:
                new_doc = self.acad.Documents.Add(template)
            else:
                new_doc = self.acad.Documents.Add()

            self.doc = new_doc
            return "Successfully created new drawing"
        except Exception as e:
            return f"Failed to create new drawing: {str(e)}"

    def get_all_layers(self):
        """Retrieve information for all layers"""
        layers = []
        for i in range(self.doc.Layers.Count):
            layer = self.doc.Layers.Item(i)
            layers.append({
                "name": layer.Name,
                "color": layer.Color,
                "linetype": layer.Linetype,
                "is_locked": layer.Lock,
                "is_on": layer.LayerOn,
                "is_frozen": layer.Freeze
            })
        return layers

    def create_layer(self, layer_name, color=7):
        """Create a new layer"""
        try:
            # Check whether the layer already exists
            try:
                existing_layer = self.doc.Layers.Item(layer_name)
                return f"Layer '{layer_name}' already exists"
            except:
                # Layer does not exist — create it
                new_layer = self.doc.Layers.Add(layer_name)
                new_layer.Color = color
                return f"Successfully created layer '{layer_name}'"
        except Exception as e:
            return f"Failed to create layer: {str(e)}"

    def delete_layer(self, layer_name):
        """Delete a layer"""
        try:
            layer = self.doc.Layers.Item(layer_name)
            # Check whether the layer is the current active layer
            if layer.Name == self.doc.ActiveLayer.Name:
                return "Cannot delete the current active layer"
            layer.Delete()
            return f"Successfully deleted layer '{layer_name}'"
        except Exception as e:
            return f"Failed to delete layer: {str(e)}"

    def draw_line(self, start_x, start_y, end_x, end_y, layer=None):
        """Draw a line on the specified layer"""
        try:
            # Save the current active layer
            current_layer = self.doc.ActiveLayer.Name

            # If a layer was specified, switch to it
            if layer:
                try:
                    self.doc.ActiveLayer = self.doc.Layers.Item(layer)
                except:
                    return f"Layer '{layer}' does not exist"

            # Draw the line
            start_point = APoint(start_x, start_y)
            end_point = APoint(end_x, end_y)
            self.doc.ModelSpace.AddLine(start_point, end_point)

            # Restore the original active layer
            if layer:
                self.doc.ActiveLayer = self.doc.Layers.Item(current_layer)

            return "Successfully drew line"
        except Exception as e:
            return f"Failed to draw line: {str(e)}"

    def export_drawing(self, file_path, file_type="DWG"):
        """Export the current drawing to the specified format"""
        try:
            # Ensure the destination directory exists
            os.makedirs(os.path.dirname(file_path), exist_ok=True)

            # If the file exists already, remove it first
            if os.path.exists(file_path):
                os.remove(file_path)

            # Choose export method based on file type
            if file_type.upper() == "DWG":
                self.doc.SaveAs(file_path)
            elif file_type.upper() == "DXF":
                self.doc.Export(file_path, "DXF")
            elif file_type.upper() == "PDF":
                self.doc.Export(file_path, "PDF")
            else:
                return f"Unsupported file type: {file_type}"
                
            return f"Successfully exported drawing to {file_path}"
        except Exception as e:
            return f"Failed to export drawing: {str(e)}"

    def get_entity_stats(self):
        """Get statistics of entity types present in the drawing"""
        entity_counts = {}

        # Count entities in model space
        for entity in self.doc.ModelSpace:
            entity_type = entity.ObjectName
            if entity_type in entity_counts:
                entity_counts[entity_type] += 1
            else:
                entity_counts[entity_type] = 1

        return entity_counts