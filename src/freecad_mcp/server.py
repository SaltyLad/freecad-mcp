import json
import logging
import xmlrpc.client
from contextlib import asynccontextmanager
from typing import AsyncIterator, Dict, Any, Literal

from mcp.server.fastmcp import FastMCP, Context
from mcp.types import TextContent, ImageContent

# Configure logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger("FreeCADMCPserver")


_only_text_feedback = False
_rpc_host = "localhost"


class FreeCADConnection:
    def __init__(self, host: str = "localhost", port: int = 9875):
        self.server = xmlrpc.client.ServerProxy(f"http://{host}:{port}", allow_none=True)

    def disconnect(self):
        """Clean up the connection (no-op for XML-RPC)."""
        pass

    def ping(self) -> bool:
        return self.server.ping()

    def create_document(self, name: str) -> dict[str, Any]:
        return self.server.create_document(name)

    def create_object(self, doc_name: str, obj_data: dict[str, Any]) -> dict[str, Any]:
        return self.server.create_object(doc_name, obj_data)

    def edit_object(self, doc_name: str, obj_name: str, obj_data: dict[str, Any]) -> dict[str, Any]:
        return self.server.edit_object(doc_name, obj_name, obj_data)

    def delete_object(self, doc_name: str, obj_name: str) -> dict[str, Any]:
        return self.server.delete_object(doc_name, obj_name)

    def insert_part_from_library(self, relative_path: str) -> dict[str, Any]:
        return self.server.insert_part_from_library(relative_path)

    def execute_code(self, code: str) -> dict[str, Any]:
        return self.server.execute_code(code)

    def get_active_screenshot(self, view_name: str = "Isometric", width: int | None = None, height: int | None = None, focus_object: str | None = None) -> str | None:
        try:
            # The addon RPC server already checks view compatibility
            # and returns None if screenshots aren't supported
            result = self.server.get_active_screenshot(view_name, width, height, focus_object)
            return result if result else None
        except Exception as e:
            logger.error(f"Error getting screenshot: {e}")
            return None

    def save_document(self, doc_name: str, file_path: str = "") -> dict[str, Any]:
        return self.server.save_document(doc_name, file_path)

    def export_document(self, doc_name: str, obj_names: list[str], file_path: str, file_format: str) -> dict[str, Any]:
        return self.server.export_document(doc_name, obj_names, file_path, file_format)

    def recompute_document(self, doc_name: str) -> dict[str, Any]:
        return self.server.recompute_document(doc_name)

    def get_objects(self, doc_name: str) -> list[dict[str, Any]]:
        return self.server.get_objects(doc_name)

    def get_object(self, doc_name: str, obj_name: str) -> dict[str, Any]:
        return self.server.get_object(doc_name, obj_name)

    def get_parts_list(self) -> list[str]:
        return self.server.get_parts_list()

    def inspect_geometry(self, doc_name: str, obj_name: str, what: str = "summary") -> dict[str, Any]:
        return self.server.inspect_geometry(doc_name, obj_name, what)

    def list_documents(self) -> list[str]:
        return self.server.list_documents()


@asynccontextmanager
async def server_lifespan(server: FastMCP) -> AsyncIterator[Dict[str, Any]]:
    try:
        logger.info("FreeCADMCP server starting up")
        try:
            _ = get_freecad_connection()
            logger.info("Successfully connected to FreeCAD on startup")
        except Exception as e:
            logger.warning(f"Could not connect to FreeCAD on startup: {str(e)}")
            logger.warning(
                "Make sure the FreeCAD addon is running before using FreeCAD resources or tools"
            )
        yield {}
    finally:
        # Clean up the global connection on shutdown
        global _freecad_connection
        if _freecad_connection:
            logger.info("Disconnecting from FreeCAD on shutdown")
            _freecad_connection.disconnect()
            _freecad_connection = None
        logger.info("FreeCADMCP server shut down")


mcp = FastMCP(
    "FreeCADMCP",
    instructions="FreeCAD integration through the Model Context Protocol",
    lifespan=server_lifespan,
)


_freecad_connection: FreeCADConnection | None = None


def get_freecad_connection():
    """Get or create a persistent FreeCAD connection"""
    global _freecad_connection
    if _freecad_connection is None:
        _freecad_connection = FreeCADConnection(host=_rpc_host, port=9875)
        if not _freecad_connection.ping():
            logger.error("Failed to ping FreeCAD")
            _freecad_connection = None
            raise Exception(
                "Failed to connect to FreeCAD. Make sure the FreeCAD addon is running."
            )
    return _freecad_connection


# Helper function to safely add screenshot to response
def add_screenshot_if_available(response, screenshot, force_include=False):
    """Safely add screenshot to response only if it's available.

    Args:
        response: The response list to append to.
        screenshot: Base64-encoded screenshot or None.
        force_include: If True, include screenshot even when _only_text_feedback is set.
    """
    should_include = force_include or not _only_text_feedback
    if screenshot is not None and should_include:
        response.append(ImageContent(type="image", data=screenshot, mimeType="image/png"))
    elif should_include and screenshot is None:
        response.append(TextContent(
            type="text",
            text="Note: Visual preview is unavailable in the current view type. "
                 "Switch to a 3D view to see visual feedback."
        ))
    return response


@mcp.tool()
def create_document(ctx: Context, name: str) -> list[TextContent]:
    """Create a new document in FreeCAD.

    Args:
        name: The name of the document to create.

    Returns:
        A message indicating the success or failure of the document creation.

    Examples:
        If you want to create a document named "MyDocument", you can use the following data.
        ```json
        {
            "name": "MyDocument"
        }
        ```
    """
    freecad = get_freecad_connection()
    try:
        res = freecad.create_document(name)
        if res["success"]:
            return [
                TextContent(type="text", text=f"Document '{res['document_name']}' created successfully")
            ]
        else:
            return [
                TextContent(type="text", text=f"Failed to create document: {res['error']}")
            ]
    except Exception as e:
        logger.error(f"Failed to create document: {str(e)}")
        return [
            TextContent(type="text", text=f"Failed to create document: {str(e)}")
        ]


@mcp.tool()
def create_object(
    ctx: Context,
    doc_name: str,
    obj_type: str,
    obj_name: str,
    analysis_name: str | None = None,
    obj_properties: dict[str, Any] = None,
) -> list[TextContent | ImageContent]:
    """Create a new object in FreeCAD.
    Object types start with "Part::", "Draft::", "PartDesign::", or "Fem::".

    Args:
        doc_name: The name of the document to create the object in.
        obj_type: The type of the object to create.
        obj_name: The name of the object to create.
        analysis_name: Optional FEM analysis name to add the object to.
        obj_properties: The properties of the object to create.

    Returns:
        A message indicating the success or failure of the object creation.

    Common types: Part::Box, Part::Cylinder, Part::Sphere, Part::Cone, Part::Torus,
    PartDesign::Body, Draft::Circle, Draft::Wire, Fem::AnalysisPython, Fem::FemMeshGmsh.

    Properties support Placement (Base + Rotation), ViewObject (ShapeColor), and
    object references (Base, Tool, Source, Profile as string names, References as
    list of [object_name, face] pairs). For FEM mesh, 'Part' property is required.
    """
    freecad = get_freecad_connection()
    try:
        obj_data = {"Name": obj_name, "Type": obj_type, "Properties": obj_properties or {}, "Analysis": analysis_name}
        res = freecad.create_object(doc_name, obj_data)
        screenshot = freecad.get_active_screenshot() if not _only_text_feedback else None
        
        if res["success"]:
            response = [
                TextContent(type="text", text=f"Object '{res['object_name']}' created successfully"),
            ]
            return add_screenshot_if_available(response, screenshot)
        else:
            response = [
                TextContent(type="text", text=f"Failed to create object: {res['error']}"),
            ]
            return add_screenshot_if_available(response, screenshot)
    except Exception as e:
        logger.error(f"Failed to create object: {str(e)}")
        return [
            TextContent(type="text", text=f"Failed to create object: {str(e)}")
        ]


@mcp.tool()
def edit_object(
    ctx: Context, doc_name: str, obj_name: str, obj_properties: dict[str, Any]
) -> list[TextContent | ImageContent]:
    """Edit an object in FreeCAD.
    This tool is used when the `create_object` tool cannot handle the object creation.

    Args:
        doc_name: The name of the document to edit the object in.
        obj_name: The name of the object to edit.
        obj_properties: The properties of the object to edit.

    Returns:
        A message indicating the success or failure of the object editing and a screenshot of the object.
    """
    freecad = get_freecad_connection()
    try:
        res = freecad.edit_object(doc_name, obj_name, {"Properties": obj_properties})
        screenshot = freecad.get_active_screenshot() if not _only_text_feedback else None

        if res["success"]:
            response = [
                TextContent(type="text", text=f"Object '{res['object_name']}' edited successfully"),
            ]
            return add_screenshot_if_available(response, screenshot)
        else:
            response = [
                TextContent(type="text", text=f"Failed to edit object: {res['error']}"),
            ]
            return add_screenshot_if_available(response, screenshot)
    except Exception as e:
        logger.error(f"Failed to edit object: {str(e)}")
        return [
            TextContent(type="text", text=f"Failed to edit object: {str(e)}")
        ]


@mcp.tool()
def delete_object(ctx: Context, doc_name: str, obj_name: str) -> list[TextContent | ImageContent]:
    """Delete an object in FreeCAD.

    Args:
        doc_name: The name of the document to delete the object from.
        obj_name: The name of the object to delete.

    Returns:
        A message indicating the success or failure of the object deletion and a screenshot of the object.
    """
    freecad = get_freecad_connection()
    try:
        res = freecad.delete_object(doc_name, obj_name)
        screenshot = freecad.get_active_screenshot() if not _only_text_feedback else None
        
        if res["success"]:
            response = [
                TextContent(type="text", text=f"Object '{res['object_name']}' deleted successfully"),
            ]
            return add_screenshot_if_available(response, screenshot)
        else:
            response = [
                TextContent(type="text", text=f"Failed to delete object: {res['error']}"),
            ]
            return add_screenshot_if_available(response, screenshot)
    except Exception as e:
        logger.error(f"Failed to delete object: {str(e)}")
        return [
            TextContent(type="text", text=f"Failed to delete object: {str(e)}")
        ]


@mcp.tool()
def execute_code(ctx: Context, code: str, include_screenshot: bool = False) -> list[TextContent | ImageContent]:
    """Execute arbitrary Python code in FreeCAD.

    Args:
        code: The Python code to execute.
        include_screenshot: If True, include a screenshot in the response even when text-only mode is enabled.

    Returns:
        A message indicating the success or failure of the code execution and its output.
    """
    freecad = get_freecad_connection()
    try:
        res = freecad.execute_code(code)
        screenshot = freecad.get_active_screenshot() if (include_screenshot or not _only_text_feedback) else None

        if res["success"]:
            response = [
                TextContent(type="text", text=f"Code executed successfully: {res['message']}"),
            ]
            return add_screenshot_if_available(response, screenshot, force_include=include_screenshot)
        else:
            response = [
                TextContent(type="text", text=f"Failed to execute code: {res['error']}"),
            ]
            return add_screenshot_if_available(response, screenshot, force_include=include_screenshot)
    except Exception as e:
        logger.error(f"Failed to execute code: {str(e)}")
        return [
            TextContent(type="text", text=f"Failed to execute code: {str(e)}")
        ]


@mcp.tool()
def get_view(ctx: Context, view_name: Literal["Isometric", "Front", "Top", "Right", "Back", "Left", "Bottom", "Dimetric", "Trimetric"], width: int | None = None, height: int | None = None, focus_object: str | None = None) -> list[ImageContent | TextContent]:
    """Get a screenshot of the active view.

    Args:
        view_name: The name of the view to get the screenshot of.
        The following views are available:
        - "Isometric"
        - "Front"
        - "Top"
        - "Right"
        - "Back"
        - "Left"
        - "Bottom"
        - "Dimetric"
        - "Trimetric"
        width: The width of the screenshot in pixels. If not specified, uses the viewport width.
        height: The height of the screenshot in pixels. If not specified, uses the viewport height.
        focus_object: The name of the object to focus on. If not specified, fits all objects in the view.

    Returns:
        A screenshot of the active view.
    """
    freecad = get_freecad_connection()
    screenshot = freecad.get_active_screenshot(view_name, width, height, focus_object)
    
    if screenshot is not None:
        return [ImageContent(type="image", data=screenshot, mimeType="image/png")]
    else:
        return [TextContent(type="text", text="Cannot get screenshot in the current view type (such as TechDraw or Spreadsheet)")]


@mcp.tool()
def insert_part_from_library(ctx: Context, relative_path: str) -> list[TextContent | ImageContent]:
    """Insert a part from the parts library addon.

    Args:
        relative_path: The relative path of the part to insert.

    Returns:
        A message indicating the success or failure of the part insertion and a screenshot of the object.
    """
    freecad = get_freecad_connection()
    try:
        res = freecad.insert_part_from_library(relative_path)
        screenshot = freecad.get_active_screenshot() if not _only_text_feedback else None
        
        if res["success"]:
            response = [
                TextContent(type="text", text=f"Part inserted from library: {res['message']}"),
            ]
            return add_screenshot_if_available(response, screenshot)
        else:
            response = [
                TextContent(type="text", text=f"Failed to insert part from library: {res['error']}"),
            ]
            return add_screenshot_if_available(response, screenshot)
    except Exception as e:
        logger.error(f"Failed to insert part from library: {str(e)}")
        return [
            TextContent(type="text", text=f"Failed to insert part from library: {str(e)}")
        ]


@mcp.tool()
def get_objects(ctx: Context, doc_name: str, summary: bool = True) -> list[TextContent | ImageContent]:
    """Get all objects in a document.

    Args:
        doc_name: The name of the document to get the objects from.
        summary: If True (default), return only Name, Label, TypeId and Placement for each object.
                 If False, return full object details including all properties and shape info.

    Returns:
        A list of objects in the document.
    """
    freecad = get_freecad_connection()
    try:
        objects = freecad.get_objects(doc_name)
        if summary:
            objects = [
                {
                    "Name": obj.get("Name"),
                    "Label": obj.get("Label"),
                    "TypeId": obj.get("TypeId"),
                    "Placement": obj.get("Placement"),
                }
                for obj in objects
            ]
        screenshot = freecad.get_active_screenshot() if not _only_text_feedback else None
        response = [
            TextContent(type="text", text=json.dumps(objects)),
        ]
        return add_screenshot_if_available(response, screenshot)
    except Exception as e:
        logger.error(f"Failed to get objects: {str(e)}")
        return [
            TextContent(type="text", text=f"Failed to get objects: {str(e)}")
        ]


@mcp.tool()
def get_object(ctx: Context, doc_name: str, obj_name: str) -> list[TextContent | ImageContent]:
    """Get an object from a document with all its properties.

    Args:
        doc_name: The name of the document to get the object from.
        obj_name: The name of the object to get.

    Returns:
        The object's full property data.
    """
    freecad = get_freecad_connection()
    try:
        screenshot = freecad.get_active_screenshot() if not _only_text_feedback else None
        response = [
            TextContent(type="text", text=json.dumps(freecad.get_object(doc_name, obj_name))),
        ]
        return add_screenshot_if_available(response, screenshot)
    except Exception as e:
        logger.error(f"Failed to get object: {str(e)}")
        return [
            TextContent(type="text", text=f"Failed to get object: {str(e)}")
        ]


@mcp.tool()
def save_document(ctx: Context, doc_name: str, file_path: str = "") -> list[TextContent]:
    """Save a FreeCAD document.

    Args:
        doc_name: The name of the document to save.
        file_path: Optional file path to save to. If empty, saves to the existing file path.

    Returns:
        A message indicating the success or failure of the save operation.
    """
    freecad = get_freecad_connection()
    try:
        res = freecad.save_document(doc_name, file_path)
        if res["success"]:
            return [TextContent(type="text", text=f"Document '{doc_name}' saved to {res['file_path']}")]
        else:
            return [TextContent(type="text", text=f"Failed to save document: {res['error']}")]
    except Exception as e:
        logger.error(f"Failed to save document: {str(e)}")
        return [TextContent(type="text", text=f"Failed to save document: {str(e)}")]


@mcp.tool()
def export_objects(
    ctx: Context,
    doc_name: str,
    obj_names: list[str],
    file_path: str,
    file_format: str,
) -> list[TextContent]:
    """Export objects from a FreeCAD document to STEP, STL, or IGES format.

    Args:
        doc_name: The name of the document containing the objects.
        obj_names: List of object names to export.
        file_path: The file path to export to (e.g. '/tmp/part.step').
        file_format: The export format: 'STEP', 'STL', or 'IGES'.

    Returns:
        A message indicating the success or failure of the export.
    """
    freecad = get_freecad_connection()
    try:
        res = freecad.export_document(doc_name, obj_names, file_path, file_format)
        if res["success"]:
            return [TextContent(type="text", text=f"Exported to {res['file_path']}")]
        else:
            return [TextContent(type="text", text=f"Failed to export: {res['error']}")]
    except Exception as e:
        logger.error(f"Failed to export: {str(e)}")
        return [TextContent(type="text", text=f"Failed to export: {str(e)}")]


@mcp.tool()
def recompute(ctx: Context, doc_name: str) -> list[TextContent]:
    """Force recompute of a FreeCAD document.

    Args:
        doc_name: The name of the document to recompute.

    Returns:
        A message indicating the success or failure of the recompute.
    """
    freecad = get_freecad_connection()
    try:
        res = freecad.recompute_document(doc_name)
        if res["success"]:
            return [TextContent(type="text", text=f"Document '{doc_name}' recomputed successfully")]
        else:
            return [TextContent(type="text", text=f"Failed to recompute: {res['error']}")]
    except Exception as e:
        logger.error(f"Failed to recompute: {str(e)}")
        return [TextContent(type="text", text=f"Failed to recompute: {str(e)}")]


@mcp.tool()
def inspect_geometry(
    ctx: Context,
    doc_name: str,
    obj_name: str,
    what: Literal["summary", "faces", "edges", "sketches", "feature_tree", "all"] = "summary",
) -> list[TextContent]:
    """Inspect structured geometry data from a FreeCAD object. Returns precise numeric data
    (bounding box, face normals, edge lengths, sketch geometry, feature tree) for verification
    without the token cost of a screenshot.

    Args:
        doc_name: The name of the document.
        obj_name: The name of the object to inspect.
        what: What data to return:
            - "summary" (default): bounding box, volume, area, face/edge/vertex counts, validity
            - "faces": per-face center, normal, area, surface type
            - "edges": per-edge midpoint, length, curve type
            - "sketches": sketch geometry, constraints, placement (obj must be a sketch)
            - "feature_tree": feature list with validity and key properties (obj should be a Body)
            - "all": everything above combined

    Returns:
        Structured geometry data as JSON text.
    """
    freecad = get_freecad_connection()
    try:
        res = freecad.inspect_geometry(doc_name, obj_name, what)
        if res["success"]:
            return [TextContent(type="text", text=json.dumps(res["data"], indent=2))]
        else:
            return [TextContent(type="text", text=f"Failed to inspect geometry: {res['error']}")]
    except Exception as e:
        logger.error(f"Failed to inspect geometry: {str(e)}")
        return [TextContent(type="text", text=f"Failed to inspect geometry: {str(e)}")]


@mcp.tool()
def get_parts_list(ctx: Context) -> list[TextContent]:
    """Get the list of parts in the parts library addon.
    """
    freecad = get_freecad_connection()
    parts = freecad.get_parts_list()
    if parts:
        return [
            TextContent(type="text", text=json.dumps(parts))
        ]
    else:
        return [
            TextContent(type="text", text=f"No parts found in the parts library. You must add parts_library addon.")
        ]


@mcp.tool()
def list_documents(ctx: Context) -> list[TextContent]:
    """Get the list of open documents in FreeCAD.

    Returns:
        A list of document names.
    """
    freecad = get_freecad_connection()
    docs = freecad.list_documents()
    return [TextContent(type="text", text=json.dumps(docs))]


@mcp.prompt()
def asset_creation_strategy() -> str:
    return """
Asset Creation Strategy for FreeCAD MCP

When creating content in FreeCAD, always follow these steps:

0. Before starting any task, always use get_objects() to confirm the current state of the document.

1. Utilize the parts library:
   - Check available parts using get_parts_list().
   - If the required part exists in the library, use insert_part_from_library() to insert it into your document.

2. If the appropriate asset is not available in the parts library:
   - Create basic shapes (e.g., cubes, cylinders, spheres) using create_object().
   - Adjust and define detailed properties of the shapes as necessary using edit_object().

3. Always assign clear and descriptive names to objects when adding them to the document.

4. Explicitly set the position, scale, and rotation properties of created or inserted objects using edit_object() to ensure proper spatial relationships.

5. After editing an object, always verify that the set properties have been correctly applied by using get_object().

6. If detailed customization or specialized operations are necessary, use execute_code() to run custom Python scripts.

Only revert to basic creation methods in the following cases:
- When the required asset is not available in the parts library.
- When a basic shape is explicitly requested.
- When creating complex shapes requires custom scripting.
"""


def _validate_host(value: str) -> str:
    """Validate that *value* is a valid IP address or hostname.

    Used as the ``type`` callback for the ``--host`` argparse argument.
    Raises ``argparse.ArgumentTypeError`` on invalid input.
    """
    import argparse

    import validators

    if validators.ipv4(value) or validators.ipv6(value) or validators.hostname(value):
        return value
    raise argparse.ArgumentTypeError(
        f"Invalid host: '{value}'. Must be a valid IP address or hostname."
    )


def main():
    """Run the MCP server"""
    global _only_text_feedback, _rpc_host
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--only-text-feedback", action="store_true", help="Only return text feedback")
    parser.add_argument("--host", type=_validate_host, default="localhost", help="Host address of the FreeCAD RPC server to connect to (default: localhost)")
    args = parser.parse_args()
    _only_text_feedback = args.only_text_feedback
    _rpc_host = args.host
    logger.info(f"Only text feedback: {_only_text_feedback}")
    logger.info(f"Connecting to FreeCAD RPC server at: {_rpc_host}")
    mcp.run()