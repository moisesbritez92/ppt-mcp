import json
import logging
from typing import Optional, List, Union, Any
from mcp.server.fastmcp import FastMCP

from mcp_pptx_server.presentation_manager import PresentationManager
import mcp_pptx_server.presentation_reader as reader
import mcp_pptx_server.presentation_editor as editor
import mcp_pptx_server.slide_builder as builder
import mcp_pptx_server.formatting as formatting

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("mcp-pptx-server")

mcp = FastMCP("pptx-editor")
manager = PresentationManager()

# ==========================================
# 1. Lifecycle Tools
# ==========================================

@mcp.tool()
def create_presentation(prs_id: str, title: str = "") -> str:
    """Creates a new empty presentation in memory.

    Args:
        prs_id: Unique identifier for the presentation.
        title: Optional title for the initial title slide.
    """
    try:
        manager.create(prs_id, title)
        return json.dumps({
            "success": True,
            "message": f"Presentation '{prs_id}' created in memory successfully.",
            "total_slides": 1 if title else 0
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)

@mcp.tool()
def open_presentation(prs_id: str, file_path: str) -> str:
    """Opens an existing PowerPoint (.pptx) file from disk.

    Args:
        prs_id: Unique identifier to assign to this presentation in memory.
        file_path: Absolute or relative path to the .pptx file.
    """
    try:
        entry = manager.open(prs_id, file_path)
        return json.dumps({
            "success": True,
            "message": f"Presentation '{prs_id}' loaded from '{file_path}' successfully.",
            "slides_count": len(entry.presentation.slides)
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)

@mcp.tool()
def save_presentation(prs_id: str, file_path: Optional[str] = None) -> str:
    """Saves an open presentation to disk.

    Args:
        prs_id: Identifier of the open presentation to save.
        file_path: Optional path to save to (saves to original path if not specified).
    """
    try:
        saved_path = manager.save(prs_id, file_path)
        return json.dumps({
            "success": True,
            "message": f"Presentation '{prs_id}' saved successfully.",
            "saved_path": saved_path
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)

@mcp.tool()
def close_presentation(prs_id: str) -> str:
    """Closes a presentation, freeing it from memory.

    Args:
        prs_id: Identifier of the presentation to close.
    """
    try:
        manager.close(prs_id)
        return json.dumps({
            "success": True,
            "message": f"Presentation '{prs_id}' closed and freed from memory."
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)

@mcp.tool()
def list_presentations() -> str:
    """Lists all PowerPoint presentations currently open in memory."""
    try:
        open_list = manager.list_presentations()
        return json.dumps({
            "success": True,
            "presentations": open_list
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)


# ==========================================
# 2. Reading Tools
# ==========================================

@mcp.tool()
def get_presentation_structure(prs_id: str) -> str:
    """Gets the hierarchical layout structure of the presentation, listing slides, layouts, and shape counts.

    Args:
        prs_id: Identifier of the presentation.
    """
    try:
        entry = manager.get(prs_id)
        structure = reader.get_presentation_structure(entry.presentation)
        return json.dumps({
            "success": True,
            "structure": structure
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)

@mcp.tool()
def read_slide(prs_id: str, slide_index: int) -> str:
    """Returns detailed information about a slide, including shapes, types, positions, texts, tables, and notes.

    Args:
        prs_id: Identifier of the presentation.
        slide_index: 0-based index of the slide to read.
    """
    try:
        entry = manager.get(prs_id)
        slide_info = reader.read_slide(entry.presentation, slide_index)
        return json.dumps({
            "success": True,
            "slide": slide_info
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)

@mcp.tool()
def read_slide_notes(prs_id: str, slide_index: int) -> str:
    """Reads the speaker notes of a specific slide.

    Args:
        prs_id: Identifier of the presentation.
        slide_index: 0-based index of the slide.
    """
    try:
        entry = manager.get(prs_id)
        notes = reader.read_slide_notes(entry.presentation, slide_index)
        return json.dumps({
            "success": True,
            "slide_index": slide_index,
            "notes": notes
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)

@mcp.tool()
def read_full_text(prs_id: str) -> str:
    """Extracts and organizes all textual content (from shapes, tables, and notes) across all slides.

    Args:
        prs_id: Identifier of the presentation.
    """
    try:
        entry = manager.get(prs_id)
        full_text = reader.read_full_text(entry.presentation)
        return json.dumps({
            "success": True,
            "slides_text": full_text
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)

@mcp.tool()
def list_slides(prs_id: str) -> str:
    """Returns a quick summary list of all slides in the presentation.

    Args:
        prs_id: Identifier of the presentation.
    """
    try:
        entry = manager.get(prs_id)
        slides_list = reader.list_slides(entry.presentation)
        return json.dumps({
            "success": True,
            "slides": slides_list
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)

@mcp.tool()
def search_text(prs_id: str, query: str, case_sensitive: bool = False) -> str:
    """Searches for a text term across shapes, tables, and speaker notes on all slides.

    Args:
        prs_id: Identifier of the presentation.
        query: Text term to search for.
        case_sensitive: Whether the search should be case sensitive.
    """
    try:
        entry = manager.get(prs_id)
        results = reader.search_text(entry.presentation, query, case_sensitive)
        return json.dumps({
            "success": True,
            "query": query,
            "matches_count": len(results),
            "results": results
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)


# ==========================================
# 3. Slide Editing Tools
# ==========================================

@mcp.tool()
def edit_shape_text(prs_id: str, slide_index: int, shape_name_or_index: Union[str, int], new_text: str) -> str:
    """Edits/replaces the text of a specific shape on a slide.

    Args:
        prs_id: Identifier of the presentation.
        slide_index: 0-based index of the slide containing the shape.
        shape_name_or_index: Name (e.g., 'TextBox 1'), ID, or 0-based index of the shape to edit.
        new_text: New text content to assign. Supports multiple lines.
    """
    try:
        entry = manager.get(prs_id)
        result = editor.edit_shape_text(entry.presentation, slide_index, shape_name_or_index, new_text)
        return json.dumps({
            "success": True,
            "message": f"Shape '{shape_name_or_index}' on slide {slide_index} updated successfully.",
            "details": result
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)

@mcp.tool()
def edit_slide_title(prs_id: str, slide_index: int, new_title: str) -> str:
    """Edits the title of a slide.

    Args:
        prs_id: Identifier of the presentation.
        slide_index: 0-based index of the slide.
        new_title: New title text.
    """
    try:
        entry = manager.get(prs_id)
        result = editor.edit_slide_title(entry.presentation, slide_index, new_title)
        return json.dumps({
            "success": True,
            "message": f"Slide {slide_index} title updated successfully.",
            "details": result
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)

@mcp.tool()
def edit_slide_notes(prs_id: str, slide_index: int, new_text: str) -> str:
    """Edits or sets the speaker/presenter notes of a slide.

    Args:
        prs_id: Identifier of the presentation.
        slide_index: 0-based index of the slide.
        new_text: New speaker notes.
    """
    try:
        entry = manager.get(prs_id)
        result = editor.edit_slide_notes(entry.presentation, slide_index, new_text)
        return json.dumps({
            "success": True,
            "message": f"Slide {slide_index} notes updated successfully.",
            "details": result
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)

@mcp.tool()
def replace_text(prs_id: str, search: str, replace: str, case_sensitive: bool = False) -> str:
    """Performs a find-and-replace operation across all text boxes, tables, and notes.

    Args:
        prs_id: Identifier of the presentation.
        search: Text term to search for.
        replace: Text to replace the search term with.
        case_sensitive: Whether the replacement should be case sensitive.
    """
    try:
        entry = manager.get(prs_id)
        result = editor.replace_text(entry.presentation, search, replace, case_sensitive)
        return json.dumps({
            "success": True,
            "message": f"Global replace completed.",
            "details": result
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)

@mcp.tool()
def delete_slide(prs_id: str, slide_index: int) -> str:
    """Deletes a slide from the presentation.

    Args:
        prs_id: Identifier of the presentation.
        slide_index: 0-based index of the slide to delete.
    """
    try:
        entry = manager.get(prs_id)
        result = editor.delete_slide(entry.presentation, slide_index)
        return json.dumps({
            "success": True,
            "message": f"Slide {slide_index} deleted successfully.",
            "details": result
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)

@mcp.tool()
def duplicate_slide(prs_id: str, slide_index: int) -> str:
    """Duplicates a slide, copying its text box layouts, tables, and pictures.

    Args:
        prs_id: Identifier of the presentation.
        slide_index: 0-based index of the slide to duplicate.
    """
    try:
        entry = manager.get(prs_id)
        result = editor.duplicate_slide(entry.presentation, slide_index)
        return json.dumps({
            "success": True,
            "message": f"Slide {slide_index} duplicated successfully.",
            "details": result
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)

@mcp.tool()
def move_slide(prs_id: str, from_index: int, to_index: int) -> str:
    """Reorders slides by moving a slide from one position index to another.

    Args:
        prs_id: Identifier of the presentation.
        from_index: 0-based source index of the slide.
        to_index: 0-based target index to place the slide.
    """
    try:
        entry = manager.get(prs_id)
        result = editor.move_slide(entry.presentation, from_index, to_index)
        return json.dumps({
            "success": True,
            "message": f"Slide moved from {from_index} to {to_index} successfully.",
            "details": result
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)


# ==========================================
# 4. Content Creation Tools
# ==========================================

@mcp.tool()
def add_slide(prs_id: str, layout_index: int = 1, title: str = "", content: str = "") -> str:
    """Adds a new slide with optional title and body content.

    Args:
        prs_id: Identifier of the presentation.
        layout_index: Index of the slide layout to use (list layouts using list_layouts tool).
        title: Optional title string to write into the slide.
        content: Optional body text to write into the content placeholder or textbox.
    """
    try:
        entry = manager.get(prs_id)
        result = builder.add_slide(entry.presentation, layout_index, title, content)
        return json.dumps({
            "success": True,
            "message": "Slide added successfully.",
            "details": result
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)

@mcp.tool()
def add_text_box(prs_id: str, slide_index: int, text: str, left_cm: float, top_cm: float, width_cm: float, height_cm: float) -> str:
    """Adds a textbox to a slide with specified centimeter dimensions and coordinates.

    Args:
        prs_id: Identifier of the presentation.
        slide_index: 0-based index of the slide.
        text: Text content of the new textbox.
        left_cm: Distance from left of slide in centimeters.
        top_cm: Distance from top of slide in centimeters.
        width_cm: Width of text box in centimeters.
        height_cm: Height of text box in centimeters.
    """
    try:
        entry = manager.get(prs_id)
        result = builder.add_text_box(entry.presentation, slide_index, text, left_cm, top_cm, width_cm, height_cm)
        return json.dumps({
            "success": True,
            "message": "Textbox added successfully.",
            "details": result
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)

@mcp.tool()
def add_image(prs_id: str, slide_index: int, image_path: str, left_cm: float, top_cm: float, width_cm: Optional[float] = None, height_cm: Optional[float] = None) -> str:
    """Inserts an image into a slide at centimeter dimensions and coordinates.

    Args:
        prs_id: Identifier of the presentation.
        slide_index: 0-based index of the slide.
        image_path: Absolute or relative path to the image file on disk.
        left_cm: Distance from left of slide in centimeters.
        top_cm: Distance from top of slide in centimeters.
        width_cm: Optional width of image in centimeters. If omitted, scales automatically.
        height_cm: Optional height of image in centimeters. If omitted, scales automatically.
    """
    try:
        entry = manager.get(prs_id)
        result = builder.add_image(entry.presentation, slide_index, image_path, left_cm, top_cm, width_cm, height_cm)
        return json.dumps({
            "success": True,
            "message": "Image added successfully.",
            "details": result
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)

@mcp.tool()
def add_table(prs_id: str, slide_index: int, rows: int, cols: int, left_cm: float, top_cm: float, width_cm: float, height_cm: float, data: Optional[List[List[Any]]] = None) -> str:
    """Adds a structured grid table to a slide.

    Args:
        prs_id: Identifier of the presentation.
        slide_index: 0-based index of the slide.
        rows: Number of table rows.
        cols: Number of table columns.
        left_cm: Distance from left of slide in centimeters.
        top_cm: Distance from top of slide in centimeters.
        width_cm: Width of table in centimeters.
        height_cm: Height of table in centimeters.
        data: Optional list-of-lists containing initial text values for each cell (row major order).
    """
    try:
        entry = manager.get(prs_id)
        result = builder.add_table(entry.presentation, slide_index, rows, cols, left_cm, top_cm, width_cm, height_cm, data)
        return json.dumps({
            "success": True,
            "message": "Table added successfully.",
            "details": result
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)

@mcp.tool()
def list_layouts(prs_id: str) -> str:
    """Lists all available slide layout indexes and names from the presentation template.

    Args:
        prs_id: Identifier of the presentation.
    """
    try:
        entry = manager.get(prs_id)
        layouts_list = builder.list_layouts(entry.presentation)
        return json.dumps({
            "success": True,
            "layouts": layouts_list
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)


# ==========================================
# 5. Formatting Tools
# ==========================================

@mcp.tool()
def set_shape_font(prs_id: str, slide_index: int, shape_name_or_index: Union[str, int], font_name: str = "", font_size: Optional[float] = None, bold: Optional[bool] = None, italic: Optional[bool] = None, color_hex: str = "") -> str:
    """Applies font styles (family, size, bold, italic, color) to text within a shape.

    Args:
        prs_id: Identifier of the presentation.
        slide_index: 0-based index of the slide containing the shape.
        shape_name_or_index: Name, ID, or 0-based index of the shape to format.
        font_name: Name of font family (e.g., 'Arial', 'Calibri') (optional).
        font_size: Size of font in points (e.g., 18.5) (optional).
        bold: Apply bold style (boolean) (optional).
        italic: Apply italic style (boolean) (optional).
        color_hex: Apply text color in hex format (e.g., 'FF0000' or '#FF0000') (optional).
    """
    try:
        entry = manager.get(prs_id)
        result = formatting.set_shape_font(entry.presentation, slide_index, shape_name_or_index, font_name, font_size, bold, italic, color_hex)
        return json.dumps({
            "success": True,
            "message": f"Shape '{shape_name_or_index}' font updated successfully.",
            "details": result
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)

@mcp.tool()
def set_slide_background(prs_id: str, slide_index: int, color_hex: str) -> str:
    """Sets a solid color background for a slide.

    Args:
        prs_id: Identifier of the presentation.
        slide_index: 0-based index of the slide.
        color_hex: Color in hex format (e.g., 'FFFFFF', '#000000').
    """
    try:
        entry = manager.get(prs_id)
        result = formatting.set_slide_background(entry.presentation, slide_index, color_hex)
        return json.dumps({
            "success": True,
            "message": f"Slide {slide_index} background color updated successfully.",
            "details": result
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)


def main():
    """Main entry point to start the FastMCP server."""
    logger.info("Starting pptx-editor MCP Server...")
    mcp.run()

if __name__ == "__main__":
    main()
