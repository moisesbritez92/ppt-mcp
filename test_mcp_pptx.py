import os
import json
from mcp_pptx_server.presentation_manager import PresentationManager
import mcp_pptx_server.presentation_reader as reader
import mcp_pptx_server.presentation_editor as editor
import mcp_pptx_server.slide_builder as builder
import mcp_pptx_server.formatting as formatting

def main():
    print("=== STARTING INTEGRATION TESTS FOR mcp-pptx-server ===")
    
    manager = PresentationManager()
    prs_id = "test_presentation"
    output_path = "test_output.pptx"
    modified_path = "test_output_modified.pptx"
    
    # Clean up old test files
    for path in [output_path, modified_path]:
        if os.path.exists(path):
            os.remove(path)
            print(f"Removed old test file: {path}")

    # 1. Create presentation
    print("\n1. Testing create_presentation...")
    entry = manager.create(prs_id, title="Main Title Test")
    prs = entry.presentation
    print(f"Presentation created. Number of slides: {len(prs.slides)}")
    assert len(prs.slides) == 1, "Should have 1 slide (the title slide)"
    assert reader.get_slide_title(prs.slides[0]) == "Main Title Test", "Title should match"
    print("✓ create_presentation test passed.")

    # 2. List layouts
    print("\n2. Testing list_layouts...")
    layouts = builder.list_layouts(prs)
    print(f"Available layouts: {len(layouts)}")
    for l in layouts[:3]:
        print(f"  - Layout {l['layout_index']}: {l['layout_name']}")
    assert len(layouts) > 0, "Should have at least one layout"
    print("✓ list_layouts test passed.")

    # 3. Add slide
    print("\n3. Testing add_slide...")
    new_slide_idx = builder.add_slide(prs, layout_index=1, title="Slide 2 Title", content="This is body content of slide 2.\nWith multiple lines.")["slide_index"]
    print(f"New slide added at index {new_slide_idx}. Total slides: {len(prs.slides)}")
    assert len(prs.slides) == 2, "Should have 2 slides"
    assert reader.get_slide_title(prs.slides[1]) == "Slide 2 Title", "Title should match"
    print("✓ add_slide test passed.")

    # 4. Add textbox
    print("\n4. Testing add_text_box...")
    tb_info = builder.add_text_box(prs, slide_index=1, text="Hello Custom TextBox\nLine 2", left_cm=2.0, top_cm=8.0, width_cm=10.0, height_cm=4.0)
    print(f"Textbox added: {tb_info}")
    assert tb_info["shape_name"] is not None, "Textbox should be named"
    print("✓ add_text_box test passed.")

    # 5. Add table
    print("\n5. Testing add_table...")
    table_data = [
        ["Month", "Revenue", "Margin"],
        ["Jan", "$10,000", "45%"],
        ["Feb", "$12,500", "48%"]
    ]
    table_info = builder.add_table(prs, slide_index=1, rows=3, cols=3, left_cm=13.0, top_cm=8.0, width_cm=10.0, height_cm=5.0, data=table_data)
    print(f"Table added: {table_info}")
    assert table_info["rows"] == 3, "Rows should match"
    assert table_info["columns"] == 3, "Columns should match"
    print("✓ add_table test passed.")

    # 6. Read slide
    print("\n6. Testing read_slide...")
    slide_details = reader.read_slide(prs, slide_index=1)
    print(f"Slide index: {slide_details['slide_index']}")
    print(f"Title: {slide_details['title']}")
    print(f"Number of shapes: {len(slide_details['shapes'])}")
    assert len(slide_details['shapes']) >= 3, "Should contain layout placeholders + custom textbox + table"
    print("✓ read_slide test passed.")

    # 7. Edit shape text and edit title
    print("\n7. Testing editing tools...")
    # Edit title
    editor.edit_slide_title(prs, slide_index=1, new_title="Updated Slide 2 Title")
    assert reader.get_slide_title(prs.slides[1]) == "Updated Slide 2 Title", "Title should have changed"
    
    # Edit textbox text
    # Let's find the textbox shape name first
    tb_shape_name = tb_info["shape_name"]
    editor.edit_shape_text(prs, slide_index=1, shape_name_or_index=tb_shape_name, new_text="Edited TextBox text")
    
    # Verify change
    updated_slide = reader.read_slide(prs, slide_index=1)
    found_edited = False
    for shape in updated_slide["shapes"]:
        if shape["name"] == tb_shape_name:
            assert shape["text"] == "Edited TextBox text", "Text should have been updated"
            found_edited = True
            break
    assert found_edited, "Edited textbox shape should exist"
    print("✓ edit_slide_title and edit_shape_text passed.")

    # 8. Set Font and Background formatting
    print("\n8. Testing formatting tools...")
    # Set background to solid custom color
    formatting.set_slide_background(prs, slide_index=1, color_hex="E6F2FF")
    
    # Set font on the edited textbox
    font_info = formatting.set_shape_font(prs, slide_index=1, shape_name_or_index=tb_shape_name, font_name="Arial", font_size=20, bold=True, italic=True, color_hex="990000")
    print(f"Font updated: {font_info}")
    assert font_info["font_name"] == "Arial", "Font family should match"
    assert font_info["font_size"] == 20, "Font size should match"
    print("✓ formatting tools passed.")

    # 9. Set and Read Speaker Notes
    print("\n9. Testing speaker notes...")
    editor.edit_slide_notes(prs, slide_index=1, new_text="Remember to highlight the growth figures in the table.")
    notes_read = reader.read_slide_notes(prs, slide_index=1)
    print(f"Notes read: '{notes_read}'")
    assert "highlight" in notes_read, "Notes text should contain key phrase"
    print("✓ speaker notes test passed.")

    # 10. Global Search and Replace
    print("\n10. Testing replace_text...")
    # Let's search and replace 'Jan' to 'January'
    replace_info = editor.replace_text(prs, search="Jan", replace="January")
    print(f"Replace result: {replace_info}")
    assert replace_info["occurrences_replaced"] >= 1, "Should have replaced at least 1 occurrence"
    
    # Search text
    search_results = reader.search_text(prs, query="January")
    print(f"Search result for 'January': {search_results}")
    assert len(search_results) >= 1, "Should have found 'January' in table or shape text"
    print("✓ search_text and replace_text passed.")

    # 11. Save presentation and read structure
    print("\n11. Testing save and open lifecycle...")
    manager.save(prs_id, file_path=output_path)
    assert os.path.exists(output_path), "File should be saved to disk"
    print(f"Saved file size: {os.path.getsize(output_path)} bytes")
    
    # Close it
    manager.close(prs_id)
    # Ensure it's closed
    try:
        manager.get(prs_id)
        assert False, "Should raise KeyError since it was closed"
    except KeyError:
        pass
    
    # Open from disk
    open_id = "test_opened"
    manager.open(open_id, file_path=output_path)
    entry_open = manager.get(open_id)
    prs_open = entry_open.presentation
    assert len(prs_open.slides) == 2, "Opened presentation should have 2 slides"
    print("✓ save/close/open lifecycle passed.")

    # 12. Duplicate, Move and Delete slide
    print("\n12. Testing slide organization...")
    # Add a third slide first
    builder.add_slide(prs_open, title="Slide 3 Title", content="This is slide 3.")
    assert len(prs_open.slides) == 3
    
    # Duplicate Slide 1 (index 1)
    dup_info = editor.duplicate_slide(prs_open, slide_index=1)
    print(f"Duplicated slide: {dup_info}")
    assert len(prs_open.slides) == 4, "Should have 4 slides after duplication"
    assert reader.get_slide_title(prs_open.slides[3]) == "Updated Slide 2 Title", "Duplicated slide title should match source"
    
    # Move Slide from 3 to 1
    # Slide at 3 is the duplicated one. Let's move it to index 1.
    move_info = editor.move_slide(prs_open, from_index=3, to_index=1)
    print(f"Moved slide: {move_info}")
    assert reader.get_slide_title(prs_open.slides[1]) == "Updated Slide 2 Title", "Slide 1 should now be the moved slide"
    
    # Delete slide at index 2
    delete_info = editor.delete_slide(prs_open, slide_index=2)
    print(f"Deleted slide: {delete_info}")
    assert len(prs_open.slides) == 3, "Should have 3 slides after deletion"
    print("✓ slide duplication, moving, and deletion passed.")

    # 13. Save modified and verify structure
    print("\n13. Testing final save and full text read...")
    manager.save(open_id, file_path=modified_path)
    assert os.path.exists(modified_path), "Modified file should be saved"
    
    full_text = reader.read_full_text(prs_open)
    print("\n--- Final Presentation Full Text Output ---")
    for s in full_text:
        print(f"Slide {s['slide_index']}: {s['title']}")
        for elem in s['text_elements']:
            print(f"  [{elem['shape_name']}]: {elem['text'].replace(chr(10), ' | ')}")
        if s['notes']:
            print(f"  Notes: {s['notes']}")
    print("-------------------------------------------")
    
    manager.close(open_id)
    print("\n✓ ALL INTEGRATION TESTS PASSED SUCCESSFULLY!")

if __name__ == "__main__":
    main()
