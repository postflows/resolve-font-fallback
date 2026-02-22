# ================================================
# Font Fallback
# Part of PostFlows toolkit for DaVinci Resolve
# https://github.com/postflows
# ================================================

"""
Font Fallback v1.4
Advanced font management tool for DaVinci Resolve with comprehensive fallback and restoration system

Original by nizar / version 1.0,
Modified by Sergey Knyazkov:
- Changed UI to use built-in ui.dispatcher in DaVinci Resolve
- Added ability to replace missing fonts in text blocks
- Added configurable font replacement settings with UI selection
- Fixed font style detection to work correctly with DaVinci Resolve API
- Added proper missing style detection
- Added user interface for font/style selection
- Improved UI layout and styling
- Added font restoration system with embedded tags and JSON logging
- Fixed parent detection to handle Text+ nodes with effects correctly
- Added MultiText node support with individual text block processing
- Enhanced restoration system with text block-level tracking
- Improved error handling and debugging capabilities
- Added detailed logging with session tracking

Original repository: https://github.com/neezr/Text-Toolbox-for-DaVinci-Resolve

MAIN FEATURES:
1. Font Detection & Analysis
   - Scans current timeline for all used fonts
   - Identifies missing fonts and unavailable styles
   - Supports both Text+ and MultiText nodes
   - Real-time font availability checking

2. Font Replacement
   - Batch replace missing fonts with user-selected alternatives
   - Configurable replacement font and style
   - Preserves original font information for restoration
   - Handles complex node hierarchies and effects

3. Restoration System
   - Embedded restoration tags in node comments
   - JSON logging of all replacement operations
   - Session-based tracking with unique IDs
   - One-click restore to original fonts
   - Text block-level restoration for MultiText nodes

4. MultiText Support
   - Individual font detection per text block
   - Granular replacement and restoration
   - Text block identification (Text1, Text2, etc.)
   - Preserves styling within complex text compositions

5. User Interface
   - Modern, responsive UI design
   - Real-time font style updates
   - Status reporting and progress tracking
   - Clipboard integration for missing fonts list

USAGE:
- Run from DaVinci Resolve (Workspace > Scripts)
- Select replacement font and style from dropdowns
- Click "Refresh Timeline" to scan for missing fonts
- Use "Replace Missing" to apply font replacements
- Use "Restore Fonts" to revert to original fonts
- "Copy Missing" exports missing fonts list to clipboard

NOTES:
- Legacy Text (not Text+) is not supported due to API limitations
- Requires `pyperclip` for reliable clipboard copying (install via `pip install pyperclip`)
- Restoration requires original fonts to be installed on the system
- MultiText nodes are fully supported with individual text block processing
- All operations are logged to desktop for audit purposes

TECHNICAL DETAILS:
- Uses Fusion API for font management
- Implements non-destructive font replacement
- Maintains backward compatibility with previous versions
- Handles complex node structures and parent groups
- Robust error handling for production use
"""

import os
import json
import re
from datetime import datetime

project = resolve.GetProjectManager().GetCurrentProject()
timeline = project.GetCurrentTimeline()

try:
    import pyperclip
except ImportError:
    pyperclip = None
    print("Warning: `pyperclip` module not found. Clipboard copying may not work reliably. Install it via `pip install pyperclip`.")

# Default font replacement settings (fallback values)
DEFAULT_REPLACEMENT_FONT = "Open Sans"
DEFAULT_REPLACEMENT_STYLE = "Regular"

# Restoration system templates
RESTORE_TAG_TEMPLATE = """
[PostFlows_FONT_RESTORE]
original_font: {original_font}
original_style: {original_style}
replaced_with: {replacement_font}|{replacement_style}
timestamp: {timestamp}
restore_id: {restore_id}
[/PostFlows_FONT_RESTORE]"""

def generate_unique_id():
    """Generate unique ID for restoration session"""
    from datetime import datetime
    return datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]

def get_installed_fonts():
    """
    Retrieve installed fonts with their available styles.
    
    Returns:
        dict: Dictionary with font names as keys and sets of available styles as values.
    """
    try:
        font_list = fusion.FontManager.GetFontList()
        installed_fonts = {}
        
        for font_name, font_data in font_list.items():
            # Get styles as dictionary keys
            styles = [key for key in font_data.keys() if isinstance(key, str)]
            installed_fonts[font_name] = set(styles)
        
        return installed_fonts
    except Exception as e:
        print(f"Error getting installed fonts: {e}")
        return {}

def check_font_style_availability(font_name, style_name):
    """
    Check availability of a specific font style.
    
    Args:
        font_name (str): Font name
        style_name (str): Style name
    
    Returns:
        bool: True if style is available, False otherwise
    """
    try:
        font_list = fusion.FontManager.GetFontList()
        
        if font_name not in font_list:
            return False
            
        available_styles = list(font_list[font_name].keys())
        return style_name in available_styles
        
    except Exception as e:
        print(f"Error checking font style: {e}")
        return False

def get_font_styles(font_name):
    """
    Retrieve available styles for a given font.
    
    Args:
        font_name (str): Name of the font to query styles for.
    
    Returns:
        list: Sorted list of available styles for the font.
    """
    try:
        font_list = fusion.FontManager.GetFontList()
        if font_name in font_list:
            styles = [key for key in font_list[font_name].keys() if isinstance(key, str)]
            return sorted(styles)
        return ["Regular"]
    except Exception as e:
        print(f"Error getting font styles for {font_name}: {e}")
        return ["Regular"]

def get_used_fonts():
    """
    Collect fonts used in the current timeline's video tracks with style validation.
    Now supports MultiText nodes with multiple text blocks.
    """
    used_fonts = {}
    installed_fonts = get_installed_fonts()
    
    if not timeline:
        print("No timeline is currently open!")
        return used_fonts
    
    for j in range(1, timeline.GetTrackCount("video") + 1):
        for tl_item in timeline.GetItemListInTrack("video", j):
            for k in range(1, tl_item.GetFusionCompCount() + 1):
                comp = tl_item.GetFusionCompByIndex(k)
                if comp:
                    for node in comp.GetToolList().values():
                        # Check for regular Text+ nodes
                        if hasattr(node, 'Font') and node.Font:
                            try:
                                font_name = node.Font[1]
                                try:
                                    font_style = node.Style[1]
                                except:
                                    font_style = "Regular"
                                
                                process_font_usage(font_name, font_style, used_fonts, installed_fonts)
                            except Exception as e:
                                print(f"Error processing Text+ node {node.Name}: {e}")
                        
                        # Check for MultiText nodes
                        elif hasattr(node, 'GetAttrs'):
                            try:
                                attrs = node.GetAttrs()
                                node_id = attrs.get('TOOLS_RegID', '')
                                
                                if node_id == 'MultiText':
                                    extract_multitext_fonts(node, used_fonts, installed_fonts)
                            except Exception as e:
                                print(f"Error checking MultiText node: {e}")
    
    return used_fonts

def extract_multitext_fonts(node, used_fonts, installed_fonts):
    """Alternative approach to extract fonts from MultiText node"""
    try:
        # Try a different approach - manually check for common input patterns
        text_blocks = {}
        
        # Common MultiText input patterns
        common_patterns = [
            "Text1.Font", "Text1.Style",
            "Text2.Font", "Text2.Style", 
            "Text3.Font", "Text3.Style",
            "Text4.Font", "Text4.Style",
            "Text5.Font", "Text5.Style"
        ]
        
        for pattern in common_patterns:
            try:
                if '.Font' in pattern:
                    text_block = pattern.split('.Font')[0]
                    font_value = node.GetInput(pattern)
                    if font_value:
                        if text_block not in text_blocks:
                            text_blocks[text_block] = {}
                        text_blocks[text_block]['font'] = font_value
                
                elif '.Style' in pattern:
                    text_block = pattern.split('.Style')[0]
                    style_value = node.GetInput(pattern)
                    if text_block not in text_blocks:
                        text_blocks[text_block] = {}
                    text_blocks[text_block]['style'] = style_value if style_value else "Regular"
                    
            except Exception as e:
                print(f"Error checking pattern {pattern}: {e}")
        
        # Process each text block
        for text_block, font_info in text_blocks.items():
            if 'font' in font_info:
                font_name = font_info['font']
                font_style = font_info.get('style', 'Regular')
                
                process_font_usage(font_name, font_style, used_fonts, installed_fonts)
                
    except Exception as e:
        print(f"Error in alternative MultiText extraction: {e}")
        import traceback
        traceback.print_exc()

def process_font_usage(font_name, font_style, used_fonts, installed_fonts):
    """Helper function to process font usage (shared by Text+ and MultiText)"""
    if font_name not in used_fonts:
        used_fonts[font_name] = {
            'used_styles': set(),
            'missing_styles': set(),
            'font_missing': font_name not in installed_fonts
        }
    
    used_fonts[font_name]['used_styles'].add(font_style)
    
    # Check style availability
    if font_name in installed_fonts:
        if font_style not in installed_fonts[font_name]:
            used_fonts[font_name]['missing_styles'].add(font_style)

def get_selected_replacement_font():
    """Get user-selected replacement font and style"""
    try:
        font = itm["FontCombo"].CurrentText
        style = itm["StyleCombo"].CurrentText
        return font, style
    except:
        return DEFAULT_REPLACEMENT_FONT, DEFAULT_REPLACEMENT_STYLE

def update_style_combo(ev):
    """Update style list when font selection changes"""
    try:
        selected_font = itm["FontCombo"].CurrentText
        if selected_font:
            available_styles = get_font_styles(selected_font)
            itm["StyleCombo"].Clear()
            for style in available_styles:
                itm["StyleCombo"].AddItem(style)
            
            # Set Regular as default if available
            if "Regular" in available_styles:
                regular_index = available_styles.index("Regular")
                itm["StyleCombo"].CurrentIndex = regular_index
            elif available_styles:
                itm["StyleCombo"].CurrentIndex = 0
                
    except Exception as e:
        print(f"Error updating style combo: {e}")

# Restoration system functions
def create_restoration_log():
    """Create a new restoration log"""
    return {
        "project_name": project.GetName() if project else "Unknown",
        "timeline_name": timeline.GetName() if timeline else "Unknown", 
        "timestamp": datetime.now().isoformat(),
        "restore_session_id": generate_unique_id(),
        "replacements": []
    }

def add_replacement_to_log(log, node, original_font, original_style, replacement_font, replacement_style):
    """Add a replacement entry to the log"""
    log["replacements"].append({
        "node_name": node.Name,
        "original_font": original_font,
        "original_style": original_style,
        "replacement_font": replacement_font,
        "replacement_style": replacement_style,
        "timestamp": datetime.now().isoformat()
    })

def save_restoration_log(log):
    """Save restoration log to file"""
    try:
        # Save to user's desktop or home directory
        home_path = os.path.expanduser("~")
        desktop_path = os.path.join(home_path, "Desktop")
        log_path = desktop_path if os.path.exists(desktop_path) else home_path
        
        filename = f"PostFlows_Font_Restoration_Log_{log['restore_session_id']}.json"
        filepath = os.path.join(log_path, filename)
        
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(log, f, indent=2, ensure_ascii=False)
        
        print(f"Restoration log saved to: {filepath}")
        return filepath
    except Exception as e:
        print(f"Error saving restoration log: {e}")
        return None

def create_restore_tag(original_font, original_style, replacement_font, replacement_style, restore_id):
    """Create restoration tag for embedding in node comments with better formatting"""
    return RESTORE_TAG_TEMPLATE.format(
        original_font=original_font,
        original_style=original_style,
        replacement_font=replacement_font,
        replacement_style=replacement_style,
        timestamp=datetime.now().isoformat(),
        restore_id=restore_id
    ).strip()

def parse_restore_tag_from_comments(node):
    """Extract restoration data from node comments"""
    try:
        comments = node.GetInput("Comments") or ""
        
        if "[PostFlows_FONT_RESTORE]" in comments and "[/PostFlows_FONT_RESTORE]" in comments:
            start_tag = comments.find("[PostFlows_FONT_RESTORE]")
            end_tag = comments.find("[/PostFlows_FONT_RESTORE]")
            
            if start_tag != -1 and end_tag != -1:
                restore_section = comments[start_tag:end_tag + len("[/PostFlows_FONT_RESTORE]")]
                
                # Parse parameters using regex
                original_font = extract_tag_value(restore_section, "original_font")
                original_style = extract_tag_value(restore_section, "original_style")
                
                if original_font and original_style:
                    return {
                        "original_font": original_font,
                        "original_style": original_style,
                        "restore_section": restore_section
                    }
    except Exception as e:
        print(f"Error parsing restore tag from {node.Name}: {e}")
    
    return None

def extract_tag_value(text, param_name):
    """Extract parameter value from tag"""
    pattern = f"{param_name}: (.+)"
    match = re.search(pattern, text)
    return match.group(1).strip() if match else None

def remove_restore_tag_from_comments(node):
    """Remove restoration tag from node comments"""
    try:
        comments = node.GetInput("Comments") or ""
        
        if "[PostFlows_FONT_RESTORE]" in comments and "[/PostFlows_FONT_RESTORE]" in comments:
            start_tag = comments.find("[PostFlows_FONT_RESTORE]")
            end_tag = comments.find("[/PostFlows_FONT_RESTORE]") + len("[/PostFlows_FONT_RESTORE]")
            
            if start_tag != -1 and end_tag != -1:
                # Remove the tag section
                new_comments = comments[:start_tag] + comments[end_tag:]
                # Clean up extra newlines
                new_comments = re.sub(r'\n\s*\n\s*\n', '\n\n', new_comments).strip()
                
                node.SetInput("Comments", new_comments)
                print(f"Removed restore tag from {node.Name}")
    except Exception as e:
        print(f"Error removing restore tag from {node.Name}: {e}")

def find_parent_group_or_macro(comp, target_node):
    """
    Find parent GroupOperator or MacroOperator for the given node.
    Ignores effect nodes like Glow, Transform, etc.
    
    Args:
        comp: Fusion composition
        target_node: Node to find parent for
    
    Returns:
        Parent node (GroupOperator/MacroOperator) or None
    """
    try:
        all_nodes = comp.GetToolList()
        
        # Look specifically for GroupOperator and MacroOperator
        valid_parents = []
        for node_name, node in all_nodes.items():
            try:
                node_attrs = node.GetAttrs()
                node_id = node_attrs.get('TOOLS_RegID', '')
                
                # Only accept actual Groups and Macros, not effect nodes
                if node_id == 'GroupOperator' or node_id == 'MacroOperator':
                    valid_parents.append((node_name, node, node_id))
                elif 'Group' in node_id or 'Macro' in node_id:
                    valid_parents.append((node_name, node, node_id))
            except Exception as e:
                print(f"Error checking node {node_name}: {e}")
                continue
        
        if valid_parents:
            # Return the first valid parent found
            return valid_parents[0][1]
                
    except Exception as e:
        print(f"Error finding parent group/macro: {e}")
    
    return None

def should_use_node_comments(comp, node):
    """
    Determine if we should write restore tags directly to the node
    instead of looking for a parent.
    
    Returns True if node is a standalone Text+ or has only effect nodes around it.
    """
    try:
        all_nodes = comp.GetToolList()
        
        # Count actual parent containers (not effect nodes)
        parent_containers = 0
        for node_name, check_node in all_nodes.items():
            try:
                node_attrs = check_node.GetAttrs()
                node_id = node_attrs.get('TOOLS_RegID', '')
                
                if node_id in ['GroupOperator', 'MacroOperator'] or ('Group' in node_id and 'Group' != node_id):
                    parent_containers += 1
            except:
                pass
        
        # If no real parent containers, use node itself
        return parent_containers == 0
        
    except Exception as e:
        print(f"Error in should_use_node_comments: {e}")
        return True  # Default to using node comments

def restore_original_fonts(ev):
    """Restore original fonts from embedded tags with proper style matching"""
    if not timeline:
        print("No timeline is currently open!")
        itm["StatusLabel"].Text = "No timeline open"
        return
    
    found_tags = 0
    restored_count = 0
    skipped_unavailable = 0
    
    for j in range(1, timeline.GetTrackCount("video") + 1):
        for tl_item in timeline.GetItemListInTrack("video", j):
            for k in range(1, tl_item.GetFusionCompCount() + 1):
                comp = tl_item.GetFusionCompByIndex(k)
                if comp:
                    comp.Lock()
                    
                    # Collect all restoration tags from all nodes
                    all_nodes = comp.GetToolList()
                    nodes_with_tags = {}
                    
                    # Collect all restoration tags
                    for node_name, node in all_nodes.items():
                        try:
                            tags = parse_all_restore_tags_from_comments(node)
                            if tags:
                                nodes_with_tags[node_name] = {
                                    'node': node,
                                    'tags': tags,
                                    'node_type': node.GetAttrs().get('TOOLS_RegID', 'Unknown') if hasattr(node, 'GetAttrs') else 'Text+'
                                }
                        except Exception as e:
                            print(f"Error checking comments in {node_name}: {e}")
                    
                    # Restore fonts in Text+ nodes
                    for node_name, node in all_nodes.items():
                        try:
                            node_type = node.GetAttrs().get('TOOLS_RegID', 'Unknown') if hasattr(node, 'GetAttrs') else 'Text+'
                            
                            if node_type == 'TextPlus' and hasattr(node, 'Font') and node.Font:
                                current_font = node.Font[1]
                                current_style = node.Style[1] if node.Style else "Regular"
                                
                                # Find tag matching current node by replacement_font
                                matching_tag = find_matching_restore_tag(node_name, current_font, current_style, nodes_with_tags)
                                
                                if matching_tag and 'tag' in matching_tag:
                                    found_tags += 1
                                    tag = matching_tag['tag']
                                    original_font = tag.get("original_font")
                                    original_style = tag.get("original_style")
                                    
                                    if original_font and original_style:
                                        if check_font_style_availability(original_font, original_style):
                                            try:
                                                node.SetInput("Font", original_font)
                                                node.SetInput("Style", original_style)
                                                
                                                # Remove used tag
                                                remove_specific_restore_tag(matching_tag['source_node'], tag)
                                                
                                                restored_count += 1
                                                print(f"✓ Restored {node_name}: {original_font} {original_style}")
                                            except Exception as e:
                                                print(f"✗ Failed to restore {node_name}: {e}")
                                                skipped_unavailable += 1
                                        else:
                                            skipped_unavailable += 1
                                            print(f"✗ Cannot restore {node_name}: {original_font} {original_style} not available")
                                    else:
                                        print(f"⚠ Warning: Incomplete tag data for {node_name}")
                            
                            # Process MultiText nodes
                            elif node_type == 'MultiText':
                                multitext_restored, multitext_skipped = restore_multitext_fonts(node, nodes_with_tags)
                                restored_count += multitext_restored
                                skipped_unavailable += multitext_skipped
                                
                        except Exception as e:
                            print(f"Error processing node {node_name}: {e}")
                    
                    comp.Unlock()
    
    # Build result message
    if found_tags == 0:
        status_msg = "No restoration tags found"
    elif restored_count == 0:
        status_msg = f"Found {found_tags} restoration tags, but original fonts not available"
    elif skipped_unavailable == 0:
        status_msg = f"Restored all {restored_count} fonts successfully"
    else:
        status_msg = f"Restored {restored_count} fonts, {skipped_unavailable} still unavailable"
    
    itm["StatusLabel"].Text = status_msg
    print(f"Font restoration completed: {status_msg}")
    
    # Refresh the display
    refresh_fonts(None)

def find_matching_restore_tag(node_name, current_font, current_style, nodes_with_tags, text_block=None):
    """Find matching restoration tag for a node with precise matching"""
    best_match = None
    
    for tag_node_name, tag_data in nodes_with_tags.items():
        for tag in tag_data['tags']:
            # Check match by replacement_font and replacement_style
            if (tag.get('replacement_font') == current_font and 
                tag.get('replacement_style') == current_style):
                
                # For MultiText also check text_block match
                if text_block and tag_node_name == node_name:
                    # Если тег содержит информацию о text_block, проверяем соответствие
                    if f"{text_block}:" in tag.get('full_tag', ''):
                        return {
                            'tag': tag,
                            'source_node': tag_data['node'],
                            'source_node_name': tag_node_name
                        }
                    # Или сохраняем как возможный match
                    elif best_match is None:
                        best_match = {
                            'tag': tag,
                            'source_node': tag_data['node'],
                            'source_node_name': tag_node_name
                        }
                else:
                    # For regular nodes or when text_block is not set
                    return {
                        'tag': tag,
                        'source_node': tag_data['node'],
                        'source_node_name': tag_node_name
                    }
    
    return best_match

def restore_multitext_fonts(node, nodes_with_tags):
    """Restore fonts in MultiText node with proper style handling"""
    restored_count = 0
    skipped_count = 0
    
    try:
        # Use the same pattern-based approach
        common_patterns = [
            "Text1.Font", "Text1.Style",
            "Text2.Font", "Text2.Style", 
            "Text3.Font", "Text3.Style",
            "Text4.Font", "Text4.Style",
            "Text5.Font", "Text5.Style"
        ]
        
        # Process each text block
        for pattern in common_patterns:
            try:
                if '.Font' in pattern:
                    text_block = pattern.split('.Font')[0]
                    style_pattern = f"{text_block}.Style"
                    
                    current_font = node.GetInput(pattern)
                    current_style = node.GetInput(style_pattern) or "Regular"
                    
                    if current_font:
                        # Find matching restoration tag considering text_block
                        matching_tag = find_matching_restore_tag_for_multitext(
                            node.Name, 
                            text_block,
                            current_font, 
                            current_style, 
                            nodes_with_tags
                        )
                        
                        if matching_tag and 'tag' in matching_tag:
                            tag = matching_tag['tag']
                            original_font = tag.get("original_font")
                            original_style = tag.get("original_style")
                            
                            if original_font and original_style:
                                if check_font_style_availability(original_font, original_style):
                                    try:
                                        node.SetInput(pattern, original_font)
                                        node.SetInput(style_pattern, original_style)
                                        
                                        # Remove used tag
                                        remove_specific_restore_tag(matching_tag['source_node'], tag)
                                        
                                        restored_count += 1
                                        print(f"✓ Restored MultiText {node.Name}.{text_block}: {original_font} {original_style}")
                                    except Exception as e:
                                        print(f"✗ Failed to restore MultiText {text_block}: {e}")
                                        skipped_count += 1
                                else:
                                    skipped_count += 1
                                    print(f"✗ Cannot restore MultiText {text_block}: {original_font} {original_style} not available")
                            else:
                                print(f"⚠ Warning: Incomplete tag data for MultiText {node.Name}.{text_block}")
                                
            except Exception as e:
                print(f"Error restoring MultiText pattern {pattern}: {e}")
        
    except Exception as e:
        print(f"Error in restore_multitext_fonts: {e}")
    
    return restored_count, skipped_count

def find_matching_restore_tag_for_multitext(node_name, text_block, current_font, current_style, nodes_with_tags):
    """Find matching restoration tag for MultiText node with precise text_block matching"""
    for tag_node_name, tag_data in nodes_with_tags.items():
        for tag in tag_data['tags']:
            tag_replacement_font = tag.get('replacement_font')
            tag_replacement_style = tag.get('replacement_style')
            tag_text_block = tag.get('text_block')
            
            # Check match by replacement_font and replacement_style
            if (tag_replacement_font == current_font and 
                tag_replacement_style == current_style):
                
                # If tag contains text_block info, check exact match
                if tag_text_block:
                    if tag_text_block == text_block:
                        return {
                            'tag': tag,
                            'source_node': tag_data['node'],
                            'source_node_name': tag_node_name
                        }
                else:
                    # If text_block not specified in tag, use this tag
                    return {
                        'tag': tag,
                        'source_node': tag_data['node'],
                        'source_node_name': tag_node_name
                    }
    
    return None

def parse_all_restore_tags_from_comments(node):
    """Parse all restoration tags from node comments with enhanced data extraction"""
    tags = []
    try:
        comments = node.GetInput("Comments") or ""
        
        # Split comments into lines for more accurate parsing
        lines = comments.split('\n')
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            
            # Look for TextBlock line
            if line.startswith('TextBlock:'):
                text_block = line.replace('TextBlock:', '').strip()
                # Look for next [PostFlows_FONT_RESTORE] tag
                i += 1
                while i < len(lines) and '[PostFlows_FONT_RESTORE]' not in lines[i]:
                    i += 1
                
                if i < len(lines) and '[PostFlows_FONT_RESTORE]' in lines[i]:
                    # Collect full tag
                    tag_content = []
                    j = i
                    while j < len(lines) and '[/PostFlows_FONT_RESTORE]' not in lines[j]:
                        tag_content.append(lines[j])
                        j += 1
                    if j < len(lines):
                        tag_content.append(lines[j])
                        full_tag_content = '\n'.join(tag_content)
                        
                        # Parse tag content
                        original_font = extract_tag_value(full_tag_content, "original_font")
                        original_style = extract_tag_value(full_tag_content, "original_style")
                        replaced_with = extract_tag_value(full_tag_content, "replaced_with")
                        
                        # Parse replaced_with format: "font|style"
                        replacement_font = None
                        replacement_style = None
                        if replaced_with and '|' in replaced_with:
                            replacement_font, replacement_style = replaced_with.split('|', 1)
                        
                        if original_font and original_style:
                            tag_data = {
                                "original_font": original_font,
                                "original_style": original_style,
                                "replacement_font": replacement_font,
                                "replacement_style": replacement_style,
                                "text_block": text_block,
                                "full_tag": '\n'.join([f"TextBlock: {text_block}"] + tag_content)
                            }
                            tags.append(tag_data)
                        i = j + 1
                        continue
            elif '[PostFlows_FONT_RESTORE]' in line:
                # Tag without TextBlock (legacy format)
                tag_content = []
                j = i
                while j < len(lines) and '[/PostFlows_FONT_RESTORE]' not in lines[j]:
                    tag_content.append(lines[j])
                    j += 1
                if j < len(lines):
                    tag_content.append(lines[j])
                    full_tag_content = '\n'.join(tag_content)
                    
                    original_font = extract_tag_value(full_tag_content, "original_font")
                    original_style = extract_tag_value(full_tag_content, "original_style")
                    replaced_with = extract_tag_value(full_tag_content, "replaced_with")
                    
                    replacement_font = None
                    replacement_style = None
                    if replaced_with and '|' in replaced_with:
                        replacement_font, replacement_style = replaced_with.split('|', 1)
                    
                    if original_font and original_style:
                        tag_data = {
                            "original_font": original_font,
                            "original_style": original_style,
                            "replacement_font": replacement_font,
                            "replacement_style": replacement_style,
                            "full_tag": full_tag_content
                        }
                        tags.append(tag_data)
                    i = j + 1
                    continue
            
            i += 1
            
    except Exception as e:
        print(f"Error parsing restore tags: {e}")
    
    return tags

def remove_specific_restore_tag(node, tag_data):
    """Remove a specific restoration tag from node comments with better tracking"""
    try:
        comments = node.GetInput("Comments") or ""
        
        # Remove specific tag
        full_tag = tag_data["full_tag"]
        if full_tag in comments:
            new_comments = comments.replace(full_tag, "")
            
            # Очищаем лишние переносы строк
            new_comments = re.sub(r'\n\s*\n\s*\n', '\n\n', new_comments).strip()
            
            node.SetInput("Comments", new_comments)
            print(f"Removed specific restore tag from {node.Name}")
        else:
            print(f"Warning: Restore tag not found in {node.Name}, might have been already removed")
    except Exception as e:
        print(f"Error removing specific restore tag from {node.Name}: {e}")

def replace_missed_fonts(ev):
    """
    Replace fonts with user-selected replacement font and style.
    """
    if not timeline:
        print("No timeline is currently open!")
        return
    
    # Get user-selected replacement font and style
    replacement_font, replacement_style = get_selected_replacement_font()
    
    installed_fonts = get_installed_fonts()
    
    if replacement_font not in installed_fonts:
        print(f"Error: Selected font '{replacement_font}' is not installed. Please select a different font.")
        return
    
    if replacement_style not in installed_fonts[replacement_font]:
        print(f"Warning: Selected style '{replacement_style}' not available for {replacement_font}")
        available_styles = list(installed_fonts[replacement_font])
        replacement_style = available_styles[0] if available_styles else "Regular"
        print(f"Using fallback style: {replacement_style}")
        
        # Update UI with fallback style
        try:
            style_index = available_styles.index(replacement_style)
            itm["StyleCombo"].CurrentIndex = style_index
        except:
            pass
    
    # Create restoration log
    restoration_log = create_restoration_log()
    restore_id = restoration_log["restore_session_id"]
    
    replacements_made = 0
    
    for j in range(1, timeline.GetTrackCount("video") + 1):
        for tl_item in timeline.GetItemListInTrack("video", j):
            for k in range(1, tl_item.GetFusionCompCount() + 1):
                comp = tl_item.GetFusionCompByIndex(k)
                if comp:
                    comp.Lock()
                    
                    parent_replacements = {}
                    
                    for node in comp.GetToolList().values():
                        if node.Font:
                            original_font = node.Font[1]
                            original_style = node.Style[1] if node.Style else "Regular"
                            
                            needs_replacement = False
                            reason = ""
                            
                            # Check if replacement is needed
                            if original_font not in installed_fonts:
                                needs_replacement = True
                                reason = "font missing"
                            elif original_style not in installed_fonts[original_font]:
                                needs_replacement = True
                                reason = "style missing"
                            
                            if needs_replacement:
                                # Create restoration tag
                                restore_tag = create_restore_tag(
                                    original_font, original_style,
                                    replacement_font, replacement_style,
                                    restore_id
                                )
                                
                                # Add to restoration log
                                add_replacement_to_log(
                                    restoration_log, node,
                                    original_font, original_style,
                                    replacement_font, replacement_style
                                )
                                
                                # Improved parent detection
                                parent_node = find_parent_group_or_macro(comp, node)
                                use_node_directly = should_use_node_comments(comp, node)
                                
                                if parent_node and not use_node_directly:
                                    # Use parent logic for real macros/groups
                                    parent_name = parent_node.Name
                                    if parent_name not in parent_replacements:
                                        parent_replacements[parent_name] = {
                                            'node_object': parent_node,
                                            'replacements': []
                                        }
                                    parent_replacements[parent_name]['replacements'].append({
                                        'node': node,
                                        'original_font': original_font,
                                        'original_style': original_style,
                                        'reason': reason,
                                        'restore_tag': restore_tag
                                    })
                                else:
                                    # Use node directly for standalone Text+ or Text+ with only effects
                                    try:
                                        current_comments = node.GetInput("Comments") or ""
                                        if current_comments:
                                            new_comments = f"{current_comments}\n{restore_tag}"
                                        else:
                                            new_comments = restore_tag
                                        node.SetInput("Comments", new_comments)
                                    except Exception as e:
                                        print(f"Failed to add restore tag to node {node.Name}: {e}")
                                
                                try:
                                    node.SetInput("Font", replacement_font)
                                    node.SetInput("Style", replacement_style)
                                    replacements_made += 1
                                except Exception as e:
                                    print(f"Failed to set font {replacement_font} or style {replacement_style}: {e}")
                                            # Handle MultiText nodes
                        # Handle MultiText nodes
                        elif hasattr(node, 'GetAttrs'):
                            try:
                                attrs = node.GetAttrs()
                                if attrs and attrs.get('TOOLS_RegID', '') == 'MultiText':
                                    multitext_replacements = replace_multitext_fonts(node, replacement_font, replacement_style, 
                                                                                installed_fonts, restore_id, restoration_log)
                                    replacements_made += multitext_replacements
                            except Exception as e:
                                print(f"Error processing MultiText node in replacement: {e}")
                    
                    # Handle parent node comments with restore tags
                    for parent_name, parent_data in parent_replacements.items():
                        parent_node = parent_data['node_object']
                        replacements = parent_data['replacements']
                        
                        comment_lines = []
                        restore_tags = []
                        
                        for repl in replacements:
                            comment_lines.append(f"Original Font: {repl['original_font']}, Style: {repl['original_style']} ({repl['reason']})")
                            restore_tags.append(repl['restore_tag'])
                        
                        combined_comment = "\n".join(comment_lines)
                        combined_restore_tags = "\n".join(restore_tags)
                        
                        try:
                            existing_comment = ""
                            try:
                                existing_comment = parent_node.GetInput("Comments") or ""
                            except:
                                existing_comment = ""
                            
                            if existing_comment:
                                final_comment = f"{existing_comment}\n{combined_comment}\n{combined_restore_tags}"
                            else:
                                final_comment = f"{combined_comment}\n{combined_restore_tags}"
                            
                            parent_node.SetInput("Comments", final_comment)
                        except Exception as e:
                            print(f"Failed to set comment on parent node {parent_node.Name}: {e}")
                            # Fallback to individual nodes
                            for repl in replacements:
                                try:
                                    node_comment = f"Original Font: {repl['original_font']}, Style: {repl['original_style']} ({repl['reason']})"
                                    current_comments = repl['node'].GetInput("Comments") or ""
                                    if current_comments:
                                        final_node_comment = f"{current_comments}\n{node_comment}\n{repl['restore_tag']}"
                                    else:
                                        final_node_comment = f"{node_comment}\n{repl['restore_tag']}"
                                    repl['node'].SetInput("Comments", final_node_comment)
                                    print(f"Fallback: Added comment and restore tag to node {repl['node'].Name}")
                                except Exception as node_e:
                                    print(f"Failed to set fallback comment: {node_e}")
                    
                    comp.Unlock()
    
    # Save restoration log if any replacements were made
    if replacements_made > 0:
        log_path = save_restoration_log(restoration_log)
        status_msg = f"Replaced {replacements_made} fonts. Log saved to Desktop."
        if log_path:
            print(f"Restoration log saved to: {log_path}")
    else:
        status_msg = "No fonts needed replacement."
    
    print(f"Font replacement completed. Total replacements made: {replacements_made}")
    itm["StatusLabel"].Text = status_msg
    
    refresh_fonts(None)

def replace_multitext_fonts(node, replacement_font, replacement_style, 
                           installed_fonts, restore_id, restoration_log):
    """Replace fonts in MultiText node using the same pattern-based approach"""
    try:
        replacements_made = 0
        restore_tags = []
        
        # Use the same pattern-based approach as in extraction
        common_patterns = [
            "Text1.Font", "Text1.Style",
            "Text2.Font", "Text2.Style", 
            "Text3.Font", "Text3.Style",
            "Text4.Font", "Text4.Style",
            "Text5.Font", "Text5.Style"
        ]
        
        # Process each text block
        for pattern in common_patterns:
            try:
                if '.Font' in pattern:
                    text_block = pattern.split('.Font')[0]
                    
                    # Get original font and style
                    original_font = node.GetInput(pattern)
                    style_pattern = f"{text_block}.Style"
                    original_style = node.GetInput(style_pattern) or "Regular"
                    
                    if original_font:
                        needs_replacement = False
                        reason = ""
                        
                        # Check if replacement is needed
                        if original_font not in installed_fonts:
                            needs_replacement = True
                            reason = "font missing"
                        elif original_style not in installed_fonts.get(original_font, set()):
                            needs_replacement = True
                            reason = "style missing"
                        
                        if needs_replacement:
                            # Create restore tag for this text block with text_block info
                            restore_tag = create_restore_tag(
                                original_font, original_style,
                                replacement_font, replacement_style,
                                restore_id
                            )
                            # Always add text block identifier to the tag
                            restore_tag = f"TextBlock: {text_block}\n{restore_tag}"
                            restore_tags.append(restore_tag)
                            
                            # Replace font but preserve the STYLE if possible
                            try:
                                node.SetInput(pattern, replacement_font)
                                # Only replace style if the replacement font has this style available
                                if replacement_style in installed_fonts.get(replacement_font, set()):
                                    node.SetInput(style_pattern, replacement_style)
                                else:
                                    # Keep original style name or use Regular as fallback
                                    fallback_style = "Regular"
                                    if installed_fonts.get(replacement_font):
                                        fallback_style = list(installed_fonts[replacement_font])[0] if installed_fonts[replacement_font] else "Regular"
                                    node.SetInput(style_pattern, fallback_style)
                                    print(f"Warning: Style {replacement_style} not available for {replacement_font}, using {fallback_style}")
                                
                                replacements_made += 1
                                print(f"MultiText {node.Name}.{text_block}: {original_font}|{original_style} → {replacement_font}|{replacement_style}")
                                
                                # Add to log
                                add_replacement_to_log(restoration_log, node, original_font, original_style, 
                                                     replacement_font, replacement_style)
                                
                            except Exception as e:
                                print(f"Error replacing font in {text_block}: {e}")
                                
            except Exception as e:
                print(f"Error processing pattern {pattern} in replacement: {e}")
        
        # Add restore tags to node comments
        if restore_tags:
            combined_tags = "\n".join(restore_tags)
            try:
                current_comments = node.GetInput("Comments") or ""
                if current_comments:
                    new_comments = f"{current_comments}\n{combined_tags}"
                else:
                    new_comments = combined_tags
                node.SetInput("Comments", new_comments)
            except Exception as e:
                print(f"Error adding restore tags to MultiText node: {e}")
        
        return replacements_made
        
    except Exception as e:
        print(f"Error in replace_multitext_fonts: {e}")
        import traceback
        traceback.print_exc()
        return 0

def copy_missed_to_clipboard(ev):
    """
    Copy the list of missed fonts and their styles to the system clipboard.
    """
    used_fonts_data = get_used_fonts()
    missed_items = []
    
    for font_name, font_data in used_fonts_data.items():
        if font_data['font_missing']:
            # Entire font is missing
            styles_list = ', '.join(font_data['used_styles'])
            missed_items.append(f"{font_name} ({styles_list}) - FONT MISSING")
        elif font_data['missing_styles']:
            # Only some styles are missing
            missing_styles_list = ', '.join(font_data['missing_styles'])
            missed_items.append(f"{font_name} ({missing_styles_list}) - STYLES MISSING")
    
    copy_text = "\n".join(missed_items)
    
    if copy_text:
        try:
            if pyperclip:
                pyperclip.copy(copy_text)
            else:
                fusion.SetClipboard(copy_text)
            print(f"Missed fonts copied to clipboard:\n{copy_text}")
        except Exception as e:
            print(f"Error copying to clipboard: {e}")
    else:
        print("No missing fonts or styles to copy!")

def refresh_fonts(ev):
    """
    Refresh the font list table with updated data from the current timeline.
    """
    global used_fonts, timeline
    
    # Update reference to current timeline
    timeline = project.GetCurrentTimeline()
    if not timeline:
        print("No timeline is currently open!")
        return
    
    used_fonts = get_used_fonts()
    
    itm["FontList"].Clear()
    
    for font_name, font_data in used_fonts.items():
        for style in font_data['used_styles']:
            row = itm["FontList"].NewItem()
            row.SetText(0, font_name)
            row.SetText(1, style)
            
            # Определяем статус
            if font_data['font_missing']:
                status = "Font missing!"
            elif style in font_data['missing_styles']:
                status = "Style missing!"
            else:
                status = "Available"
            
            row.SetText(2, status)
            itm["FontList"].AddTopLevelItem(row)

# UI Styles (adapted from Offline Clips v03.py)
PRIMARY_COLOR = "#c0c0c0"
BORDER_COLOR = "#3a6ea5"
TEXT_COLOR = "#ebebeb"

PRIMARY_ACTION_BUTTON_STYLE = f"""
    QPushButton {{
        border: 1px solid #2C6E49;
        max-height: 30px;
        border-radius: 14px;
        background-color: #2C6E49;
        color: #FFFFFF;
        min-height: 20px;
        font-size: 12px;
        font-weight: bold;
        padding: 5px 15px;
    }}
    QPushButton:hover {{
        border: 1px solid {PRIMARY_COLOR};
        background-color: #4C956C;
    }}
    QPushButton:pressed {{
        border: 2px solid {PRIMARY_COLOR};
        background-color: #76C893;
    }}
"""

SECONDARY_ACTION_BUTTON_STYLE = f"""
    QPushButton {{
        border: 1px solid #2A4A7A;
        max-height: 30px;
        border-radius: 14px;
        background-color: #2A4A7A;
        color: #D0D0D0;
        min-height: 20px;
        font-size: 12px;
        font-weight: bold;
        padding: 5px 15px;
    }}
    QPushButton:hover {{
        border: 1px solid #3A5A9A;
        background-color: #3A5A9A;
        color: #FFFFFF;
    }}
    QPushButton:pressed {{
        border: 2px solid #4A6ABA;
        background-color: #4A6ABA;
        color: #FFFFFF;
    }}
    QPushButton:disabled {{
        border: 1px solid #333333;
        background-color: #333333;
        color: #666666;
    }}
"""

RESTORE_BUTTON_STYLE = f"""
    QPushButton {{
        border: 1px solid #A0522D;
        max-height: 30px;
        border-radius: 14px;
        background-color: #A0522D;
        color: #FFFFFF;
        min-height: 20px;
        font-size: 12px;
        font-weight: bold;
        padding: 5px 15px;
    }}
    QPushButton:hover {{
        border: 1px solid #CD853F;
        background-color: #CD853F;
    }}
    QPushButton:pressed {{
        border: 2px solid #DEB887;
        background-color: #DEB887;
    }}
    QPushButton:disabled {{
        border: 1px solid #333333;
        background-color: #333333;
        color: #666666;
    }}
"""

START_LOGO_CSS = """
    QLabel {
        color: #62b6cb;
        font-size: 22px;
        font-weight: bold;
        letter-spacing: 1px;
        font-family: 'Futura';
    }
"""

END_LOGO_CSS = """
    QLabel {
        color: rgb(255, 255, 255);
        font-size: 22px;
        font-weight: bold;
        letter-spacing: 1px;
        font-family: 'Futura';
    }
"""

STATUS_LABEL_STYLE = """
    QLabel {
        color: #c0c0c0;
        font-size: 13px;
        font-weight: bold;
        max-height: 25px;
        padding: 5px 0;
    }
"""

COMBO_STYLE = f"""
    QComboBox {{
        border: 1px solid {BORDER_COLOR};
        border-radius: 5px;
        background-color: #2A2A2A;
        color: {TEXT_COLOR};
        font-size: 12px;
        padding: 5px;
        min-width: 120px;
        min-height: 20px;
    }}
    QComboBox:focus {{
        border: 1px solid #4C956C;
    }}
    QComboBox::drop-down {{
        subcontrol-origin: padding;
        subcontrol-position: top right;
        width: 15px;
        border-left-width: 1px;
        border-left-color: {BORDER_COLOR};
        border-left-style: solid;
    }}
    QComboBox QAbstractItemView {{
        background-color: #2A2A2A;
        color: {TEXT_COLOR};
        selection-background-color: #4C956C;
        border: 1px solid {BORDER_COLOR};
    }}
"""

LIST_STYLE = f"""
    QTreeWidget {{
        border: 1px solid {BORDER_COLOR};
        border-radius: 5px;
        background-color: #2A2A2A;
        color: {TEXT_COLOR};
        font-size: 12px;
    }}
    QTreeWidget::item {{
        padding: 5px;
        height: 20px;
    }}
    QTreeWidget::item:selected {{
        background-color: #4C956C;
        color: #FFFFFF;
    }}
    QHeaderView::section {{
        background-color: #3A3A3A;
        color: {TEXT_COLOR};
        padding: 5px;
        border: 1px solid {BORDER_COLOR};
    }}
"""

def main_ui():
    """Creates and returns the main UI layout for the script."""
    return ui.VGroup({"Spacing": 10}, [
        # Header
        ui.HGroup({"Spacing": 5, "Weight": 0}, [
            ui.Label({"Text": "Font", "StyleSheet": START_LOGO_CSS, "Weight": 0}),
            ui.Label({"Text": "Fallback", "StyleSheet": END_LOGO_CSS, "Weight": 0})
        ]),
        
        # Replacement font selection section
        ui.VGroup({"Spacing": 5, "Weight": 0}, [
            ui.Label({
                "Text": "Font Replacement Settings", 
                "StyleSheet": STATUS_LABEL_STYLE,
                "Weight": 0
            }),
            ui.HGroup({"Spacing": 10}, [
                ui.Label({
                    "Text": "Font:", 
                    "StyleSheet": STATUS_LABEL_STYLE,
                    "Weight": 0
                }),
                ui.ComboBox({
                    "ID": "FontCombo",
                    "StyleSheet": COMBO_STYLE,
                    "Weight": 0,
                    "MinimumSize": [150, 25],
                    "MaximumSize": [200, 25]
                }),
                ui.Label({
                    "Text": "Style:", 
                    "StyleSheet": STATUS_LABEL_STYLE,
                    "Weight": 0
                }),
                ui.ComboBox({
                    "ID": "StyleCombo", 
                    "StyleSheet": COMBO_STYLE,
                    "Weight": 0,
                    "MinimumSize": [120, 25],
                    "MaximumSize": [150, 25]
                }),
                ui.Label({"Text": "", "Weight": 1})  # Right spacer
            ])
        ]),
        
        ui.VGap(10),
        
        # Основная таблица шрифтов
        ui.VGroup({"Spacing": 5}, [
            ui.Label({
                "Text": "Fonts in Current Timeline", 
                "StyleSheet": STATUS_LABEL_STYLE,
                "Weight": 0
            }),
            ui.Tree({
                "ID": "FontList",
                "HeaderText": "Font Name|Style|Status",
                "ColumnCount": 3,
                "ColumnWidth": "200,120,120",
                "Weight": 4,  # Увеличиваем вес таблицы
                "AlternatingRowColors": True,
                "RootIsDecorated": False,
                "SortingEnabled": True,
                "StyleSheet": LIST_STYLE
            })
        ]),
        
        ui.VGap(5),
        
        # Action buttons
        ui.HGroup({"Spacing": 5, "Weight": 0}, [
            ui.Button({
                "ID": "RefreshButton", 
                "Text": "Refresh Timeline",
                "StyleSheet": PRIMARY_ACTION_BUTTON_STYLE
            }),
            ui.Button({
                "ID": "ReplaceButton", 
                "Text": "Replace Missing",
                "StyleSheet": SECONDARY_ACTION_BUTTON_STYLE
            }),
            ui.Button({
                "ID": "RestoreButton", 
                "Text": "Restore Fonts",
                "StyleSheet": RESTORE_BUTTON_STYLE,
                "Enabled": True  # Всегда активна
            }),
            ui.Button({
                "ID": "CopyMissedButton", 
                "Text": "Copy Missing",
                "StyleSheet": SECONDARY_ACTION_BUTTON_STYLE
            })
        ]),
        
        # Status line
        ui.Label({
            "ID": "StatusLabel",
            "Text": "Press 'Refresh Timeline' to check for missing fonts and styles",
            "StyleSheet": STATUS_LABEL_STYLE,
            "Weight": 0,
            "Alignment": {"AlignHCenter": True}
        })
    ])

# Create UI
ui = fusion.UIManager
disp = bmd.UIDispatcher(ui)

win = disp.AddWindow({
    "WindowTitle": "PostFlows Font Fallback",
    "ID": "FontViewerWin",
    "Geometry": [250, 250, 550, 600]  # Увеличиваем ширину для новой кнопки
}, main_ui())

# Initialize UI Items
itm = win.GetItems()

# Initialize font combo boxes
installed_fonts = get_installed_fonts()
font_names = sorted(installed_fonts.keys())

for font_name in font_names:
    itm["FontCombo"].AddItem(font_name)

# Set default font
if DEFAULT_REPLACEMENT_FONT in font_names:
    default_index = font_names.index(DEFAULT_REPLACEMENT_FONT)
    itm["FontCombo"].CurrentIndex = default_index
elif font_names:
    itm["FontCombo"].CurrentIndex = 0

# Update styles for selected font
update_style_combo(None)

# Initialize font list from timeline
used_fonts = get_used_fonts()
for font_name, font_data in used_fonts.items():
    for style in font_data['used_styles']:
        row = itm["FontList"].NewItem()
        row.SetText(0, font_name)
        row.SetText(1, style)
        
        # Определяем статус
        if font_data['font_missing']:
            status = "Font missing!"
        elif style in font_data['missing_styles']:
            status = "Style missing!"
        else:
            status = "Available"
        
        row.SetText(2, status)
        itm["FontList"].AddTopLevelItem(row)

# Event bindings
win.On["FontCombo"].CurrentIndexChanged = update_style_combo
win.On["RefreshButton"].Clicked = refresh_fonts
win.On["ReplaceButton"].Clicked = replace_missed_fonts
win.On["RestoreButton"].Clicked = restore_original_fonts
win.On["CopyMissedButton"].Clicked = copy_missed_to_clipboard

def on_close(ev):
    """
    Close the window and exit the event loop.
    """
    disp.ExitLoop()

win.On.FontViewerWin.Close = on_close

win.Show()
disp.RunLoop()
