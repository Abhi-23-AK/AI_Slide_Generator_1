import io
import re
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.oxml.ns import qn

def analyze_template_v2(template_bytes: bytes, max_slides: int = 30) -> dict:
    """Analyze a PPTX template and return a V2 layout schema of its structure.
    
    This is the V2 pipeline analyzer. It extracts basic shape properties
    without any card clustering, layout inference, or decorative filtering.
    It maps shapes strictly to: title, subtitle, body1-4, image1-3, icon1-2, footer1.
    Small boxes and excess text shapes become other_<shape_id> and are never filled.
    """
    print(f"[TEMPLATE_V2] Starting V2 analysis of template ({len(template_bytes)} bytes)")
    prs = Presentation(io.BytesIO(template_bytes))

    slide_width = prs.slide_width
    slide_height = prs.slide_height

    # Extract theme info
    theme_info = _extract_theme_info(prs)

    slides_info = []
    for idx, slide in enumerate(prs.slides):
        if idx >= max_slides:
            break
            
        print(f"[TEMPLATE_V2] Analyzing Slide {idx + 1}/{len(prs.slides)}")
        slide_data = _analyze_slide_v2(slide, idx, slide_width, slide_height)
        slides_info.append(slide_data)

    schema = {
        "slide_count": len(slides_info),
        "theme": theme_info,
        "slides": slides_info,
    }
    
    print("[TEMPLATE_V2] V2 Analysis completed successfully.")
    return schema


def _extract_theme_info(prs):
    theme_info = {
        "heading_font": "Calibri",
        "body_font": "Calibri",
        "colors": {
            "dk1": "#000000",
            "lt1": "#FFFFFF",
            "dk2": "#1F4E79",
            "lt2": "#D9E1F2",
            "accent1": "#4472C4",
            "accent2": "#ED7D31",
            "accent3": "#A5A5A5",
            "accent4": "#FFC000",
            "accent5": "#5B9BD5",
            "accent6": "#70AD47",
        }
    }
    try:
        if not prs.slide_masters:
            return theme_info
        master = prs.slide_masters[0]
        font_scheme = master.element.findall('.//' + qn('a:fontScheme'))
        if font_scheme:
            major = font_scheme[0].find(qn('a:majorFont'))
            if major is not None:
                latin = major.find(qn('a:latin'))
                if latin is not None and latin.get('typeface'):
                    theme_info['heading_font'] = latin.get('typeface')
            minor = font_scheme[0].find(qn('a:minorFont'))
            if minor is not None:
                latin = minor.find(qn('a:latin'))
                if latin is not None and latin.get('typeface'):
                    theme_info['body_font'] = latin.get('typeface')

        clr_scheme = master.element.findall('.//' + qn('a:clrScheme'))
        if clr_scheme:
            color_names = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']
            for name in color_names:
                color_el = clr_scheme[0].find(qn(f'a:{name}'))
                if color_el is not None:
                    srgb = color_el.find(qn('a:srgbClr'))
                    if srgb is not None and srgb.get('val'):
                        theme_info['colors'][name] = f'#{srgb.get("val")}'
                    else:
                        sys_clr = color_el.find(qn('a:sysClr'))
                        if sys_clr is not None and sys_clr.get('lastClr'):
                            theme_info['colors'][name] = f'#{sys_clr.get("lastClr")}'
    except Exception as e:
        print(f"[TEMPLATE_V2] Warning: Theme parsing failed: {e}")
    return theme_info


def _get_all_shapes_recursive(shapes, slide_width, slide_height, group_offset_left=0, group_offset_top=0):
    """Recursively flatten GROUP containers so child shapes are analyzed.

    Returns a list of (shape, abs_left_offset, abs_top_offset) tuples.
    For non-group shapes the offsets are simply passed through;
    for group children the group's position is accumulated as an offset.
    """
    result = []
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            new_offset_left = group_offset_left + shape.left
            new_offset_top = group_offset_top + shape.top
            result.extend(
                _get_all_shapes_recursive(
                    shape.shapes, slide_width, slide_height,
                    new_offset_left, new_offset_top,
                )
            )
        else:
            result.append((shape, group_offset_left, group_offset_top))
    return result


def _analyze_slide_v2(slide, idx, slide_width, slide_height):
    bg_info = _extract_background_info(slide)

    shapes_list = []
    all_shapes = _get_all_shapes_recursive(slide.shapes, slide_width, slide_height)
    for shape_idx, (shape, offset_left, offset_top) in enumerate(all_shapes):
        shape_data = _analyze_shape_v2(shape, shape_idx, slide_width, slide_height, offset_left, offset_top)
        shapes_list.append(shape_data)
        
        # If it's a table, also extract each cell as an individual text zone
        if shape_data.get("content_type") == "table":
            try:
                table = shape.table
                for r_idx, row in enumerate(table.rows):
                    for c_idx, cell in enumerate(row.cells):
                        cell_role = f"table_{shape_data['shape_id']}_r{r_idx}c{c_idx}"
                        cell_orig_text = cell.text_frame.text.strip()
                        
                        font_name = "Calibri"
                        font_size = 12.0
                        font_color = None
                        bold = False
                        italic = False
                        
                        if cell.text_frame.paragraphs and cell.text_frame.paragraphs[0].runs:
                            run = cell.text_frame.paragraphs[0].runs[0]
                            font_name = run.font.name or font_name
                            if run.font.size:
                                font_size = float(run.font.size.pt)
                            if run.font.color and run.font.color.type is not None:
                                try:
                                    font_color = f"#{run.font.color.rgb}"
                                except Exception:
                                    pass
                            bold = bool(run.font.bold)
                            italic = bool(run.font.italic)
                            
                        cell_w = shape.width.inches / len(table.columns)
                        cell_h = shape.height.inches / len(table.rows)
                        usable_w = max(0.2, cell_w - 0.1)
                        usable_h = max(0.1, cell_h - 0.05)
                        chars_per_line = max(1, int(usable_w / (font_size * 0.006)))
                        max_lines = max(1, int(usable_h / (font_size * 0.0166)))
                        cell_max_chars = chars_per_line * max_lines
                        cell_max_chars = max(5, min(300, cell_max_chars))
                        
                        cell_data = {
                            "shape_id": shape.shape_id,
                            "name": f"{shape.name}_r{r_idx}c{c_idx}",
                            "shape_type": "table_cell",
                            "z_order": shape_idx,
                            "left_pct": shape_data["left_pct"],
                            "top_pct": shape_data["top_pct"],
                            "width_pct": shape_data["width_pct"],
                            "height_pct": shape_data["height_pct"],
                            "left_emu": shape_data["left_emu"],
                            "top_emu": shape_data["top_emu"],
                            "width_emu": shape_data["width_emu"],
                            "height_emu": shape_data["height_emu"],
                            "has_text_frame": True,
                            "font_name": font_name,
                            "font_size": font_size,
                            "font_color": font_color,
                            "bold": bold,
                            "italic": italic,
                            "underline": False,
                            "alignment": "left",
                            "capacity": max(5, int(cell_max_chars / 5.5)),
                            "max_chars": cell_max_chars,
                            "original_text": cell_orig_text,
                            "role": cell_role,
                            "content_type": "text",
                            "row_idx": r_idx,
                            "col_idx": c_idx,
                        }
                        shapes_list.append(cell_data)
            except Exception as e:
                print(f"[TEMPLATE_V2] Failed to parse table cells: {e}")
                
        # If it's a chart, parse categories & series metadata
        elif shape_data.get("content_type") == "chart":
            try:
                chart = shape.chart
                categories = []
                series_names = []
                try:
                    categories = [str(cat) for cat in chart.plots[0].categories]
                except Exception:
                    try:
                        categories = [str(cat) for cat in chart.data.categories]
                    except Exception:
                        pass
                try:
                    series_names = [str(s.name) for s in chart.series]
                except Exception:
                    try:
                        series_names = [str(s.name) for s in chart.data.series]
                    except Exception:
                        pass
                    
                shape_data["categories"] = categories
                shape_data["series"] = series_names
                shape_data["original_text"] = f"Chart categories: {categories}, series: {series_names}"
                shape_data["role"] = f"chart_{shape_data['shape_id']}"
            except Exception as e:
                print(f"[TEMPLATE_V2] Failed to parse chart data: {e}")

    # Assign V2 roles strictly following the instructions
    _assign_v2_roles(shapes_list)

    # Classify layout pattern
    inferred_layout = _classify_layout(shapes_list)
    column_count = _count_columns(shapes_list)

    return {
        "slide_index": idx,
        "inferred_layout": inferred_layout,
        "zones": shapes_list,
        "column_count": column_count,
        "has_image_placeholder": any(s["role"].startswith("image") for s in shapes_list),
        "background_color": bg_info.get("color"),
        "background_type": bg_info.get("type", "none"),
        "background_gradient": bg_info.get("gradient"),
    }


def _extract_background_info(slide):
    """Extract background info: solid color, gradient, or image."""
    info = {"type": "none", "color": None, "gradient": None}
    try:
        bg = slide.background
        if bg and bg.fill:
            fill_type = bg.fill.type
            if fill_type == 1:  # Solid
                info["type"] = "solid"
                info["color"] = f"#{bg.fill.fore_color.rgb}"
            elif fill_type == 2:  # Gradient
                info["type"] = "gradient"
                try:
                    stops = []
                    for stop in bg.fill.gradient_stops:
                        stops.append({
                            "color": f"#{stop.color.rgb}" if stop.color and stop.color.rgb else None,
                            "position": stop.position,
                        })
                    info["gradient"] = stops
                except Exception:
                    pass
            elif fill_type == 3:  # Pattern
                info["type"] = "pattern"
                try:
                    info["color"] = f"#{bg.fill.fore_color.rgb}"
                except Exception:
                    pass
    except Exception:
        pass
    return info


def _classify_layout(shapes_list):
    """Classify a slide's layout pattern based on shape positions and roles."""
    image_shapes = [s for s in shapes_list if s["role"].startswith("image")]
    icon_shapes = [s for s in shapes_list if s["role"].startswith("icon")]
    body_shapes = [s for s in shapes_list if s["role"].startswith("body")]

    # Hero: single large image covering > 40% of slide area
    for img in image_shapes:
        area = img["width_pct"] * img["height_pct"] / 100
        if area > 40:
            return "hero"

    # Icon row: 3+ icons at similar y-position
    if len(icon_shapes) >= 3:
        y_positions = [s["top_pct"] for s in icon_shapes]
        y_spread = max(y_positions) - min(y_positions)
        if y_spread < 15:
            return "icon_row"

    # Grid: 4 cards/sections (typically 4 body shapes arranged in 2 columns)
    if len(body_shapes) >= 4:
        left_positions = [s["left_pct"] for s in body_shapes]
        columns = _group_into_columns(left_positions)
        if len(columns) == 2:
            return "grid"
        elif len(columns) >= 3:
            return "three_column"

    if len(body_shapes) == 3:
        return "three_column"
    elif len(body_shapes) == 2:
        return "two_column"
    elif len(body_shapes) == 1:
        return "single_column"

    return "content"


def _group_into_columns(left_positions, threshold=10):
    """Group left positions into columns based on proximity."""
    if not left_positions:
        return []
    sorted_pos = sorted(set(left_positions))
    columns = [[sorted_pos[0]]]
    for pos in sorted_pos[1:]:
        if pos - columns[-1][-1] < threshold:
            columns[-1].append(pos)
        else:
            columns.append([pos])
    return columns


def _count_columns(shapes_list):
    """Count the number of content columns in a slide."""
    body_shapes = [s for s in shapes_list if s["role"].startswith("body")]
    if not body_shapes:
        return 1
    left_positions = [s["left_pct"] for s in body_shapes]
    columns = _group_into_columns(left_positions)
    return max(1, len(columns))


def _analyze_shape_v2(shape, z_order, slide_width, slide_height, offset_left=0, offset_top=0):
    abs_left = shape.left + offset_left
    abs_top = shape.top + offset_top
    left_pct = (abs_left / slide_width) * 100 if slide_width else 0
    top_pct = (abs_top / slide_height) * 100 if slide_height else 0
    width_pct = (shape.width / slide_width) * 100 if slide_width else 0
    height_pct = (shape.height / slide_height) * 100 if slide_height else 0

    shape_type_name = "unknown"
    if shape.shape_type is not None:
        try:
            shape_type_name = shape.shape_type.name
        except Exception:
            shape_type_name = str(shape.shape_type)

    placeholder_idx = None
    placeholder_type = None
    if shape.is_placeholder:
        try:
            placeholder_idx = shape.placeholder_format.idx
            ph_type = shape.placeholder_format.type
            placeholder_type = ph_type.name if hasattr(ph_type, "name") else str(ph_type)
        except Exception:
            pass

    is_picture = False
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            is_picture = True
    except Exception:
        pass
    if not is_picture and shape.is_placeholder:
        try:
            if shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
                is_picture = True
        except Exception:
            pass

    # ── Extract size in inches ──
    width_in = shape.width.inches
    height_in = shape.height.inches

    # ── Extract colors ──
    fill_color = None
    line_color = None
    try:
        if hasattr(shape, "fill") and shape.fill:
            if shape.fill.type == 1:  # Solid
                fill_color = f"#{shape.fill.fore_color.rgb}"
            elif shape.fill.type == 2:  # Gradient
                fill_color = "gradient"
    except Exception:
        pass
    try:
        if hasattr(shape, "line") and shape.line:
            if shape.line.color and shape.line.color.type is not None:
                line_color = f"#{shape.line.color.rgb}"
    except Exception:
        pass

    # ── Classify content type ──
    content_type = "decorative"
    
    is_table = False
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
            is_table = True
    except Exception:
        pass

    is_chart = False
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.CHART:
            is_chart = True
    except Exception:
        pass

    if is_table:
        content_type = "table"
    elif is_chart:
        content_type = "chart"
    elif is_picture:
        # Check size and aspect ratio to distinguish image vs icon
        is_small = (width_in < 1.5 and height_in < 1.5)
        aspect_ratio = (width_in / height_in) if height_in > 0 else 1.0
        is_square_like = (0.7 <= aspect_ratio <= 1.4)
        if is_small and is_square_like:
            content_type = "icon"
        else:
            content_type = "image"
    elif shape.has_text_frame and shape.text_frame.text.strip():
        content_type = "text"
    else:
        # Check if it could be an icon (small, no text, vector shape or circle/symbol)
        name_lower = (shape.name or "").lower()
        is_small_shape = (width_in < 1.2 and height_in < 1.2)
        has_icon_name = any(k in name_lower for k in ("icon", "logo", "vector", "symbol", "svg"))
        if is_small_shape and (has_icon_name or shape_type_name in ("OVAL", "ELLIPSE") or shape.shape_type in (MSO_SHAPE_TYPE.FREEFORM, MSO_SHAPE_TYPE.AUTO_SHAPE)):
            content_type = "icon"

    # ── Extract per-paragraph/per-run font properties ──
    font_name = None
    font_size = None
    font_color = None
    bold = False
    italic = False
    underline = False
    alignment = "left"
    paragraphs_info = []

    if shape.has_text_frame:
        tf = shape.text_frame
        for p_idx, p in enumerate(tf.paragraphs):
            p_alignment = "left"
            if p.alignment is not None:
                try:
                    p_alignment = p.alignment.name.lower()
                except Exception:
                    pass

            runs_info = []
            for run in p.runs:
                run_data = {
                    "font_name": run.font.name,
                    "font_size": None,
                    "font_color": None,
                    "bold": bool(run.font.bold),
                    "italic": bool(run.font.italic),
                    "underline": bool(run.font.underline),
                }
                if run.font.size:
                    try:
                        run_data["font_size"] = float(run.font.size.pt)
                    except Exception:
                        pass
                if run.font.color and run.font.color.type is not None:
                    try:
                        run_data["font_color"] = f"#{run.font.color.rgb}"
                    except Exception:
                        pass
                runs_info.append(run_data)

            paragraphs_info.append({
                "alignment": p_alignment,
                "runs": runs_info,
            })

            # Populate top-level font fields from the FIRST paragraph's FIRST run
            if p_idx == 0:
                alignment = p_alignment
                for run in p.runs:
                    if run.font.name:
                        font_name = run.font.name
                    if run.font.size:
                        try:
                            font_size = float(run.font.size.pt)
                        except Exception:
                            pass
                    if run.font.color and run.font.color.type is not None:
                        try:
                            font_color = f"#{run.font.color.rgb}"
                        except Exception:
                            pass
                    bold = bool(run.font.bold)
                    italic = bool(run.font.italic)
                    underline = bool(run.font.underline)
                    break

    # ── Compute capacity (word count estimate) ──
    capacity = 40
    if shape.has_text_frame:
        try:
            area = width_in * height_in
            f_sz = font_size if font_size else 14.0
            word_area = (f_sz ** 2 * 3.6) / 5184
            capacity = int(area / word_area) if word_area > 0 else 0
            capacity = max(5, min(500, capacity))
        except Exception:
            capacity = 40

    # ── Compute max_chars (character-level limit for overflow prevention) ──
    max_chars = 200  # safe default
    if shape.has_text_frame:
        try:
            f_sz = font_size if font_size else 14.0
            usable_w = max(0.5, width_in - 0.2)
            usable_h = max(0.2, height_in - 0.1)
            # Average character width ≈ font_size * 0.006 inches
            chars_per_line = max(1, int(usable_w / (f_sz * 0.006)))
            # Line height ≈ font_size * 0.0166 inches (includes 1.2x leading)
            max_lines = max(1, int(usable_h / (f_sz * 0.0166)))
            max_chars = chars_per_line * max_lines
            max_chars = max(5, min(2000, max_chars))
        except Exception:
            max_chars = 200

    image_filename = ""
    image_content_type = ""
    if is_picture:
        try:
            if hasattr(shape, "image") and shape.image is not None:
                image_filename = getattr(shape.image, "filename", "") or ""
                image_content_type = getattr(shape.image, "content_type", "") or ""
        except Exception:
            pass

    return {
        "shape_id": shape.shape_id,
        "name": shape.name or "",
        "shape_type": shape_type_name,
        "placeholder_idx": placeholder_idx,
        "placeholder_type": placeholder_type,
        "z_order": z_order,
        "left_pct": round(left_pct, 2),
        "top_pct": round(top_pct, 2),
        "width_pct": round(width_pct, 2),
        "height_pct": round(height_pct, 2),
        "left_emu": int(abs_left),
        "top_emu": int(abs_top),
        "width_emu": int(shape.width),
        "height_emu": int(shape.height),
        "width_in": round(width_in, 3),
        "height_in": round(height_in, 3),
        "has_text_frame": shape.has_text_frame,
        "is_picture": is_picture,
        "image_filename": image_filename,
        "image_content_type": image_content_type,
        "font_name": font_name,
        "font_size": font_size,
        "font_color": font_color,
        "bold": bold,
        "italic": italic,
        "underline": underline,
        "fill_color": fill_color,
        "line_color": line_color,
        "content_type": content_type,
        "alignment": alignment,
        "paragraphs_info": paragraphs_info,
        "capacity": capacity,
        "max_chars": max_chars,
        "original_text": shape.text_frame.text.strip() if shape.has_text_frame else "",
        "role": "other",
    }



def _is_decorative_artwork(s, shapes):
    # Genuine picture placeholders are NEVER decorative artwork
    if s.get("placeholder_type") == "PICTURE" or s.get("placeholder_idx") is not None:
        return False
        
    name_lower = (s.get("name") or "").lower()
    
    # Genuine icon placeholders or shapes explicitly named icon/logo/symbol are NOT decorative background
    has_explicit_icon_name = any(k in name_lower for k in ("icon", "logo", "symbol", "avatar"))
    if has_explicit_icon_name and not any(k in name_lower for k in ("bg", "background", "decor", "pattern")):
        return False

    decorative_keywords = (
        "leaf", "tree", "bg", "background", "decor", "art", "graphic", "logo", "watermark", 
        "illustration", "ornament", "flower", "plant", "abstract", "pattern",
        "border", "corner", "banner", "watermark", "decoration", "artwork", "ornamentation"
    )

    # Check shape name
    if any(kw in name_lower for kw in decorative_keywords):
        return True

    # Check image filename & content type if it is a picture
    image_filename_lower = s.get("image_filename", "").lower()
    image_content_type_lower = s.get("image_content_type", "").lower()
    
    if "svg" in image_content_type_lower or image_filename_lower.endswith(".svg"):
        return True
        
    if any(kw in image_filename_lower for kw in decorative_keywords):
        return True

    # Low z-order:
    # If the shape is in the bottom-most layer (z_order < 3) and it's not a placeholder, it is likely background artwork
    if s.get("z_order", 0) < 3 and s.get("placeholder_idx") is None:
        return True

    # Very small relative to slide:
    # If area is less than 1.5% of slide area, and it's not a placeholder, it's decorative/ornamental
    area_pct = s["width_pct"] * s["height_pct"]
    if area_pct < 150 and s.get("placeholder_idx") is None:
        return True

    # Positioned in a corner purely as ornamentation:
    # Check if the shape is located in one of the 4 corners:
    in_top_left = (s["left_pct"] < 25 and s["top_pct"] < 25)
    in_top_right = (s["left_pct"] + s["width_pct"] > 75 and s["top_pct"] < 25)
    in_bottom_left = (s["left_pct"] < 25 and s["top_pct"] + s["height_pct"] > 75)
    in_bottom_right = (s["left_pct"] + s["width_pct"] > 75 and s["top_pct"] + s["height_pct"] > 75)
    
    if in_top_left or in_top_right or in_bottom_left or in_bottom_right:
        # If it's in a corner, and not a placeholder, and area is less than 12% of the slide, treat as decorative
        if area_pct < 1200 and s.get("placeholder_idx") is None:
            return True

    # Check overlap with background/text elements (e.g. page numbers, footers, headers)
    for other in shapes:
        if other == s:
            continue
        if other.get("has_text_frame") and other.get("original_text"):
            # Bounding box overlap check
            s_left = s["left_pct"]
            s_right = s["left_pct"] + s["width_pct"]
            s_top = s["top_pct"]
            s_bottom = s["top_pct"] + s["height_pct"]
            
            o_left = other["left_pct"]
            o_right = other["left_pct"] + other["width_pct"]
            o_top = other["top_pct"]
            o_bottom = other["top_pct"] + other["height_pct"]
            
            # Simple rectangle intersection check
            if not (s_right <= o_left or s_left >= o_right or s_bottom <= o_top or s_top >= o_bottom):
                text_clean = other["original_text"].strip()
                # If they overlap, and the other is a page number, slide number, footer, or short header/footer text
                is_num = text_clean.isdigit() or bool(re.match(r'^\d+/\d+$', text_clean))
                if is_num or other["top_pct"] > 80 or other["top_pct"] < 15 or len(text_clean) < 15:
                    return True

    # When in doubt, do not touch the shape:
    # If it is a picture shape but has no placeholder format, and is not a substantial main image
    if s.get("placeholder_idx") is None and s.get("shape_type") == "PICTURE":
        # If it is smaller than 2.0 inches wide or 1.5 inches high, we err on the side of caution
        if s.get("width_in", 0) < 2.0 or s.get("height_in", 0) < 1.5:
            return True

    return False



def _assign_v2_roles(shapes):
    """Assign strict V2 roles to slide shapes.

    ALLOWED ROLES ONLY:
        title, subtitle, body1 to bodyN,
        image1 to imageM, icon1 to iconK, footer1,
        other_<shape_id> (never filled).
    """
    footers = []
    large_pictures = []
    icons = []
    text_boxes = []

    for s in shapes:
        # Skip if shape already has a custom role assigned (like table cells or charts)
        role = s.get("role", "other")
        if role.startswith("table_") or role.startswith("chart_"):
            continue

        name_lower = s["name"].lower()
        stype_upper = s["shape_type"].upper()
        text_clean = s["original_text"].strip()

        # ── Background shapes / pictures: do NOT classify as content/image candidates (preserve them!) ──
        is_background_shape = (
            s["left_pct"] <= 5 and 
            s["top_pct"] <= 5 and 
            s["width_pct"] >= 90 and 
            s["height_pct"] >= 90
        )
        if is_background_shape:
            continue

        # Check if text is page number, slide number, date, copyright, etc.
        is_page_num_or_date = False
        if s["has_text_frame"] and text_clean:
            is_slide_num = text_clean.isdigit() and int(text_clean) < 150 and s["top_pct"] > 85
            is_fraction_num = bool(re.match(r'^\d+/\d+$', text_clean)) and s["top_pct"] > 85
            
            is_page_num_or_date = (
                is_slide_num or
                is_fraction_num or 
                (bool(re.match(r'^\d{4}-\d{2}-\d{2}$', text_clean)) and s["top_pct"] > 85) or
                ("page" in text_clean.lower() and s["top_pct"] > 85) or
                "copyright" in text_clean.lower() or
                "confidential" in text_clean.lower()
            )

        # ── Footer: bottom-anchored small text or page numbers/dates ──
        if s["has_text_frame"] and (
            (s["top_pct"] > 88 and s["height_pct"] < 10 and s["width_pct"] < 40) or
            is_page_num_or_date
        ):
            footers.append(s)
            continue

        # ── Distinguish Large Images vs Small Icons ──
        is_large_image = s.get("content_type") == "image"
        is_icon = s.get("content_type") == "icon"

        if is_large_image:
            if not _is_decorative_artwork(s, shapes):
                large_pictures.append(s)
                continue
        elif is_icon:
            if not _is_decorative_artwork(s, shapes):
                icons.append(s)
                continue

        # ── Decorative sequential numbers (01, 02, 03, etc.) — never filled ──
        if s["has_text_frame"] and text_clean:
            is_decorative_number = bool(re.match(r'^\d{1,2}$', text_clean)) and len(text_clean) <= 2
            # Also check for numbers with leading zeros like '01', '02'
            if is_decorative_number:
                # Check context: if it's small and there are other shapes with similar numbers nearby, it's decorative
                if s["width_pct"] < 15 and s["height_pct"] < 15:
                    s["role"] = f"other_{s['shape_id']}"
                    continue

        # ── Text boxes (excluding chart/table/graph labels and empty text) ──
        if s["has_text_frame"] and text_clean and not any(k in name_lower for k in ("table", "chart", "graph")):
            text_boxes.append(s)
            continue

    # Qualified text boxes — no small box filter anymore!
    qualified_text = text_boxes

    # ── 1. Title: largest top textbox (top_pct < 35) ──
    title_shape = None
    top_textboxes = [s for s in qualified_text if s["top_pct"] < 35]
    if top_textboxes:
        # Sort by font size DESC, then by area DESC
        top_textboxes.sort(key=lambda s: (s.get("font_size") or 0, s["width_pct"] * s["height_pct"]), reverse=True)
        title_shape = top_textboxes[0]
        title_shape["role"] = "title"

    # ── 2. Subtitle: second largest textbox (top_pct < 45) ──
    subtitle_shape = None
    remaining_top = [s for s in qualified_text if s != title_shape and s["top_pct"] < 45]
    if remaining_top:
        remaining_top.sort(key=lambda s: (s.get("font_size") or 0, s["width_pct"] * s["height_pct"]), reverse=True)
        subtitle_shape = remaining_top[0]
        subtitle_shape["role"] = "subtitle"

    # ── 3. Body1-N: remaining textboxes sorted by reading order (top-to-bottom, left-to-right) ──
    bodies = [s for s in qualified_text if s != title_shape and s != subtitle_shape]
    # Reading order sort: group by top position rounded to nearest 5% to handle minor misalignments, then sort by left_pct
    bodies.sort(key=lambda s: (round(s["top_pct"] / 5) * 5, s["left_pct"]))
    for i, s in enumerate(bodies):
        s["role"] = f"body{i + 1}"

    # ── 4. Images: image1, image2, image3 ... (left-to-right) ──
    large_pictures.sort(key=lambda s: (round(s["top_pct"] / 10) * 10, s["left_pct"]))
    for i, s in enumerate(large_pictures):
        s["role"] = f"image{i + 1}"

    # ── 5. Icons: icon1, icon2 ... (left-to-right) ──
    icons.sort(key=lambda s: (round(s["top_pct"] / 10) * 10, s["left_pct"]))
    for i, s in enumerate(icons):
        s["role"] = f"icon{i + 1}"

    # ── 6. Footer: first only ──
    footers.sort(key=lambda s: s["left_pct"])
    for i, s in enumerate(footers):
        if i == 0:
            s["role"] = "footer1"
        else:
            s["role"] = f"other_{s['shape_id']}"

    # ── 7. Everything still marked "other" gets other_<id> ──
    for s in shapes:
        if s.get("role") == "other":
            s["role"] = f"other_{s['shape_id']}"
