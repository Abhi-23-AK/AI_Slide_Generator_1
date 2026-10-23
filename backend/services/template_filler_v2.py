"""
Template Filler V2 — Direct Shape Replacement
==============================================
Fills a PPTX template by matching shapes via shape_id and replacing
content in-place. Preserves all template formatting.

This is the V2 pipeline filler. It:
  - Matches shapes by shape_id (primary) → placeholder_idx → name → coords
  - Preserves existing paragraph structure, run formatting, fonts, colors
  - Replaces image blobs in-place (no shape resizing or repositioning)
  - Skips icons, decorative shapes, and footers entirely
  - Truncates gracefully with "..." instead of overflowing
  - Never uses fallback shape matching — missing content is better than wrong content
"""

import io
import os
import json
import copy
import uuid
import threading
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.oxml.ns import qn
from pptx.enum.shapes import MSO_SHAPE_TYPE

thread_local = threading.local()


def fill_template_v2(
    template_path: str,
    template_schema: dict,
    slide_contents: list,
    output_dir: str = "outputs",
    topic: str = "",
) -> str:
    """Fill a PPTX template with V2-generated content.

    Args:
        template_path: Path to the original .pptx template file.
        template_schema: V2 schema dict from analyze_template_v2().
        slide_contents: List of flat dicts, one per slide, with role-keyed content.
        output_dir: Directory for saving output.
        topic: Overall presentation topic for contextual image queries.

    Returns:
        Absolute path to the saved .pptx file.
    """
    print(f"[FILLER_V2] Starting template fill: {template_path}")
    print(f"[FILLER_V2] Slides to fill: {len(slide_contents)}")

    # Pre-download all images in parallel if visual engine is available
    try:
        preloaded_maps = _download_all_template_images(slide_contents, topic)
        for idx, content in enumerate(slide_contents):
            if idx < len(preloaded_maps) and preloaded_maps[idx]:
                content["_preloaded_image"] = preloaded_maps[idx]
    except Exception as e:
        print(f"[FILLER_V2] Parallel image pre-download failed: {e}")

    prs = Presentation(template_path)

    # Generate output filename
    os.makedirs(output_dir, exist_ok=True)
    out_name = f"template_{uuid.uuid4().hex[:10]}.pptx"
    out_path = os.path.join(output_dir, out_name)

    slide_width = prs.slide_width
    slide_height = prs.slide_height

    slides_schema = template_schema.get("slides", [])

    # ── Capture initial shape state for validation ──
    initial_state = _capture_shape_state(prs)

    # ── Fill each slide ──
    for slide_idx, slide in enumerate(prs.slides):
        if slide_idx >= len(slide_contents):
            print(f"[FILLER_V2] Slide {slide_idx + 1}: no content provided, skipping")
            continue

        content = slide_contents[slide_idx]
        if not content:
            print(f"[FILLER_V2] Slide {slide_idx + 1}: empty content, skipping")
            continue

        schema_slide = slides_schema[slide_idx] if slide_idx < len(slides_schema) else None
        if not schema_slide:
            print(f"[FILLER_V2] Slide {slide_idx + 1}: no schema found, skipping")
            continue

        _fill_slide_v2(slide, content, schema_slide, slide_idx, slide_width, slide_height)

    # ── Remove slides explicitly marked for deletion in reverse order ──
    for del_idx in range(len(prs.slides) - 1, -1, -1):
        if del_idx < len(slide_contents):
            content = slide_contents[del_idx]
            if content and (content.get("delete") is True or content.get("action") == "delete"):
                try:
                    _delete_slide(prs, del_idx)
                    print(f"[FILLER_V2] Deleted explicitly excluded slide {del_idx + 1}")
                except Exception as e:
                    print(f"[FILLER_V2] Warning: Could not delete slide {del_idx + 1}: {e}")

    # ── Post-fill validation ──
    _validate_shape_state(prs, initial_state)

    # ── Post-fill verification pass ──
    try:
        _verify_and_shorten_text_frames(prs, template_schema, slide_contents)
    except Exception as e:
        print(f"[FILLER_V2] Warning: Post-fill verification pass failed: {e}")

    # ── Apply complex script font fix ──
    fix_presentation_complex_scripts(prs)

    # ── Save ──
    prs.save(out_path)
    abs_path = os.path.abspath(out_path)
    print(f"[FILLER_V2] Saved: {abs_path}")
    return abs_path


# ═══════════════════════════════════════════════════════════════════════
# SLIDE FILLING
# ═══════════════════════════════════════════════════════════════════════

# ── STRICT WHITELIST: only these roles receive content ──
FILLABLE_TEXT_ROLES = {"title", "subtitle", "body1", "body2", "body3", "body4"}
FILLABLE_IMAGE_ROLES = {"image1", "image2", "image3"}


def _fill_slide_v2(slide, content, schema_slide, slide_idx, slide_width, slide_height):
    """Fill a single slide with content.

    Allows title, subtitle, body*, image*, icon*, table_*, and chart_* roles dynamically.
    """
    thread_local.current_slide_data = content
    try:
        zones = schema_slide.get("zones", [])
        if not zones:
            print(f"[FILLER_V2] Slide {slide_idx + 1}: no zones in schema")
            return

        # Map unified "body" key from frontend editor back to individual body1, body2... roles
        body_text = content.get("body")
        if body_text and isinstance(body_text, str):
            lines = [line.strip() for line in body_text.split("\n") if line.strip()]
            for i, line in enumerate(lines):
                content[f"body{i+1}"] = line
            # If there are fewer lines than body shapes in template, blank out the rest
            total_body_zones = sum(1 for z in zones if z.get("role", "").startswith("body"))
            for i in range(len(lines), total_body_zones):
                content[f"body{i+1}"] = ""

        filled_shape_ids = set()

        for zone in zones:
            role = zone.get("role", "other")
            role_lower = role.lower()

            is_text_role = (
                role in ("title", "subtitle") or 
                role_lower.startswith("body") or 
                role_lower.startswith("title") or 
                role_lower.startswith("subtitle")
            )
            is_table_role = role_lower.startswith("table_")
            is_chart_role = role_lower.startswith("chart_")
            is_image_role = role_lower.startswith("image")
            is_icon_role = role_lower.startswith("icon")

            if not is_text_role and not is_image_role and not is_table_role and not is_chart_role and not is_icon_role:
                continue

            # Get content value for this role
            val = _get_value_for_role(role, content)

            # Skip empty content — never overwrite a shape with empty text
            if val is None:
                continue
            if isinstance(val, str) and not val.strip():
                continue

            # Find the matching shape
            shape = _find_shape_v2(slide, zone, slide_width, slide_height)
            if shape is None:
                print(f"[FILLER_V2] WARNING: Slide {slide_idx + 1}, role '{role}' — shape not found, skipping")
                continue

            # Table shapes are shared across cells, so do not block table cells by shape_id
            if not is_table_role and shape.shape_id in filled_shape_ids:
                print(f"[FILLER_V2] WARNING: Slide {slide_idx + 1}, shape {shape.shape_id} already filled, skipping role '{role}'")
                continue

            # Fill based on type
            if is_image_role or is_icon_role:
                _replace_image_in_shape(slide, shape, val, is_icon=is_icon_role, role=role)
            elif is_table_role:
                row_idx = zone.get("row_idx")
                col_idx = zone.get("col_idx")
                if row_idx is not None and col_idx is not None and hasattr(shape, "has_table") and shape.has_table:
                    try:
                        cell = shape.table.cell(row_idx, col_idx)
                        max_chars = zone.get("max_chars", 200)
                        val = fit_text_with_llm_rewrites(val, max_chars, role)
                        _fill_text_preserving_format(cell, val)
                        if len(val) > max_chars:
                            ratio = max_chars / len(val)
                            _reduce_font_size_in_shape(cell, ratio=ratio)
                    except Exception as e:
                        print(f"[FILLER_V2] Error filling table cell '{role}': {e}")
            elif is_chart_role:
                if hasattr(shape, "has_chart") and shape.has_chart:
                    try:
                        chart = shape.chart
                        chart_data_dict = None
                        if isinstance(val, dict):
                            chart_data_dict = val
                        elif isinstance(val, str):
                            try:
                                chart_data_dict = json.loads(val)
                            except Exception:
                                import re
                                match = re.search(r'\{.*\}', val, re.DOTALL)
                                if match:
                                    try:
                                        chart_data_dict = json.loads(match.group(0))
                                    except Exception:
                                        pass
                        if chart_data_dict:
                            categories = chart_data_dict.get("categories", [])
                            series_list = chart_data_dict.get("series", [])
                            if categories and series_list:
                                from pptx.chart.data import CategoryChartData
                                chart_data = CategoryChartData()
                                chart_data.categories = categories
                                for s in series_list:
                                    name = s.get("name", "")
                                    values = s.get("values", [])
                                    chart_data.add_series(name, tuple(values))
                                chart.replace_data(chart_data)
                                print(f"[FILLER_V2] Replaced chart data for role '{role}'")
                    except Exception as e:
                        print(f"[FILLER_V2] Error filling chart '{role}': {e}")
            elif is_text_role and hasattr(shape, "text_frame"):
                max_chars = zone.get("max_chars", 2000)
                
                # 1. Fit text using iterative LLM rewrites to ensure complete sentences and compliance
                val = fit_text_with_llm_rewrites(val, max_chars, role)
                
                # 2. Write to shape preserving formats
                _fill_text_preserving_format(shape, val)
                
                # 3. Reduce font size slightly if it still exceeds max_chars as a last resort
                if len(val) > max_chars:
                    ratio = max_chars / len(val)
                    _reduce_font_size_in_shape(shape, ratio=ratio)

            if not is_table_role:
                filled_shape_ids.add(shape.shape_id)
            print(f"[FILLER_V2] Slide {slide_idx + 1}: filled role '{role}' → shape_id={shape.shape_id}")
    finally:
        thread_local.current_slide_data = None


# ═══════════════════════════════════════════════════════════════════════
# SHAPE MATCHING — strict priority: shape_id → placeholder_idx → name → coords
# ═══════════════════════════════════════════════════════════════════════

def _get_all_slide_shapes(slide):
    """Recursively collect all shapes from a slide, including children of GROUP shapes."""
    result = []
    def _collect(shapes_collection):
        for shape in shapes_collection:
            try:
                if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                    _collect(shape.shapes)
                else:
                    result.append(shape)
            except Exception:
                result.append(shape)
    _collect(slide.shapes)
    return result


def _find_shape_v2(slide, zone, slide_width, slide_height):
    """Find the shape matching a zone definition.

    Priority:
    1. shape_id match
    2. placeholder_idx match
    3. name match
    4. coordinate similarity (< 5.0 Manhattan distance)

    Returns None if no match found. Never falls back to arbitrary shapes.
    """
    all_shapes = _get_all_slide_shapes(slide)

    target_shape_id = zone.get("shape_id")
    target_ph_idx = zone.get("placeholder_idx")
    target_name = zone.get("name", "")

    # Priority 1: shape_id
    if target_shape_id is not None:
        for s in all_shapes:
            if s.shape_id == target_shape_id:
                return s

    # Priority 2: placeholder_idx
    if target_ph_idx is not None:
        for s in all_shapes:
            if s.is_placeholder:
                try:
                    if s.placeholder_format.idx == target_ph_idx:
                        return s
                except Exception:
                    pass

    # Priority 3: name
    if target_name:
        for s in all_shapes:
            if s.name == target_name:
                return s

    # Priority 4: coordinate matching
    target_left = zone.get("left_pct", -1)
    target_top = zone.get("top_pct", -1)
    target_w = zone.get("width_pct", -1)
    target_h = zone.get("height_pct", -1)

    if target_left >= 0 and target_top >= 0:
        best_shape = None
        best_dist = 5.0  # threshold

        for s in all_shapes:
            try:
                s_left = (s.left / slide_width) * 100 if slide_width else 0
                s_top = (s.top / slide_height) * 100 if slide_height else 0
                s_w = (s.width / slide_width) * 100 if slide_width else 0
                s_h = (s.height / slide_height) * 100 if slide_height else 0

                dist = (
                    abs(s_left - target_left) +
                    abs(s_top - target_top) +
                    abs(s_w - target_w) +
                    abs(s_h - target_h)
                )

                if dist < best_dist:
                    best_dist = dist
                    best_shape = s
            except Exception:
                continue

        if best_shape:
            return best_shape

    return None


# ═══════════════════════════════════════════════════════════════════════
# CONTENT VALUE LOOKUP
# ═══════════════════════════════════════════════════════════════════════

def _get_value_for_role(role: str, content: dict):
    """Look up the content value for a given role.

    Checks:
    1. Direct key: content[role]
    2. zone_content: content["zone_content"][role]
    3. Image keyword: content["{role}_keyword"] for image roles

    Returns None if nothing found.
    """
    # Direct key
    val = content.get(role)
    if val is not None:
        if isinstance(val, dict):
            return val
        if isinstance(val, list):
            return "\n".join(str(v) for v in val)
        return str(val)

    # zone_content sub-dict
    zc = content.get("zone_content")
    if isinstance(zc, dict):
        val = zc.get(role)
        if val is not None:
            if isinstance(val, dict):
                return val
            if isinstance(val, list):
                return "\n".join(str(v) for v in val)
            return str(val)

    # Image/Icon keyword lookup
    role_lower = role.lower()
    if role_lower.startswith("image") or role_lower.startswith("icon"):
        # Try {role}_keyword
        kw = content.get(f"{role}_keyword")
        if kw:
            return str(kw)
        # Try generic image_keyword or icon_keyword fallback
        fallback_keys = ["image_keyword", "icon_keyword"] if role_lower.startswith("image") else ["icon_keyword", "image_keyword"]
        for fk in fallback_keys:
            kw = content.get(fk)
            if kw:
                return str(kw)

    return None


# ═══════════════════════════════════════════════════════════════════════
# TEXT FILLING — preserves paragraph structure and run formatting
# ═══════════════════════════════════════════════════════════════════════

def _fill_text_preserving_format(shape, text_value: str):
    """Fill text into a shape while preserving existing formatting.

    Rules:
    1. Preserve existing paragraph count and structure
    2. Write into existing runs (preserves font properties)
    3. Only the first run of each paragraph is modified
    4. Extra paragraphs beyond text lines get empty string
    5. Never add new paragraphs if the template didn't have them
    """
    if not hasattr(shape, "text_frame"):
        return

    tf = shape.text_frame

    # Split text into lines
    if isinstance(text_value, list):
        lines = [str(l) for l in text_value]
    else:
        lines = text_value.split("\n") if "\n" in text_value else [text_value]

    # Remove empty trailing lines
    while lines and not lines[-1].strip():
        lines.pop()

    if not lines:
        return

    paras = list(tf.paragraphs)
    P = len(paras)

    if P == 0:
        # No paragraphs at all — add one
        p = tf.paragraphs[0] if tf.paragraphs else tf.add_paragraph()
        run = p.add_run()
        run.text = " ".join(lines)
        return

    # Map lines to paragraphs
    if P == 1:
        # Single paragraph: join all lines
        _set_paragraph_text(paras[0], " ".join(lines))
    elif P >= len(lines):
        # More paragraphs than lines: fill available, blank the rest
        for i, para in enumerate(paras):
            if i < len(lines):
                _set_paragraph_text(para, lines[i])
            else:
                _set_paragraph_text(para, "")
    else:
        # More lines than paragraphs: first P-1 get one line each, last gets the rest
        for i in range(P - 1):
            _set_paragraph_text(paras[i], lines[i])
        remaining = " ".join(lines[P - 1:])
        _set_paragraph_text(paras[P - 1], remaining)


def _set_paragraph_text(para, text: str):
    """Set text in a paragraph by modifying the first run.
    Preserves all run formatting (font, color, bold, etc.).
    """
    if para.runs:
        # Write into existing first run — formatting preserved
        para.runs[0].text = text
        # Clear any additional runs to avoid leftover text
        for run in para.runs[1:]:
            run.text = ""
    else:
        # No runs exist — add one (won't have template formatting)
        run = para.add_run()
        run.text = text


# ═══════════════════════════════════════════════════════════════════════
# IMAGE HANDLING — in-place blob replacement
# ═══════════════════════════════════════════════════════════════════════

def _replace_image_in_shape(slide, shape, image_keyword: str, is_icon: bool = False, role: str = None):
    """Replace an image in a shape.
    
    First tries in-place blob replacement (preserves cropping, shape, styles).
    If the shape doesn't have an embedded image blip (e.g. vector icon/shape),
    falls back to deleting the shape and adding a new picture at the exact same bounds.
    """
    # ── Check if the shape is actually a picture shape, picture placeholder, or has an image blip ──
    from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
    
    is_picture_shape = (shape.shape_type == MSO_SHAPE_TYPE.PICTURE)
    is_picture_placeholder = False
    if shape.is_placeholder:
        try:
            if shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
                is_picture_placeholder = True
        except Exception:
            pass
            
    has_blip = False
    try:
        blip = shape.element.find('.//' + qn('a:blip'))
        if blip is not None:
            has_blip = True
    except Exception:
        pass
        
    # If it is not a picture shape, picture placeholder, or shape with image fill,
    # then it is a vector design shape/graphic. We MUST preserve it as-is (unless it is an icon role)!
    if not (is_picture_shape or is_picture_placeholder or has_blip or is_icon):
        print(f"[FILLER_V2] Preserving design shape {shape.shape_id} (type={shape.shape_type}), skipping image replacement.")
        return

    # Download the image
    image_path = _download_image_for_template(image_keyword, is_icon=is_icon, role=role)
    if not image_path:
        print(f"[FILLER_V2] WARNING: Could not download image for '{image_keyword}'")
        return

    # Find the blip element (embedded image reference)
    blip = None
    try:
        blip = shape.element.find('.//' + qn('a:blip'))
    except Exception:
        pass
    if blip is None:
        try:
            blip = shape._element.find('.//' + qn('a:blip'))
        except Exception:
            pass

    if blip is not None:
        # 1. In-place blob replacement
        try:
            rId = blip.get(qn('r:embed'))
            if rId:
                image_part = slide.part.rels[rId].target_part
                with open(image_path, 'rb') as f:
                    image_part.blob = f.read()
                print(f"[FILLER_V2] Replaced image blob in shape {shape.shape_id} with '{image_keyword}'")
                return
        except Exception as e:
            print(f"[FILLER_V2] WARNING: In-place replacement failed, trying fallback: {e}")

    # 2. Fallback: delete shape and add picture
    try:
        left = shape.left
        top = shape.top
        width = shape.width
        height = shape.height
        
        _replace_shape_with_picture(slide, shape, image_path, left, top, width, height)
        print(f"[FILLER_V2] Deleted shape {shape.shape_id} and added new picture with z-order preservation for '{image_keyword}'")
    except Exception as e:
        print(f"[FILLER_V2] ERROR: Fallback image placement failed for shape {shape.shape_id}: {e}")


def _replace_shape_with_picture(slide, shape, image_path, left, top, width, height):
    """Replaces a shape with a new picture at the exact same z-order position in the XML tree."""
    old_el = shape._element
    parent = old_el.getparent()
    if parent is not None:
        idx = parent.index(old_el)
        new_pic = slide.shapes.add_picture(image_path, left, top, width, height)
        new_el = new_pic._element
        slide.shapes._spTree.remove(new_el)
        parent.insert(idx, new_el)
        parent.remove(old_el)
    else:
        slide.shapes.add_picture(image_path, left, top, width, height)


def _verify_and_shorten_text_frames(prs, template_schema: dict, slide_contents: list):
    """Post-fill verification pass.
    Checks every filled text zone and table cell against its max_chars limit.
    If it exceeds max_chars, it triggers fit_text_with_llm_rewrites.
    """
    slides_schema = template_schema.get("slides", [])
    for slide_idx, slide in enumerate(prs.slides):
        if slide_idx >= len(slides_schema):
            continue
        schema_slide = slides_schema[slide_idx]
        zones = schema_slide.get("zones", [])
        
        for zone in zones:
            role = zone.get("role", "other")
            role_lower = role.lower()
            
            is_text = role in ("title", "subtitle") or role_lower.startswith("body") or role_lower.startswith("title") or role_lower.startswith("subtitle")
            is_table = role_lower.startswith("table_")
            
            if not is_text and not is_table:
                continue
                
            max_chars = zone.get("max_chars", 200)
            
            # Find the shape
            shape = _find_shape_v2(slide, zone, prs.slide_width, prs.slide_height)
            if shape is None:
                continue
                
            if is_table:
                row_idx = zone.get("row_idx")
                col_idx = zone.get("col_idx")
                if row_idx is not None and col_idx is not None and hasattr(shape, "has_table") and shape.has_table:
                    try:
                        cell = shape.table.cell(row_idx, col_idx)
                        text = cell.text_frame.text.strip()
                        if len(text) > max_chars:
                            print(f"[VERIFY] Table cell {role} (len={len(text)}) exceeds limit={max_chars}. Shortening...")
                            new_text = fit_text_with_llm_rewrites(text, max_chars, role)
                            _fill_text_preserving_format(cell, new_text)
                            if len(new_text) > max_chars:
                                ratio = max_chars / len(new_text)
                                _reduce_font_size_in_shape(cell, ratio=ratio)
                    except Exception as e:
                        print(f"[VERIFY] Error verifying table cell {role}: {e}")
            else:
                if hasattr(shape, "text_frame"):
                    text = shape.text_frame.text.strip()
                    if len(text) > max_chars:
                        print(f"[VERIFY] Text shape {role} (len={len(text)}) exceeds limit={max_chars}. Shortening...")
                        new_text = fit_text_with_llm_rewrites(text, max_chars, role)
                        _fill_text_preserving_format(shape, new_text)
                        if len(new_text) > max_chars:
                            ratio = max_chars / len(new_text)
                            _reduce_font_size_in_shape(shape, ratio=ratio)


def _download_image_for_template(keyword: str, is_icon: bool = False, role: str = None):
    """Download an image using the visual engine or resolve local/remote URLs."""
    if not keyword:
        return None
        
    keyword_str = str(keyword).strip()
    
    # Check thread-local cache for pre-downloaded image first
    if hasattr(thread_local, "current_slide_data") and thread_local.current_slide_data:
        preloaded = thread_local.current_slide_data.get("_preloaded_image")
        if preloaded:
            if isinstance(preloaded, dict) and role:
                path = preloaded.get(role)
                if path and os.path.exists(path):
                    print(f"[FILLER_V2] Using thread-local preloaded image for role '{role}': {path}")
                    return path
            elif isinstance(preloaded, str) and os.path.exists(preloaded):
                print(f"[FILLER_V2] Using thread-local preloaded image: {preloaded}")
                return preloaded
    
    # 1. Handle local static URLs (from the web app preview)
    if "/static/" in keyword_str:
        filename = keyword_str.split("/static/")[-1]
        import tempfile
        local_path = os.path.join(tempfile.gettempdir(), filename)
        if os.path.exists(local_path):
            return local_path

    # 2. Handle remote URLs
    if keyword_str.startswith("http://") or keyword_str.startswith("https://"):
        try:
            import requests
            import tempfile
            import uuid
            resp = requests.get(keyword_str, timeout=15)
            resp.raise_for_status()
            ext = ".jpg"
            if "png" in resp.headers.get("Content-Type", ""):
                ext = ".png"
            local_path = os.path.join(tempfile.gettempdir(), f"downloaded_{uuid.uuid4().hex[:10]}{ext}")
            with open(local_path, "wb") as f:
                f.write(resp.content)
            return local_path
        except Exception as e:
            print(f"[FILLER_V2] Failed to download URL '{keyword_str}': {e}")
            return None

    # 3. Handle standard search keyword
    try:
        from services.visual_engine import get_hero_image
        if is_icon:
            # Look for transparent vector/illustration on Pixabay first
            try:
                from services.visual_engine import PIXABAY_API_KEY
                import requests
                import uuid
                if PIXABAY_API_KEY:
                    resp = requests.get(
                        "https://pixabay.com/api/",
                        params={
                            "key": PIXABAY_API_KEY,
                            "q": keyword_str,
                            "image_type": "vector",
                            "per_page": 3,
                            "safesearch": "true"
                        },
                        timeout=8
                    )
                    if resp.status_code == 200:
                        hits = resp.json().get("hits", [])
                        if hits:
                            img_url = hits[0].get("vectorSVG") or hits[0].get("largeImageURL") or hits[0].get("webformatURL")
                            if img_url:
                                r = requests.get(img_url, timeout=10)
                                if r.status_code == 200:
                                    ext = ".png" if "png" in img_url or "png" in r.headers.get("Content-Type", "") else ".svg" if "svg" in img_url else ".jpg"
                                    import tempfile
                                    p = os.path.join(tempfile.gettempdir(), f"icon_{uuid.uuid4().hex[:10]}{ext}")
                                    with open(p, "wb") as f:
                                        f.write(r.content)
                                    print(f"[FILLER_V2] Downloaded transparent Pixabay icon: {p}")
                                    return p
            except Exception as ex:
                print(f"[FILLER_V2] Pixabay vector search failed: {ex}")

            # Fallback search query
            search_query = f"{keyword_str} flat icon png"
            return get_hero_image(search_query)
        return get_hero_image(keyword_str)
    except Exception as e:
        print(f"[FILLER_V2] WARNING: Image download failed for '{keyword_str}': {e}")
        return None


# ═══════════════════════════════════════════════════════════════════════
# TEXT CAPACITY & TRUNCATION
# ═══════════════════════════════════════════════════════════════════════

def _get_font_size(shape) -> float:
    """Extract font size from the first run of the first paragraph."""
    try:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.font.size:
                        return float(run.font.size.pt)
    except Exception:
        pass
    return 14.0  # default


def _estimate_capacity(shape, font_size: float) -> int:
    """Estimate how many words fit in a shape based on its physical size."""
    try:
        if shape.has_text_frame:
            width_in = shape.width.inches
            height_in = shape.height.inches
            area = width_in * height_in
            f_sz = font_size if font_size else 14.0
            word_area = (f_sz ** 2 * 3.6) / 5184
            capacity = int(area / word_area) if word_area > 0 else 0
            return max(5, min(500, capacity))
    except Exception:
        pass
    # Fallback: estimate from font size
    if font_size >= 24:
        return 10
    elif font_size >= 18:
        return 20
    else:
        return 35


# _truncate_text removed to prevent hard truncation


# ═══════════════════════════════════════════════════════════════════════
# SHAPE STATE VALIDATION
# ═══════════════════════════════════════════════════════════════════════

def _capture_shape_state(prs) -> list:
    """Capture initial shape positions/sizes for post-fill validation."""
    state = []
    for slide_idx, slide in enumerate(prs.slides):
        slide_shapes = []
        for shape in slide.shapes:
            slide_shapes.append({
                "shape_id": shape.shape_id,
                "left": shape.left,
                "top": shape.top,
                "width": shape.width,
                "height": shape.height,
            })
        state.append(slide_shapes)
    return state


def _validate_shape_state(prs, initial_state: list):
    """Compare current shape state against initial. Log warnings only."""
    for slide_idx, slide in enumerate(prs.slides):
        if slide_idx >= len(initial_state):
            break
        initial_shapes = {s["shape_id"]: s for s in initial_state[slide_idx]}
        for shape in slide.shapes:
            init = initial_shapes.get(shape.shape_id)
            if not init:
                continue
            if shape.left != init["left"] or shape.top != init["top"]:
                print(f"[FILLER_V2] WARNING: Slide {slide_idx + 1}, shape {shape.shape_id} position changed!")
            if shape.width != init["width"] or shape.height != init["height"]:
                print(f"[FILLER_V2] WARNING: Slide {slide_idx + 1}, shape {shape.shape_id} size changed!")


# ═══════════════════════════════════════════════════════════════════════
# SLIDE DELETION
# ═══════════════════════════════════════════════════════════════════════

def _delete_slide(prs, slide_idx):
    """Delete a slide by index using XML manipulation."""
    rId = prs.slides._sldIdLst[slide_idx].get(qn('r:id'))
    prs.part.drop_rel(rId)
    sldId = prs.slides._sldIdLst[slide_idx]
    prs.slides._sldIdLst.remove(sldId)


def fix_presentation_complex_scripts(prs):
    """Post-processor that ensures all runs containing complex script text (non-ASCII characters)
    have the complex script typeface (<a:cs>) properly set in their OOXML properties.
    This prevents PowerPoint from spacing out/jumbling Indic, Arabic, and other complex characters.
    """
    from pptx.oxml.ns import qn
    from pptx.oxml.xmlchemy import OxmlElement
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    def _get_all_shapes_recursive(shapes_collection):
        result = []
        for shape in shapes_collection:
            try:
                if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                    result.extend(_get_all_shapes_recursive(shape.shapes))
                else:
                    result.append(shape)
            except Exception:
                result.append(shape)
        return result

    def is_rtl_text(text: str) -> bool:
        if not text:
            return False
        for char in text:
            o = ord(char)
            # Arabic/Persian/Urdu: 0x0600-0x06FF, 0x0750-0x077F, 0x08A0-0x08FF
            # Hebrew: 0x0590-0x05FF
            # Thaana: 0x0780-0x07BF
            # Samaritan: 0x0800-0x083F
            if (1424 <= o <= 1791) or (1872 <= o <= 1919) or (2208 <= o <= 2303) or (1920 <= o <= 1983) or (2048 <= o <= 2111):
                return True
        return False

    def _process_text_frame(tf):
        from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN

        # 1. Check if the text frame contains any complex script or RTL text
        has_complex = False
        has_rtl = False
        for para in tf.paragraphs:
            para_text = "".join(run.text for run in para.runs if run.text)
            if any(ord(c) > 127 for c in para_text):
                has_complex = True
            if is_rtl_text(para_text):
                has_rtl = True

        # Apply auto-fit and word wrap to prevent text overflow and clipping
        if has_complex:
            try:
                tf.word_wrap = True
                tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            except Exception as e:
                print(f"[COMPLEX_SCRIPT] Warning: failed to apply auto size: {e}")

        # 2. Process paragraphs and runs
        for para in tf.paragraphs:
            para_text = "".join(run.text for run in para.runs if run.text)
            if is_rtl_text(para_text):
                try:
                    pPr = para._p.get_or_add_pPr()
                    pPr.set('rtl', '1')
                    # Align RTL text to the right if it's default/left-aligned
                    if para.alignment is None or para.alignment == PP_ALIGN.LEFT:
                        para.alignment = PP_ALIGN.RIGHT
                except Exception as e:
                    print(f"[COMPLEX_SCRIPT] Warning: failed to apply RTL paragraph properties: {e}")

            for run in para.runs:
                text = run.text
                if not text:
                    continue
                if any(ord(c) > 127 for c in text):
                    try:
                        rPr = run.font._rPr
                        if rPr is not None:
                            font_name = run.font.name
                            if not font_name:
                                font_name = "Calibri"  # Fallback
                            
                            cs_font_name = font_name
                            if any(2304 <= ord(c) <= 3583 for c in text):
                                cs_font_name = "Nirmala UI"
                            
                            # Ensure <a:latin> element exists
                            latin = rPr.find(qn('a:latin'))
                            if latin is None:
                                latin = OxmlElement('a:latin')
                                rPr.append(latin)
                            if not latin.get('typeface'):
                                latin.set('typeface', font_name)
                            
                            # Ensure <a:cs> (complex script) element exists and matches typeface
                            cs = rPr.find(qn('a:cs'))
                            if cs is None:
                                cs = OxmlElement('a:cs')
                                rPr.append(cs)
                            cs.set('typeface', cs_font_name)
                            
                            # Ensure <a:ea> (East Asian) element exists
                            ea = rPr.find(qn('a:ea'))
                            if ea is None:
                                ea = OxmlElement('a:ea')
                                rPr.append(ea)
                            if not ea.get('typeface'):
                                ea.set('typeface', font_name)
                    except Exception as e:
                        print(f"[COMPLEX_SCRIPT] Warning: failed to apply complex script font to run: {e}")

    for slide in prs.slides:
        all_shapes = _get_all_shapes_recursive(slide.shapes)
        for shape in all_shapes:
            if shape.has_text_frame:
                _process_text_frame(shape.text_frame)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        _process_text_frame(cell.text_frame)


def _download_all_template_images(slide_contents: list, topic: str) -> list:
    """Pre-download all images and icons for template slides in parallel.
    Returns a list of dicts (one per slide) mapping role -> local_path.
    """
    import time
    from concurrent.futures import ThreadPoolExecutor, as_completed
    from services.visual_engine import get_hero_image, PIXABAY_API_KEY
    import requests
    import tempfile
    import uuid

    preloaded_maps = [{} for _ in slide_contents]
    tasks = []  # (slide_index, role, keyword, is_icon)

    for idx, content in enumerate(slide_contents):
        if not content:
            continue
        for key, val in content.items():
            if not val:
                continue
            # Look for image/icon keywords
            key_lower = key.lower()
            if key_lower.endswith("_keyword"):
                # E.g. image1_keyword -> image1
                role = key[:-8]
                is_icon = role.lower().startswith("icon")
                tasks.append((idx, role, str(val).strip(), is_icon))

    if not tasks:
        return preloaded_maps

    def download_one(slide_idx, role, keyword, is_icon):
        try:
            # 1. URL download
            if keyword.startswith("http://") or keyword.startswith("https://"):
                resp = requests.get(keyword, timeout=15)
                resp.raise_for_status()
                ext = ".jpg"
                if "png" in resp.headers.get("Content-Type", ""):
                    ext = ".png"
                local_path = os.path.join(tempfile.gettempdir(), f"downloaded_{uuid.uuid4().hex[:10]}{ext}")
                with open(local_path, "wb") as f:
                    f.write(resp.content)
                return slide_idx, role, local_path

            # 2. Pixabay Icon
            if is_icon:
                try:
                    if PIXABAY_API_KEY:
                        resp = requests.get(
                            "https://pixabay.com/api/",
                            params={
                                "key": PIXABAY_API_KEY,
                                "q": keyword,
                                "image_type": "vector",
                                "per_page": 3,
                                "safesearch": "true"
                            },
                            timeout=8
                        )
                        if resp.status_code == 200:
                            hits = resp.json().get("hits", [])
                            if hits:
                                img_url = hits[0].get("vectorSVG") or hits[0].get("largeImageURL") or hits[0].get("webformatURL")
                                if img_url:
                                    r = requests.get(img_url, timeout=10)
                                    if r.status_code == 200:
                                        ext = ".png" if "png" in img_url or "png" in r.headers.get("Content-Type", "") else ".svg" if "svg" in img_url else ".jpg"
                                        p = os.path.join(tempfile.gettempdir(), f"icon_{uuid.uuid4().hex[:10]}{ext}")
                                        with open(p, "wb") as f:
                                            f.write(r.content)
                                        return slide_idx, role, p
                except Exception as ex:
                    print(f"[FILLER_V2] Parallel Pixabay vector search failed: {ex}")

                # Fallback icon keyword search
                search_query = f"{keyword} flat icon png"
                path = get_hero_image(search_query)
                return slide_idx, role, path

            # 3. Standard hero image
            path = get_hero_image(keyword)
            return slide_idx, role, path
        except Exception as e:
            print(f"[FILLER_V2] Error pre-downloading '{keyword}' for slide {slide_idx+1} role '{role}': {e}")
            return slide_idx, role, None

    print(f"[TIMING] Starting parallel template image download of {len(tasks)} items...")
    start_time = time.time()

    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {
            executor.submit(download_one, s_idx, role, kw, is_icon): (s_idx, role)
            for s_idx, role, kw, is_icon in tasks
        }
        for future in as_completed(futures):
            slide_idx, role, path = future.result()
            if path:
                preloaded_maps[slide_idx][role] = path
                print(f"[FILLER_V2] OK Preloaded image for slide {slide_idx+1} role '{role}': {path}")

    elapsed = time.time() - start_time
    print(f"[TIMING] Template images preloaded in {elapsed:.2f}s")
    return preloaded_maps


def _rewrite_shorter_with_llm(text: str, max_chars: int, role: str) -> str:
    """Use the LLM to rewrite text shorter so it fits the character limit, returning a complete sentence or complete synonym."""
    from services.llm_client import call_llm
    
    system_prompt = (
        "You are an expert editor. You receive a text snippet, a target character limit, and its slide role.\n"
        "Your task is to rewrite the text so it is strictly shorter than the limit, while preserving the key meaning, "
        "retaining a professional tone, and ensuring it consists of fully spelled, complete, grammatically correct words.\n"
        "CRITICAL RULES:\n"
        "1. Never return a response containing an ellipsis ('...') or any sliced, partial, or truncated words (like 'Nutri' instead of 'Nutrition' or 'strat' instead of 'strategy').\n"
        "2. All words in your response must be complete dictionary words. No abbreviations or clipped words.\n"
        "3. If the character limit is too small to fit the original word, rewrite it using a shorter synonym that fits completely (e.g., if the limit is 8 characters, instead of truncating 'Education' to 'Edu...', use a complete synonym like 'School' or 'Study' or 'Learn').\n"
        "4. Always output complete sentences or complete phrases with proper ending punctuation if it is a sentence.\n"
        "5. Do NOT slice mid-sentence or mid-word. Return ONLY the rewritten text, with no markdown fences, no quotes, and no explanation."
    )
    user_prompt = f"Rewrite this text for the role '{role}' to be strictly under {max_chars} characters (including spaces):\n\n\"{text}\""
    
    try:
        rewritten = call_llm(system_prompt, user_prompt)
        rewritten = rewritten.strip().strip('"').strip("'").strip()
        return rewritten
    except Exception as e:
        print(f"[REWRITE] LLM rewrite failed: {e}")
        return text


def fit_text_with_llm_rewrites(text: str, max_chars: int, role: str) -> str:
    """Fits text using iterative LLM rewrites to ensure it fits max_chars without word truncation."""
    if not text:
        return text
        
    current_text = text.strip()
    if "..." in current_text:
        current_text = current_text.replace("...", "").strip()
        
    words = current_text.split()
    is_short_or_table = (
        role.lower().startswith("table_") or 
        len(words) < 4 or 
        (current_text and (current_text[-1].isdigit() or current_text[-1] in ('%', '$', ')', ']', '}')))
    )

    # First check: if the text already fits and has no ellipsis, return it directly
    if len(current_text) <= max_chars:
        if is_short_or_table:
            return current_text
        if current_text and current_text[-1] in ('.', '!', '?', '"', "'"):
            return current_text
        if len(current_text) + 1 <= max_chars:
            return current_text + "."
        return current_text

    for attempt in range(3):
        print(f"[REWRITE] Text for role '{role}' (len={len(current_text)}) exceeds limit={max_chars}. Rewrite attempt {attempt+1}...")
        rewritten = _rewrite_shorter_with_llm(current_text, max_chars, role)
        
        # Scrub any trailing or embedded ellipses
        if "..." in rewritten:
            rewritten = rewritten.replace("...", "").strip()
            
        if len(rewritten) <= max_chars:
            if is_short_or_table:
                return rewritten
            if rewritten and rewritten[-1] not in ('.', '!', '?'):
                if len(rewritten) + 1 <= max_chars:
                    rewritten += "."
            return rewritten
            
        current_text = rewritten
        
    if "..." in current_text:
        current_text = current_text.replace("...", "").strip()
    return current_text


def _reduce_font_size_in_shape(shape, ratio=0.85):
    """Reduce the font size of all runs in all paragraphs of a shape by a ratio."""
    if not hasattr(shape, "text_frame"):
        return
    # Cap ratio to a reasonable range [0.5, 0.95]
    scale_factor = max(0.5, min(0.95, ratio))
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            if run.font.size:
                try:
                    old_size = run.font.size.pt
                    new_size = max(6.0, old_size * scale_factor)
                    run.font.size = Pt(new_size)
                    print(f"[FONT_REDUCTION] Scaled font size from {old_size}pt to {new_size:.1f}pt for shape {shape.shape_id}")
                except Exception as e:
                    print(f"[FONT_REDUCTION] Failed to change font size: {e}")

