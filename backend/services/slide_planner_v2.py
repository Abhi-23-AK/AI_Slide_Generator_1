"""
Slide Planner V2 — Template-Aware Content Generation
=====================================================
Generates LLM-driven content that maps 1:1 to template zone roles.

This is the V2 pipeline planner. It:
  - Only sends allowed roles (title, subtitle, body1-4, image1-3) to the LLM
  - Returns flat-dict slides with template_mode = True
  - Enforces strict word limits based on zone capacity
  - Generates image keywords for every image placeholder
  - Never generates content for icons, decorative shapes, footers, etc.
"""

import json
import re
from services.llm_client import call_llm


# ─── Tone descriptions ───────────────────────────────────────────────
TONE_DESCRIPTIONS = {
    "professional":  "Formal, corporate tone with clear and concise language.",
    "technical":     "Precise, data-driven tone suitable for engineering or scientific audiences.",
    "lavish":        "Luxurious, premium tone with elegant and sophisticated language.",
    "academic":      "Scholarly, well-referenced tone for educational or research contexts.",
    "creative":      "Bold, imaginative tone with vivid language and fresh metaphors.",
    "education":     "Friendly, accessible tone optimized for teaching and learning.",
    "ecommerce":     "Persuasive, benefits-focused tone for products and conversions.",
    "neumorphism":   "Modern, tech-forward tone with clean minimal language.",
}

# ─── Roles the LLM is allowed to generate content for ────────────────
ALLOWED_CONTENT_ROLES = {
    "title", "subtitle",
    "body1", "body2", "body3", "body4",
}

# Roles that should produce image/icon keywords instead of text
IMAGE_ROLES = {
    "image1", "image2", "image3",
    "image_left", "image_right", "hero_image",
}

# Roles that must NEVER reach the LLM
SKIP_KEYWORDS = (
    "logo", "vector", "symbol", "icon",
    "shape", "decorative", "other", "footer", "caption", "label",
    "chart", "table", "arrow", "note", "source", "reference",
    "separator", "line", "border",
)

# ─── Word limits per role ─────────────────────────────────────────────
DEFAULT_WORD_LIMITS = {
    "title":    8,
    "subtitle": 12,
    "body1":    18,
    "body2":    18,
    "body3":    18,
    "body4":    18,
}


def is_valid_icon_role(role: str) -> bool:
    """Check if the role is a valid content icon role (e.g. icon1-icon10 or ending with _icon)."""
    role_lower = role.lower()
    # Skip any icon roles that contain other forbidden decorative keywords
    for k in SKIP_KEYWORDS:
        if k != "icon" and k in role_lower:
            return False
            
    if role_lower.endswith("_icon") or "_icon_" in role_lower:
        return True
    m = re.match(r'^icon(\d+)$', role_lower)
    if m:
        val = int(m.group(1))
        return val <= 10
    return False


def _extract_icon_keyword_from_slide(slide_content: dict, role: str) -> str:
    """Extract a simple 1-2 word noun keyword from slide content related to the icon role."""
    role_lower = role.lower()
    prefix = None
    if "_icon" in role_lower:
        prefix = role_lower.split("_icon")[0]
    
    m = re.search(r'\d+', role_lower)
    digit = m.group(0) if m else None
    
    candidates = []
    if prefix:
        for k, v in slide_content.items():
            if k.lower().startswith(prefix) and k.lower() != role_lower and isinstance(v, str) and v.strip():
                candidates.append(v)
    
    if digit:
        for k, v in slide_content.items():
            if digit in k and k.lower() != role_lower and isinstance(v, str) and v.strip():
                candidates.append(v)
                
    if not candidates:
        for k, v in slide_content.items():
            if k not in ("slide_index", "layout") and not k.endswith("_keyword") and isinstance(v, str) and v.strip():
                candidates.append(v)
                
    stop_words = {"the", "a", "an", "and", "or", "but", "in", "on", "at", "to", "for", "with", "by", "of", "our", "your", "their", "its", "we", "us", "they", "it", "this", "that", "these", "those"}
    words = []
    for text in candidates:
        cleaned = re.sub(r'[^a-zA-Z0-9\s]', '', text)
        for w in cleaned.lower().split():
            if w not in stop_words and len(w) > 2:
                words.append(w)
                
    seen = set()
    unique_words = []
    for w in words:
        if w not in seen:
            seen.add(w)
            unique_words.append(w)
            
    if unique_words:
        return " ".join(unique_words[:2])
    return ""


def generate_outline_from_template_v2(
    topic: str,
    template_schema: dict,
    tone: str,
    source_text: str = "",
) -> dict:
    """Generate a content outline for a V2 template schema.

    Args:
        topic: Presentation topic string.
        template_schema: V2 schema dict from analyze_template_v2().
        tone: Tone name (matched against TONE_DESCRIPTIONS).
        source_text: Optional source document text for context.

    Returns:
        Dict with deck_title, template_mode=True, and slides list.
        Each slide is a flat dict with role-name keys.
    """
    print(f"[PLANNER_V2] Generating outline for topic='{topic}', tone='{tone}'")

    # Logging support for YouTube, Webpage and Document mode
    import inspect
    import os
    from datetime import datetime

    youtube_url = None
    web_url = None
    document_name = None
    frame = inspect.currentframe()
    try:
        while frame:
            locals_dict = frame.f_locals
            if "youtube_url" in locals_dict and locals_dict["youtube_url"]:
                youtube_url = locals_dict["youtube_url"]
            if "web_url" in locals_dict and locals_dict["web_url"]:
                web_url = locals_dict["web_url"]
            if "upload" in locals_dict and locals_dict["upload"]:
                upload_val = locals_dict["upload"]
                if hasattr(upload_val, "filename") and upload_val.filename:
                    document_name = upload_val.filename
            for var_name in ("pdf_file", "document_file", "doc_file", "docs_file"):
                if not document_name and var_name in locals_dict and locals_dict[var_name]:
                    val = locals_dict[var_name]
                    if hasattr(val, "filename") and val.filename:
                        document_name = val.filename
            frame = frame.f_back
    except Exception:
        pass
    finally:
        del frame

    if youtube_url and source_text:
        from services.youtube_service import extract_video_id
        try:
            video_id = extract_video_id(youtube_url)
        except Exception:
            video_id = None

        transcript_len = 0
        if video_id:
            try:
                # We can't access TRANSCRIPT_LENGTHS directly here unless imported/defined.
                # Let's import fetch_transcript or just use len(source_text) * 4 as fallback.
                from services.transcript_service import fetch_transcript
                t = fetch_transcript(video_id)
                transcript_len = len(t)
            except Exception:
                transcript_len = len(source_text) * 4

        timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        logs_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "logs")
        os.makedirs(logs_dir, exist_ok=True)

        # Append to logs/video_summary_log.txt
        txt_path = os.path.join(logs_dir, "video_summary_log.txt")
        txt_entry = (
            "====================================================\n"
            f"Timestamp:\n{timestamp_str}\n\n"
            f"Video URL:\n{youtube_url}\n\n"
            f"Topic:\n{topic}\n\n"
            f"Transcript Length:\n{transcript_len}\n\n"
            f"Summary Length:\n{len(source_text)}\n\n"
            f"Summary Preview:\n{source_text[:3000]}\n"
            "====================================================\n"
        )
        try:
            with open(txt_path, "a", encoding="utf-8") as f:
                f.write(txt_entry)
        except Exception as e:
            print(f"[PLANNER_V2] Warning: Failed to write to txt log: {e}")

        # Append to logs/video_summary_log.json
        json_path = os.path.join(logs_dir, "video_summary_log.json")
        json_entry = {
            "timestamp": timestamp_str,
            "video_url": youtube_url,
            "topic": topic,
            "transcript_length": transcript_len,
            "summary_length": len(source_text),
            "summary_preview": source_text[:3000]
        }
        try:
            with open(json_path, "a", encoding="utf-8") as f:
                f.write(json.dumps(json_entry, indent=4) + "\n")
        except Exception as e:
            print(f"[PLANNER_V2] Warning: Failed to write to json log: {e}")

    if web_url and source_text:
        timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        logs_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "logs")
        os.makedirs(logs_dir, exist_ok=True)

        # Append to logs/weblog.txt
        web_txt_path = os.path.join(logs_dir, "weblog.txt")
        web_txt_entry = (
            "====================================================\n"
            f"Timestamp:\n{timestamp_str}\n\n"
            f"Webpage URL:\n{web_url}\n\n"
            f"Topic:\n{topic}\n\n"
            f"Content Length:\n{len(source_text)}\n\n"
            f"Content Preview:\n{source_text[:3000]}\n"
            "====================================================\n"
        )
        try:
            with open(web_txt_path, "a", encoding="utf-8") as f:
                f.write(web_txt_entry)
        except Exception as e:
            print(f"[PLANNER_V2] Warning: Failed to write to weblog.txt: {e}")

        # Append to logs/weblog.json
        web_json_path = os.path.join(logs_dir, "weblog.json")
        web_json_entry = {
            "timestamp": timestamp_str,
            "web_url": web_url,
            "topic": topic,
            "content_length": len(source_text),
            "content_preview": source_text[:3000]
        }
        try:
            with open(web_json_path, "a", encoding="utf-8") as f:
                f.write(json.dumps(web_json_entry, indent=4) + "\n")
        except Exception as e:
            print(f"[PLANNER_V2] Warning: Failed to write to weblog.json: {e}")

    if document_name and source_text:
        timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        logs_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "logs")
        os.makedirs(logs_dir, exist_ok=True)

        # Append to logs/Document log.txt
        doc_txt_path = os.path.join(logs_dir, "Document log.txt")
        doc_txt_entry = (
            "====================================================\n"
            f"Timestamp:\n{timestamp_str}\n\n"
            f"Document Name:\n{document_name}\n\n"
            f"Topic:\n{topic}\n\n"
            f"Content Length:\n{len(source_text)}\n\n"
            f"Content Preview:\n{source_text[:3000]}\n"
            "====================================================\n"
        )
        try:
            with open(doc_txt_path, "a", encoding="utf-8") as f:
                f.write(doc_txt_entry)
        except Exception as e:
            print(f"[PLANNER_V2] Warning: Failed to write to Document log.txt: {e}")

        # Append to logs/Document log.json
        doc_json_path = os.path.join(logs_dir, "Document log.json")
        doc_json_entry = {
            "timestamp": timestamp_str,
            "document_name": document_name,
            "topic": topic,
            "content_length": len(source_text),
            "content_preview": source_text[:3000]
        }
        try:
            with open(doc_json_path, "a", encoding="utf-8") as f:
                f.write(json.dumps(doc_json_entry, indent=4) + "\n")
        except Exception as e:
            print(f"[PLANNER_V2] Warning: Failed to write to Document log.json: {e}")

    slide_count = template_schema.get("slide_count", 5)
    theme_info = template_schema.get("theme", {})
    slides_schema = template_schema.get("slides", [])

    # ── Build per-slide zone info for the LLM ──
    slide_data_list = []
    slide_image_roles_map = {}  # slide_index -> list of image role names

    for s_schema in slides_schema:
        slide_idx = s_schema.get("slide_index", 0)
        zones = s_schema.get("zones", [])

        filtered_zones = []
        image_roles_for_slide = []
        visual_elements = []

        for zone in zones:
            role = zone.get("role", "other")
            role_lower = role.lower()

            # Handle image/icon roles separately
            is_image = role_lower.startswith("image") or role in IMAGE_ROLES
            is_icon = is_valid_icon_role(role) or role_lower.startswith("icon")
            
            if is_image or is_icon:
                visual_elements.append({
                    "role": role,
                    "content_type": "icon" if is_icon else "image",
                    "original_name": zone.get("name", "")
                })
                image_roles_for_slide.append(role)
                continue

            # Skip anything with forbidden keywords (except custom table or chart roles)
            is_table = role_lower.startswith("table_")
            is_chart = role_lower.startswith("chart_")
            
            if not is_table and not is_chart:
                if any(k in role_lower for k in SKIP_KEYWORDS):
                    continue

            # Only allow title, subtitle, body1-N, table_*, and chart_* roles
            is_valid_content_role = (
                role in ALLOWED_CONTENT_ROLES or 
                role.startswith("body") or 
                role.startswith("title") or 
                role.startswith("subtitle") or
                is_table or
                is_chart
            )
            if not is_valid_content_role:
                continue

            # Determine word limit based on original text length to avoid overflow and overlapping
            orig_text = zone.get("original_text", "")
            orig_words = len(orig_text.split())
            max_chars = zone.get("max_chars", 200)
            
            if is_table:
                word_limit = 10
            elif is_chart:
                word_limit = 50
            else:
                if orig_words == 0:
                    word_limit = 12
                elif orig_words <= 3:
                    word_limit = orig_words + 1
                elif orig_words <= 8:
                    word_limit = orig_words + 2
                else:
                    word_limit = int(orig_words * 1.2) + 1

            # Also derive a word limit from max_chars (avg 5.5 chars per word)
            if not is_chart:
                chars_word_limit = max(3, int(max_chars / 5.5))
                word_limit = min(word_limit, chars_word_limit)

            # Ensure word_limit is within safe bounds of capacity and reasonable maximums
            capacity = zone.get("capacity", 40)
            if not is_chart:
                word_limit = min(word_limit, max(5, capacity))

            filtered_zones.append({
                "role": role,
                "word_limit": word_limit,
                "max_chars": max_chars,
                "original_text": orig_text,
                "categories": zone.get("categories") if is_chart else None,
                "series": zone.get("series") if is_chart else None,
                "row_idx": zone.get("row_idx") if is_table else None,
                "col_idx": zone.get("col_idx") if is_table else None,
            })

        slide_data_list.append({
            "slide_index": slide_idx + 1,  # 1-based for LLM
            "layout_pattern": s_schema.get("inferred_layout", "content"),
            "zones": filtered_zones,
            "visual_zones": visual_elements,
            "image_roles": image_roles_for_slide,
        })
        slide_image_roles_map[slide_idx] = image_roles_for_slide

    slides_layout_info = json.dumps(slide_data_list, indent=2)

    # ── Build LLM prompts ──
    tone_desc = TONE_DESCRIPTIONS.get(tone, TONE_DESCRIPTIONS["professional"])

    system_prompt = _build_system_prompt(tone_desc)
    user_prompt = _build_user_prompt(
        topic=topic,
        slide_count=slide_count,
        tone_desc=tone_desc,
        theme_info=theme_info,
        slides_layout_info=slides_layout_info,
        source_text=source_text,
    )

    # ── Call LLM ──
    print(f"[PLANNER_V2] Calling LLM with {slide_count} slides...")
    raw_response = call_llm(system_prompt, user_prompt, max_tokens=4000)

    # ── Parse response ──
    deck = _parse_llm_response(raw_response, topic)

    # ── Post-process: enforce allowed keys + image keyword fallbacks ──
    for slide_content in deck.get("slides", []):
        s_idx = slide_content.get("slide_index", 1) - 1  # back to 0-based

        # Remove unauthorized keys
        allowed_keys = {"slide_index", "layout"}
        for sd in slide_data_list:
            if sd["slide_index"] == s_idx + 1:
                allowed_keys.update(z["role"] for z in sd["zones"])
                allowed_keys.update(f"{ir}_keyword" for ir in sd["image_roles"])
                break

        keys_to_remove = [
            k for k in list(slide_content.keys())
            if k not in allowed_keys and not k.endswith("_keyword")
        ]
        for k in keys_to_remove:
            if k not in ("slide_index", "layout", "title", "subtitle"):
                del slide_content[k]

        # Enforce word limits AND character limits using LLM rewrites
        for sd in slide_data_list:
            if sd["slide_index"] == s_idx + 1:
                for z in sd["zones"]:
                    role = z["role"]
                    char_limit = z.get("max_chars", 2000)
                    val = slide_content.get(role, "")
                    if isinstance(val, str) and val.strip():
                        from services.template_filler_v2 import fit_text_with_llm_rewrites
                        val = fit_text_with_llm_rewrites(val, char_limit, role)
                        slide_content[role] = val
                break

        # Image/Icon keyword fallbacks
        img_roles = slide_image_roles_map.get(s_idx, [])
        for img_role in img_roles:
            kw_key = f"{img_role}_keyword"
            if not slide_content.get(kw_key):
                is_icon = is_valid_icon_role(img_role)
                if is_icon:
                    kw = _extract_icon_keyword_from_slide(slide_content, img_role)
                    if kw:
                        slide_content[kw_key] = kw
                    else:
                        default_icons = ["lightbulb", "rocket", "gear", "database", "users", "trending-up", "award"]
                        role_digit_match = re.search(r'\d+', img_role.lower())
                        idx = int(role_digit_match.group(0)) if role_digit_match else s_idx
                        slide_content[kw_key] = default_icons[idx % len(default_icons)]
                else:
                    fallback = slide_content.get("image_keyword", "")
                    if fallback:
                        slide_content[kw_key] = fallback
                    else:
                        title_text = slide_content.get("title", "") or slide_content.get("subtitle", "") or topic
                        slide_content[kw_key] = f"{title_text} professional photo".strip()

    # Force template_mode flag
    deck["template_mode"] = True
    if "deck_title" not in deck:
        deck["deck_title"] = topic

    # Pad deck["slides"] with omitted slides marked with "delete": True if not is_fallback
    is_fallback = deck.get("is_fallback", False)
    if not is_fallback:
        returned_indices = {s.get("slide_index") for s in deck.get("slides", []) if s.get("slide_index") is not None}
        for i in range(1, slide_count + 1):
            if i not in returned_indices:
                deck["slides"].append({
                    "slide_index": i,
                    "delete": True
                })
        deck["slides"] = sorted(deck["slides"], key=lambda s: s.get("slide_index", 1))

    if source_text and "slides" in deck and len(deck["slides"]) > 0:
        first_slide = deck["slides"][0]
        if topic:
            first_slide["title"] = topic.strip()
            if "subtitle" in first_slide:
                first_slide["subtitle"] = topic.strip()

    # Map Slide 1 title/subtitle to the actual analyzed shape roles
    if "slides" in deck and len(deck["slides"]) > 0:
        first_slide_content = deck["slides"][0]
        slides_schema = template_schema.get("slides", [])
        if len(slides_schema) > 0:
            first_slide_schema = slides_schema[0]
            text_zones = []
            for zone in first_slide_schema.get("zones", []):
                role = zone.get("role", "")
                role_lower = role.lower()
                is_text_role = (
                    role in ("title", "subtitle") or 
                    role_lower.startswith("body") or 
                    role_lower.startswith("title") or 
                    role_lower.startswith("subtitle")
                )
                if is_text_role:
                    text_zones.append(zone)
            
            # Sort by top_pct ascending (top-to-bottom physically)
            text_zones = sorted(text_zones, key=lambda z: z.get("top_pct", 0))
            
            # Extract generated title and subtitle
            gen_title = first_slide_content.get("title") or topic or ""
            gen_subtitle = first_slide_content.get("subtitle") or first_slide_content.get("body1") or ""
            if not gen_subtitle:
                for k in sorted(first_slide_content.keys()):
                    if k.startswith("body") and first_slide_content[k]:
                        gen_subtitle = first_slide_content[k]
                        break
            
            # Map values to their actual roles
            if len(text_zones) > 0:
                primary_role = text_zones[0].get("role", "title")
                first_slide_content[primary_role] = gen_title.strip()
                print(f"[PLANNER_V2] Slide 1 primary text zone '{primary_role}' mapped -> '{gen_title.strip()}'")
            if len(text_zones) > 1:
                secondary_role = text_zones[1].get("role", "subtitle")
                first_slide_content[secondary_role] = gen_subtitle.strip()
                print(f"[PLANNER_V2] Slide 1 secondary text zone '{secondary_role}' mapped -> '{gen_subtitle.strip()}'")
                
                # Clear subsequent text fields on Cover slide to prevent duplicating subtitle
                for tz in text_zones[2:]:
                    role = tz.get("role")
                    if role in first_slide_content:
                        first_slide_content[role] = ""

    print(f"[PLANNER_V2] Outline generation complete: {len(deck.get('slides', []))} slides")
    return deck


# ─── Private helpers ──────────────────────────────────────────────────

def _build_system_prompt(tone_desc: str) -> str:
    return f"""You are a presentation content generator.
You receive a list of slides, each with specific text zones, their word limits, max character counts, and the original template text (original_text).
Your job is to generate content for EXACTLY those zones — nothing more, nothing less.

RULES:
1. For each slide, return a flat JSON object with keys matching the zone roles.
2. Each text zone has a word_limit AND a max_chars limit. Do NOT exceed EITHER. Keep content tight and impactful.
3. For image zones (like image1, image2, image_left, hero_image), generate a key like "[role]_keyword" containing a highly specific stock photo search query (3-6 words) representing that slide's content.
4. For icon zones (like icon1, card1_icon, etc.), generate a key like "[role]_keyword" containing a simple 1-2 word noun describing a flat vector icon or symbol representing that specific card or section's concept (e.g. "rocket", "gear", "brain", "chart"). Icons are SMALL symbolic graphics, NOT photographs. Make sure each icon role on a slide gets a unique, topic-relevant icon keyword that matches its specific context (e.g. card1_icon should represent card1's content, card2_icon should represent card2's content). Never use identical or generic keywords for icons.
5. For table cell zones (role starting with 'table_'), generate a short text string (under 10 words) that represents a realistic table cell value for that row/column under the user's topic. Do not add formatting or periods unless appropriate.
6. For chart zones (role starting with 'chart_'), generate a JSON object with 'categories' (list of categories) and 'series' (list of series, each series is a dict with 'name' and 'values' list of numbers) representing the updated chart data for the topic. Keep the count of categories and series same as in the original template.
7. Do NOT add keys that are not in the zone list.
8. Use the tone: {tone_desc}
9. Content must be specific and relevant to the topic. No generic placeholders.
10. Title should be punchy and concise (maximum 8 words).
11. Subtitle should expand on the title (maximum 12 words).
12. Body text (body1, body2, etc.) must match the word limit specified for each zone. Generate concise, card-friendly text.
13. No paragraph should exceed two lines. Keep every zone short and scannable.
14. CRITICAL: Look at the "original_text" provided for each text zone. You MUST rewrite the content to match the new topic, but you must match the exact style, purpose, semantic context, and length of the original_text. For example, if the original_text is a step (e.g., 'Step 1: Planning'), rewrite it as a step for the new topic (e.g., 'Step 1: Design'). If the original_text is a short label (e.g., 'Introduction'), rewrite it as a short label (e.g., 'Overview'). If the original_text contains tech names or numbers, replace them with appropriate values for the new topic. Do NOT add extra paragraphs or content.
15. CRITICAL: ALL slides in the deck MUST be about the SAME topic. Never mix different projects, subjects, or unrelated content across slides. If the topic is "festivals", every slide discusses festivals. If the topic is "AI Study Planner", every slide discusses that single project. Never introduce unrelated subjects like random algorithms, diseases, or other projects.
16. CRITICAL: Respect the max_chars limit strictly. Count your characters including spaces. If your text exceeds max_chars, shorten it.

Return ONLY valid JSON. No markdown fences, no commentary."""


def _build_user_prompt(
    topic: str,
    slide_count: int,
    tone_desc: str,
    theme_info: dict,
    slides_layout_info: str,
    source_text: str,
) -> str:
    prompt = f"""Generate content for a {slide_count}-slide presentation about: "{topic}"

Tone: {tone_desc}
Theme fonts: heading={theme_info.get('heading_font', 'Calibri')}, body={theme_info.get('body_font', 'Calibri')}

Here are the slides with their zones, word limits, and original template text:
{slides_layout_info}

Return a JSON object with this structure:
{{
  "deck_title": "Short deck title",
  "slides": [
    {{
      "slide_index": 1,
      "title": "Slide Title Here",
      "subtitle": "Expanding subtitle text",
      "body1": "First body paragraph with details...",
      "image1_keyword": "specific photo search query",
      "icon1_keyword": "rocket"
    }}
  ]
}}

Each slide's keys must EXACTLY match the zone roles listed above.
For image and icon roles, generate "[role]_keyword" keys with search queries.
CRITICAL: You MUST generate a distinct, topic-relevant search keyword for every single icon and image role listed under "image_roles" for each slide. Ensure icon keywords are simple 1-2 word nouns (e.g., "gear", "rocket") and represent the specific sub-topic or section they reside in."""

    if source_text:
        # Truncate source text to avoid token overflow
        max_source = 2000
        if len(source_text) > max_source:
            source_text = source_text[:max_source] + "..."
        prompt += f"\n\nSource document context:\n{source_text}\n"
        prompt += f"CRITICAL: You MUST use the Source document context to extract information and details specifically as they relate to the topic '{topic}'. Focus the presentation content entirely on '{topic}', using the source document as the primary source of truth."

    return prompt


def _parse_llm_response(raw: str, topic: str) -> dict:
    """Parse LLM JSON response with fallback handling."""
    # Strip markdown code fences
    cleaned = raw.strip()
    if cleaned.startswith("```"):
        # Remove opening fence
        first_newline = cleaned.index("\n")
        cleaned = cleaned[first_newline + 1:]
    if cleaned.endswith("```"):
        cleaned = cleaned[:-3]
    cleaned = cleaned.strip()

    is_fallback = False
    try:
        deck = json.loads(cleaned)
    except json.JSONDecodeError as e:
        print(f"[PLANNER_V2] JSON parse error: {e}")
        is_fallback = True
        # Try to extract JSON from response
        json_match = re.search(r'\{[\s\S]*\}', cleaned)
        if json_match:
            try:
                deck = json.loads(json_match.group())
                is_fallback = False
            except json.JSONDecodeError:
                print("[PLANNER_V2] Fallback JSON parse also failed. Returning minimal deck.")
                deck = {
                    "deck_title": topic,
                    "slides": [{"slide_index": 1, "title": topic}]
                }
        else:
            deck = {
                "deck_title": topic,
                "slides": [{"slide_index": 1, "title": topic}]
            }

    # Ensure slides is a list
    if "slides" not in deck or not isinstance(deck["slides"], list):
        deck["slides"] = [{"slide_index": 1, "title": topic}]
        is_fallback = True

    deck["is_fallback"] = is_fallback
    return deck
