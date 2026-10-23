import unittest
from unittest.mock import patch, MagicMock
import re

from services.slide_planner_v2 import (
    is_valid_icon_role,
    _extract_icon_keyword_from_slide,
    generate_outline_from_template_v2
)
from services.template_filler_v2 import fill_template_v2

class TestTemplatePlannerAndFiller(unittest.TestCase):
    def test_is_valid_icon_role(self):
        # Whitelisted icon roles
        self.assertTrue(is_valid_icon_role("icon1"))
        self.assertTrue(is_valid_icon_role("icon10"))
        self.assertTrue(is_valid_icon_role("card1_icon"))
        self.assertTrue(is_valid_icon_role("custom_icon"))
        self.assertTrue(is_valid_icon_role("icon_left_icon"))

        # Non-whitelisted/decorative icon roles
        self.assertFalse(is_valid_icon_role("icon11"))
        self.assertFalse(is_valid_icon_role("icon786"))
        self.assertFalse(is_valid_icon_role("decorative_icon"))
        self.assertFalse(is_valid_icon_role("shape_icon_xyz")) # does not end with _icon or have _icon_

    def test_extract_icon_keyword_from_slide(self):
        # Test case 1: Matching prefix (e.g. card1_icon matching card1_title/description)
        slide_content_1 = {
            "slide_index": 2,
            "card1_title": "Innovative Hardware Architecture",
            "card1_description": "We build highly robust and modular microcontrollers.",
            "card2_title": "Software Stack",
        }
        kw_1 = _extract_icon_keyword_from_slide(slide_content_1, "card1_icon")
        # Should extract words from card1_title or card1_description
        self.assertIn("innovative", kw_1.lower() or "hardware" in kw_1.lower())

        # Test case 2: Matching digit (e.g. icon2 matching body2/title2)
        slide_content_2 = {
            "slide_index": 3,
            "title2": "Cloud Integrations",
            "body2": "Seamless scaling to AWS and GCP servers.",
            "body1": "Local network settings",
        }
        kw_2 = _extract_icon_keyword_from_slide(slide_content_2, "icon2")
        self.assertTrue(any(w in kw_2.lower() for w in ["cloud", "integrations", "scaling", "aws", "gcp"]))

        # Test case 3: General fallback from slide title if no direct match
        slide_content_3 = {
            "slide_index": 4,
            "title": "Sustainable Agriculture",
            "body": "Farming with smart IoT devices.",
        }
        kw_3 = _extract_icon_keyword_from_slide(slide_content_3, "icon_any")
        self.assertTrue(any(w in kw_3.lower() for w in ["sustainable", "agriculture", "smart", "iot"]))

    @patch('services.slide_planner_v2.call_llm')
    def test_generate_outline_padding_and_fallbacks(self, mock_call_llm):
        # Mock LLM returning content for only 1 slide out of 3
        mock_response = """
        {
          "deck_title": "Arduino Smart Home",
          "slides": [
            {
              "slide_index": 1,
              "title": "Welcome to Smart Home",
              "subtitle": "Overview of features",
              "body1": "Control lights and climate automatically.",
              "image1_keyword": "smart home dashboard"
            }
          ]
        }
        """
        mock_call_llm.return_value = mock_response

        # Minimal template schema with 3 slides
        template_schema = {
            "slide_count": 3,
            "theme": {},
            "slides": [
                {
                    "slide_index": 0,
                    "zones": [
                        {"role": "title", "original_text": "Slide 1 Title"},
                        {"role": "subtitle", "original_text": "Slide 1 Sub"},
                        {"role": "image1", "original_text": ""}
                    ]
                },
                {
                    "slide_index": 1,
                    "zones": [
                        {"role": "title", "original_text": "Slide 2 Title"},
                        {"role": "body1", "original_text": "Slide 2 Body"},
                        {"role": "card1_icon", "original_text": ""}
                    ]
                },
                {
                    "slide_index": 2,
                    "zones": [
                        {"role": "title", "original_text": "Slide 3 Title"}
                    ]
                }
            ]
        }

        # Run planner with source_text context to trigger title override
        deck = generate_outline_from_template_v2("Arduino Smart Home", template_schema, "technical", source_text="dummy context")

        # Verify that max_tokens was passed as 4000
        mock_call_llm.assert_called_once()
        self.assertEqual(mock_call_llm.call_args[1].get("max_tokens"), 4000)

        # Verify slide list size is padded to 3 slides
        slides = deck["slides"]
        self.assertEqual(len(slides), 3)

        # Slide 1 should contain LLM generated details
        self.assertEqual(slides[0]["slide_index"], 1)
        self.assertEqual(slides[0]["title"], "Arduino Smart Home") # overridden for first slide
        self.assertEqual(slides[0]["image1_keyword"], "smart home dashboard")

        # Slides 2 and 3 should be padded with delete: True
        self.assertEqual(slides[1]["slide_index"], 2)
        self.assertTrue(slides[1]["delete"])
        self.assertEqual(slides[2]["slide_index"], 3)
        self.assertTrue(slides[2]["delete"])

    @patch('services.template_filler_v2._fill_slide_v2')
    @patch('services.template_filler_v2._delete_slide')
    @patch('services.template_filler_v2.Presentation')
    def test_safe_slide_deletion_logic(self, mock_presentation, mock_delete_slide, mock_fill_slide):
        # Mock Presentation structure
        mock_prs = MagicMock()
        mock_slide_1 = MagicMock()
        mock_slide_2 = MagicMock()
        mock_slide_3 = MagicMock()
        mock_prs.slides = [mock_slide_1, mock_slide_2, mock_slide_3]
        mock_presentation.return_value = mock_prs

        # Slide contents: only slide 3 is marked for deletion, slide 2 is unmodified, slide 1 is filled
        slide_contents = [
            {"slide_index": 1, "title": "Real Content"},
            {"slide_index": 2}, # empty/no-delete
            {"slide_index": 3, "delete": True}
        ]

        # Call filler with correct parameter names
        fill_template_v2(
            template_path="dummy_template.pptx",
            template_schema={"slides": [{}, {}, {}]},
            slide_contents=slide_contents,
            output_dir="dummy_out.pptx"
        )

        # Should only call _delete_slide for the 3rd slide (index 2)
        mock_delete_slide.assert_called_once_with(mock_prs, 2)

if __name__ == "__main__":
    unittest.main()
