import unittest
from unittest.mock import patch, MagicMock
from services.pptx_builder import build_pptx

class TestPptxTone(unittest.TestCase):
    @patch('services.pptx_builder._add_1_column_slide')
    @patch('services.professional_layouts._add_prof_1_column_slide')
    def test_professional_tone_calls_professional_layout(self, mock_prof_layout, mock_std_layout):
        deck = {
            "deck_title": "Professional Test",
            "slides": [
                {
                    "title": "Intro",
                    "layout_id": "1-column",
                    "slide_type": "content",
                    "bullet_points": ["Point 1"]
                }
            ]
        }
        
        # Call build_pptx with Professional tone
        build_pptx(deck, topic="Professional_Test", tone="Professional")
        
        # Verify the professional layout builder was called
        mock_prof_layout.assert_called_once()
        mock_std_layout.assert_not_called()

    @patch('services.pptx_builder._add_1_column_slide')
    @patch('services.professional_layouts._add_prof_1_column_slide')
    def test_non_professional_tone_calls_standard_layout(self, mock_prof_layout, mock_std_layout):
        deck = {
            "deck_title": "Technical Test",
            "slides": [
                {
                    "title": "Intro",
                    "layout_id": "1-column",
                    "slide_type": "content",
                    "bullet_points": ["Point 1"]
                }
            ]
        }
        
        # Call build_pptx with Technical tone
        build_pptx(deck, topic="Technical_Test", tone="Technical")
        
        # Verify the standard layout builder was called, NOT the professional one
        mock_std_layout.assert_called_once()
        mock_prof_layout.assert_not_called()

    @patch('api.routes.build_pptx')
    def test_export_pptx_route_preserves_tone(self, mock_build_pptx):
        from fastapi import FastAPI
        from fastapi.testclient import TestClient
        from api.routes import router
        
        # Setup test client for routers
        app = FastAPI()
        app.include_router(router)
        client = TestClient(app)
        
        # Mock build_pptx return path to exist
        import os
        import tempfile
        temp_pptx = os.path.join(tempfile.gettempdir(), "test_export.pptx")
        with open(temp_pptx, "wb") as f:
            f.write(b"dummy pptx data")
            
        mock_build_pptx.return_value = temp_pptx
        
        # Call /export-pptx with tone "Neumorphism"
        response = client.post("/export-pptx", json={
            "title": "Sports in Neumorphism",
            "tone": "Neumorphism",
            "slides": [
                {
                    "title": "Soccer Overview",
                    "layout_id": "OneColumn",
                    "zone_content": {
                        "body": "Soccer is played globally."
                    }
                }
            ]
        })
        
        self.assertEqual(response.status_code, 200)
        # Verify build_pptx was called with tone="Neumorphism"
        mock_build_pptx.assert_called_once()
        self.assertEqual(mock_build_pptx.call_args[1].get("tone"), "Neumorphism")

    def test_ecommerce_and_education_tone_resolutions(self):
        from services.pptx_builder import get_pptx_theme
        from services.theme_engine import get_theme

        # Test get_pptx_theme for ecommerce variants
        theme1 = get_pptx_theme("e-commerce")
        theme2 = get_pptx_theme("ecommerce")
        theme3 = get_pptx_theme("Ecommerce")

        # They should all resolve to the same ecommerce theme, with bg_color and primary_color matching
        self.assertEqual(theme1["bg_color"], theme2["bg_color"])
        self.assertEqual(theme2["bg_color"], theme3["bg_color"])
        self.assertIsNotNone(theme1["bg_color"])
        
        # Test education theme resolution
        theme_edu = get_pptx_theme("education")
        self.assertIsNotNone(theme_edu["bg_color"])

        # Test get_theme from theme_engine
        engine_theme1 = get_theme("e-commerce")
        engine_theme2 = get_theme("ecommerce")
        self.assertEqual(engine_theme1["primary"], engine_theme2["primary"])
        self.assertEqual(engine_theme1["background"], "#0F0A28")

    @patch('services.professional_layouts._add_prof_hero_slide')
    def test_layout_casing_normalization(self, mock_add_prof_hero_slide):
        from pptx import Presentation
        prs = Presentation()
        palette = {
            "bg": None,
            "text": None,
            "text_muted": None,
            "primary": None,
            "secondary": None,
            "surface": None,
        }
        slide_data = {"layout_id": "Hero", "title": "Test Title"}
        from services.pptx_builder import _build_slide_from_scratch
        _build_slide_from_scratch(prs, slide_data, "title", palette, tone="professional")
        
        mock_add_prof_hero_slide.assert_called_once()

if __name__ == "__main__":
    unittest.main()
