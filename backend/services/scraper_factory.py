from services.web_scraper import extract_text_from_url as bs_extract_text_from_url
from services.scrapling_scraper import ScraplingScraper

# Step 9: Regression Protection Feature Flag
USE_SCRAPLING = True

print(f"[SCRAPER_FACTORY]\nUSE_SCRAPLING={USE_SCRAPLING}")

import time

class BeautifulSoupScraperWrapper:
    """
    Wrapper around the legacy BeautifulSoup scraper to provide a unified
    interface for the factory.
    """
    def extract_text_from_url(self, url: str, structured: bool = False):
        url_stripped = url.strip()
        start_total = time.time()
        source = "beautifulsoup"
        
        plain_text = bs_extract_text_from_url(url_stripped)
        
        total_time = time.time() - start_total
        
        # Step 7: Final Verification Summary
        print("=========================")
        print("SCRAPER SUMMARY")
        print("===============")
        print(f"URL: {url_stripped}")
        print(f"SOURCE: {source}")
        print(f"CHARS: {len(plain_text)}")
        print("CACHE: N/A")
        print(f"TOTAL TIME: {total_time:.2f}s")
        print("=================")

        if structured:
            return {
                "title": "Webpage Summary",
                "headings": [],
                "content": plain_text,
                "images": [],
                "source": source
            }
        return plain_text

def get_scraper():
    """
    Returns the active scraper implementation based on the feature flag.
    """
    if USE_SCRAPLING:
        return ScraplingScraper()
    else:
        return BeautifulSoupScraperWrapper()
