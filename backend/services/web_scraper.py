import requests
from bs4 import BeautifulSoup

def extract_text_from_url(url: str) -> str:
    """
    Fetches the webpage content from the given URL and parses it using BeautifulSoup.
    It removes HTML boilerplate (scripts, styles, headers, footers, etc.)
    and returns cleaned, readable plain text.
    """
    url_stripped = url.strip()
    if not url_stripped.startswith("http://") and not url_stripped.startswith("https://"):
        url_stripped = "https://" + url_stripped
        
    print(f"[BS4] Using legacy BeautifulSoup scraper for URL: {url_stripped}")

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    
    try:
        print(f"[BS4_FALLBACK] GET request to: {url_stripped}")
        response = requests.get(url_stripped, headers=headers, timeout=12)
        response.raise_for_status()
    except Exception as e:
        print(f"[BS4_FAILURE] HTTP request failed: {str(e)}")
        raise ValueError(f"HTTP request failed: {str(e)}")
        
    try:
        soup = BeautifulSoup(response.content, "html.parser")
        
        # Remove non-content tags
        for element in soup(["script", "style", "meta", "noscript", "header", "footer", "nav", "aside", "svg", "iframe"]):
            element.decompose()
            
        # Get text
        text = soup.get_text()
        
        # Format text to clean lines and remove blank sections
        lines = (line.strip() for line in text.splitlines())
        chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
        cleaned_text = '\n'.join(chunk for chunk in chunks if chunk)
        
        print(f"[BS4_SUCCESS] Extracted {len(cleaned_text)} chars")
        return cleaned_text
    except Exception as e:
        print(f"[BS4_FAILURE] HTML parsing failed: {str(e)}")
        raise ValueError(f"HTML parsing failed: {str(e)}")
