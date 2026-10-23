import time
import os
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from services.web_scraper import extract_text_from_url as bs_extract_text_from_url

def validate_dependencies():
    patchright_installed = False
    playwright_installed = False
    chromium_installed = False

    try:
        import patchright
        patchright_installed = True
    except ImportError:
        pass

    try:
        import playwright
        playwright_installed = True
    except ImportError:
        pass

    if patchright_installed and playwright_installed:
        try:
            import glob
            local_appdata = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~/AppData/Local")
            playwright_dir = os.path.join(local_appdata, "ms-playwright")
            if os.path.exists(playwright_dir):
                chromium_dirs = glob.glob(os.path.join(playwright_dir, "chromium-*"))
                for d in chromium_dirs:
                    exe_path = os.path.join(d, "chrome-win64", "chrome.exe")
                    if os.path.exists(exe_path):
                        chromium_installed = True
                        break
        except Exception:
            pass

    print("[SCRAPLING_VALIDATION]")
    print(f"Patchright Installed = {patchright_installed}")
    print(f"Playwright Installed = {playwright_installed}")
    print(f"Chromium Installed = {chromium_installed}")

# Run startup validation
validate_dependencies()

# In-memory cache to store scraped results
# Format: { url: { "title": str, "headings": list, "subheadings": list, ... } }
SCRAPER_CACHE = {}

# Step 6: Add feature flag
USE_DYNAMIC_FETCHER = True

class ScraplingScraper:
    def _is_suspicious_content(self, html: str, text: str) -> bool:
        """
        Detects if the page appears to be React/Next.js/Vue rendered but lacks
        actual text content in the raw page source, or if it is suspiciously small.
        """
        if not text or len(text.strip()) < 300:
            return True
            
        html_lower = html.lower()
        text_lower = text.lower()

        # Check for empty SPA mounting points in HTML
        spa_mounts = [
            '<div id="root"></div>', 
            '<div id="app"></div>', 
            '<div id="__next"></div>', 
            '<div id="root"/>', 
            '<div id="app"/>',
            '<div id="__next"/>'
        ]
        if any(mount in html for mount in spa_mounts):
            return True

        # Check for React/Next.js hydration blocks with generic loading text
        if "__NEXT_DATA__" in html or "window.__PRELOADED_STATE__" in html:
            if len(text.strip()) < 1000 and any(kw in text_lower for kw in ["loading", "please wait", "javascript"]):
                return True

        # Check for generic loading page/missing JS indicators
        loading_indicators = [
            "loading...", 
            "please enable javascript", 
            "you need to enable javascript",
            "javascript is required",
            "enable javascript to run this app"
        ]
        if any(indicator in text_lower for indicator in loading_indicators):
            return True

        return False

    def _try_stealthy_fetch(self, url: str):
        """
        Executes Scrapling's StealthyFetcher in a separate thread.
        """
        print(f"[STEALTHY_FETCHER] Attempting fetch")
        start_fetch = time.time()
        try:
            from scrapling.fetchers import StealthyFetcher
            from concurrent.futures import ThreadPoolExecutor

            def fetch_stealthy():
                import asyncio
                import sys
                if sys.platform == 'win32':
                    try:
                        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
                    except Exception as e:
                        print(f"Failed to set WindowsProactorEventLoopPolicy in stealthy thread: {e}")
                return StealthyFetcher.fetch(url, headless=True, timeout=15000)

            with ThreadPoolExecutor(max_workers=1) as executor:
                future = executor.submit(fetch_stealthy)
                response = future.result()

            fetch_time = time.time() - start_fetch
            if response and response.status == 200:
                print(f"[STEALTHY_FETCHER] Success")
                print(f"[STEALTHY_FETCHER_SUCCESS]")
                return response.html_content, fetch_time
            else:
                status = response.status if response else "None"
                print(f"[STEALTHY_FETCHER] Failed (status={status})")
                print(f"[STEALTHY_FETCHER_FAILURE]")
                return None, fetch_time
        except Exception as e:
            print(f"[STEALTHY_FETCHER] Failed | Exception: {str(e)}")
            print(f"[STEALTHY_FETCHER_FAILURE]")
            return None, time.time() - start_fetch

    def _try_dynamic_fetch(self, url: str):
        """
        Executes Scrapling's DynamicFetcher in a separate thread to render JS.
        """
        # Step 3: Add logs
        print(f"[DYNAMIC_FETCHER] Attempting JS-rendered extraction")
        start_fetch = time.time()
        try:
            from scrapling.fetchers import DynamicFetcher
            from concurrent.futures import ThreadPoolExecutor

            def fetch_dynamic():
                import asyncio
                import sys
                if sys.platform == 'win32':
                    try:
                        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
                    except Exception as e:
                        print(f"Failed to set WindowsProactorEventLoopPolicy in dynamic thread: {e}")
                return DynamicFetcher.fetch(url, headless=True, timeout=15000)

            with ThreadPoolExecutor(max_workers=1) as executor:
                future = executor.submit(fetch_dynamic)
                response = future.result()

            fetch_time = time.time() - start_fetch
            if response and response.status == 200:
                print(f"[DYNAMIC_FETCHER_SUCCESS] Extracted {len(response.html_content)} chars")
                print(f"[DYNAMIC_FETCHER_SUCCESS]")
                return response.html_content, fetch_time
            else:
                status = response.status if response else "None"
                print(f"[DYNAMIC_FETCHER_FAILURE] Failed (status={status})")
                print(f"[DYNAMIC_FETCHER_FAILURE]")
                return None, fetch_time
        except Exception as e:
            print(f"[DYNAMIC_FETCHER_FAILURE] Failed | Exception: {str(e)}")
            print(f"[DYNAMIC_FETCHER_FAILURE]")
            return None, time.time() - start_fetch

    def _try_standard_fetch(self, url: str):
        """
        Executes standard Scrapling Fetcher.
        """
        print(f"[FETCHER] Attempting fallback")
        start_fetch = time.time()
        try:
            from scrapling.fetchers import Fetcher
            response = Fetcher.get(url, timeout=12)
            fetch_time = time.time() - start_fetch
            if response and response.status == 200:
                print(f"[FETCHER] Success")
                return response.html_content, fetch_time
            else:
                status = response.status if response else "None"
                print(f"[FETCHER] Failed (status={status})")
                return None, fetch_time
        except Exception as e:
            print(f"[FETCHER] Failed | Exception: {str(e)}")
            return None, time.time() - start_fetch

    def fetch_url_html(self, url_stripped: str):
        """
        Coordinates the tiered web scraping fetch pipeline:
        StealthyFetcher -> DynamicFetcher (fallback/JS rendering) -> Fetcher -> BeautifulSoup
        """
        html_content = None
        fetcher_used = None
        fetch_time = 0.0
        source = None

        # Tier 1: StealthyFetcher
        html_content, fetch_time = self._try_stealthy_fetch(url_stripped)
        if html_content:
            fetcher_used = "StealthyFetcher"
            source = "scrapling_stealthy"
            
            # Check if text is suspiciously small or SPA rendered
            soup = BeautifulSoup(html_content, "html.parser")
            text = soup.get_text()
            if self._is_suspicious_content(html_content, text):
                print(f"[SCRAPLING] StealthyFetcher succeeded but content is sparse/JS-dependent. Attempting DynamicFetcher...")
                html_content = None # Trigger next fallback tier

        # Tier 2: DynamicFetcher (if Stealthy failed or content is suspicious)
        if not html_content and USE_DYNAMIC_FETCHER:
            html_content, fetch_time = self._try_dynamic_fetch(url_stripped)
            if html_content:
                fetcher_used = "DynamicFetcher"
                source = "scrapling_dynamic"

        # Tier 3: Standard Fetcher
        if not html_content:
            html_content, fetch_time = self._try_standard_fetch(url_stripped)
            if html_content:
                fetcher_used = "Fetcher"
                source = "scrapling_fetcher"

        # Tier 4: BeautifulSoup HTTP Fallback
        if not html_content:
            try:
                print(f"[BS4_FALLBACK] Attempting fallback to BeautifulSoup")
                start_fetch = time.time()
                html_content = bs_extract_text_from_url(url_stripped)
                fetch_time = time.time() - start_fetch
                fetcher_used = "BeautifulSoup"
                source = "beautifulsoup"
                print(f"[BS4_FALLBACK] Success")
            except Exception as e:
                print(f"[SCRAPLING_FAILURE] All scraper attempts failed")
                raise ValueError(f"All scraper attempts failed for URL: {url_stripped}. Error: {str(e)}")

        return html_content, fetcher_used, fetch_time, source

    def detect_page_type(self, url: str, title: str, headings: list, content: str) -> str:
        """
        Determines the page type category based on keyword density in URL, title,
        headings, and page content.
        """
        url_lower = url.lower()
        title_lower = title.lower()
        content_lower = content.lower()
        headings_lower = [h.lower() for h in headings]
        
        # 1. Pricing Page
        pricing_score = 0
        if "pricing" in url_lower or "price" in url_lower or "plan" in url_lower:
            pricing_score += 5
        if "pricing" in title_lower or "plan" in title_lower:
            pricing_score += 3
        pricing_score += content_lower.count("$") + content_lower.count("usd") + content_lower.count("subscribe")
        if any("pricing" in h or "plan" in h for h in headings_lower):
            pricing_score += 3
            
        # 2. Documentation
        docs_score = 0
        if "docs" in url_lower or "documentation" in url_lower or "api" in url_lower:
            docs_score += 5
        if "docs" in title_lower or "documentation" in title_lower or "guide" in title_lower:
            docs_score += 3
        if any(k in content_lower for k in ["api reference", "installation", "quick start", "getting started", "code example"]):
            docs_score += 4
            
        # 3. Blog
        blog_score = 0
        if "blog" in url_lower or "post" in url_lower or "article" in url_lower:
            blog_score += 5
        if "blog" in title_lower or "news" in title_lower or "published" in title_lower:
            blog_score += 3
        if any(k in content_lower for k in ["published on", "written by", "read time", "comments"]):
            blog_score += 4

        # 4. Course Page
        course_score = 0
        if "course" in url_lower or "syllabus" in url_lower or "curriculum" in url_lower:
            course_score += 5
        if "course" in title_lower or "learn" in title_lower or "class" in title_lower:
            course_score += 3
        if any(k in content_lower for k in ["syllabus", "curriculum", "lectures", "instructor", "enroll now"]):
            course_score += 4

        # 5. Portfolio
        portfolio_score = 0
        if "portfolio" in url_lower or "projects" in url_lower or "gallery" in url_lower:
            portfolio_score += 5
        if "portfolio" in title_lower or "my work" in title_lower:
            portfolio_score += 3
        if any(k in content_lower for k in ["my projects", "selected works", "resume", "skills", "about me"]):
            portfolio_score += 4

        # 6. Product Page
        product_score = 0
        if "product" in url_lower or "features" in url_lower or "item" in url_lower or "store" in url_lower:
            product_score += 3
        if "features" in title_lower or "buy" in title_lower:
            product_score += 2
        if any(k in content_lower for k in ["add to cart", "add to bag", "specifications", "features", "reviews"]):
            product_score += 4
            
        scores = {
            "pricing page": pricing_score,
            "documentation": docs_score,
            "blog": blog_score,
            "course page": course_score,
            "portfolio": portfolio_score,
            "product page": product_score
        }
        
        max_type = max(scores, key=scores.get)
        if scores[max_type] > 2:
            return max_type
        return "company website"

    def create_sections(self, title: str, headings: list, paragraphs: list, products: list, 
                        testimonials: list, team_members: list, faqs: list, contact_info: dict) -> list:
        """
        Organizes the extracted content fields into logical, structured slide-friendly sections.
        """
        sections = []
        
        # 1. Hero Section
        hero_content = [f"Title: {title}"]
        if paragraphs:
            hero_content.append(paragraphs[0])
            if len(paragraphs) > 1:
                hero_content.append(paragraphs[1])
        sections.append({
            "section_name": "Introduction",
            "content": hero_content
        })
        
        # 2. Main Content Sections
        # Map remaining paragraphs under extracted headings
        p_index = 2
        for h in headings[:4]:
            sec_content = []
            # Gather next few paragraphs
            while p_index < len(paragraphs) and len(sec_content) < 2:
                sec_content.append(paragraphs[p_index])
                p_index += 1
            if sec_content:
                sections.append({
                    "section_name": h,
                    "content": sec_content
                })
                
        # 3. Offerings / Products / Pricing
        if products:
            sections.append({
                "section_name": "Core Offerings & Pricing",
                "content": [f"{p['name']} ({p['price']}): {p['description']}" if p['price'] else f"{p['name']}: {p['description']}" for p in products]
            })
            
        # 4. Social Proof / Testimonials
        if testimonials:
            sections.append({
                "section_name": "What Our Clients Say",
                "content": testimonials[:3]
            })

        # 5. Team Section
        if team_members:
            sections.append({
                "section_name": "Our Team",
                "content": [f"{m['name']} - {m['role']}" if m['role'] else m['name'] for m in team_members]
            })

        # 6. FAQ Section
        if faqs:
            sections.append({
                "section_name": "Frequently Asked Questions",
                "content": [f"Q: {f['question']}\nA: {f['answer']}" for f in faqs[:3]]
            })

        # 7. Contact Details
        if any(contact_info.values()):
            contact_lines = []
            if contact_info.get("phone"):
                contact_lines.append(f"Phone: {contact_info['phone']}")
            if contact_info.get("email"):
                contact_lines.append(f"Email: {contact_info['email']}")
            if contact_info.get("address"):
                contact_lines.append(f"Address: {contact_info['address']}")
            if contact_info.get("social_links"):
                contact_lines.append(f"Social Channels: {', '.join(contact_info['social_links'])}")
            if contact_lines:
                sections.append({
                    "section_name": "Contact & Support",
                    "content": contact_lines
                })
                
        return sections

    def extract_structured_content(self, url: str) -> dict:
        """
        DOM-aware structured content extractor. Extracts detailed segments optimized
        for downstream AI slide planning.
        """
        url_stripped = url.strip()
        if not url_stripped.startswith("http://") and not url_stripped.startswith("https://"):
            url_stripped = "https://" + url_stripped

        # Cache Check
        if url_stripped in SCRAPER_CACHE:
            print(f"[CACHE_HIT] URL={url_stripped}")
            print(f"[CACHE] Current cache size: {len(SCRAPER_CACHE)}")
            
            cached_res = SCRAPER_CACHE[url_stripped]
            
            print("=========================")
            print("SCRAPER SUMMARY")
            print("===============")
            print(f"URL: {url_stripped}")
            print(f"SOURCE: {cached_res.get('source', 'unknown')}")
            print(f"CHARS: {len(cached_res.get('content', ''))}")
            print("CACHE: HIT")
            print("TOTAL TIME: 0.00s")
            print("=================")
            return cached_res

        print(f"[CACHE_MISS] URL={url_stripped}")
        print(f"[CACHE] Current cache size: {len(SCRAPER_CACHE)}")

        # Step 9: Structured extraction logs
        print(f"[STRUCTURED_EXTRACTION] Parsing structured content for URL: {url_stripped}")
        start_total = time.time()

        # Fetch HTML
        html_content, fetcher_used, fetch_time, source = self.fetch_url_html(url_stripped)

        # Handle BeautifulSoup fallback when HTML content is plain cleaned text instead of DOM
        if source == "beautifulsoup" and not html_content.startswith("<"):
            plain_text = html_content
            parse_time = 0.0
            total_time = time.time() - start_total
            
            structured_res = {
                "title": "Webpage Summary",
                "headings": [],
                "subheadings": [],
                "paragraphs": [p.strip() for p in plain_text.split("\n") if p.strip()],
                "images": [],
                "tables": [],
                "lists": [],
                "testimonials": [],
                "faqs": [],
                "contact_info": {"phone": "", "email": "", "address": "", "social_links": []},
                "products": [],
                "team_members": [],
                "page_type": "company website",
                "content": plain_text,
                "sections": [{"section_name": "Core Summary", "content": [plain_text[:500]]}],
                "source": source
            }
            SCRAPER_CACHE[url_stripped] = structured_res
            print(f"[SCRAPLING_FAILURE] BeautifulSoup returned text content. Summary logged.")
            return structured_res

        # DOM-aware parsing
        start_parse = time.time()
        try:
            from scrapling import Selector
            page = Selector(html_content, url=url_stripped, adaptive=True)
            print("[ADAPTIVE_SELECTORS]")
            print("Enabled = True")
            print("[ADAPTIVE_SELECTORS_ENABLED]")

            # Title
            title = page.css("title::text", adaptive=True).get() or "Webpage Summary"
            title = title.strip()
            print(f"[TITLE] Extracted title: {title}")

            # Headings & Subheadings
            headings = [h.strip() for h in page.css("h1::text, h2::text", adaptive=True).getall() if h.strip()]
            subheadings = [h.strip() for h in page.css("h3::text, h4::text", adaptive=True).getall() if h.strip()]
            print(f"[HEADINGS] Extracted {len(headings)} headings")

            # Paragraphs
            paragraphs = [p.strip() for p in page.css("p::text", adaptive=True).getall() if len(p.strip()) > 30]

            # Images
            images = []
            for img in page.css("img", adaptive=True):
                src = img.css("::attr(src)").get() or img.css("::attr(data-src)").get()
                if src:
                    abs_src = urljoin(url_stripped, src.strip())
                    if abs_src.startswith("http") and abs_src not in images:
                        src_lower = abs_src.lower()
                        if src_lower.endswith(".svg") or ".svg?" in src_lower:
                            continue
                        if any(term in src_lower for term in ["logo", "spacer", "favicon", "sprite", "tracker", "ad_"]):
                            continue
                        images.append(abs_src)
            print(f"[IMAGES] Extracted {len(images)} images")

            # Tables
            tables = []
            for table in page.css("table", adaptive=True):
                rows = []
                for tr in table.css("tr"):
                    cells = [td.get_all_text().strip() for td in tr.css("td, th")]
                    if cells:
                        rows.append(cells)
                if rows:
                    tables.append(rows)
            print(f"[TABLES] Extracted {len(tables)} tables")

            # Lists
            lists = []
            for lst in page.css("ul, ol", adaptive=True):
                items = [li.get_all_text().strip() for li in lst.css("li")]
                items = [i for i in items if i]
                if items and len(items) > 1:
                    lists.append(items)

            # Testimonials
            testimonials = []
            for element in page.css("blockquote, .testimonial, .review, [id*='testimonial'], [id*='review']", adaptive=True):
                val = element.get_all_text().strip()
                if val and len(val) > 20:
                    testimonials.append(val)
            print(f"[TESTIMONIALS] Extracted {len(testimonials)} testimonials")

            # FAQs
            faqs = []
            faq_containers = page.css(".faq, .faqs, [id*='faq']", adaptive=True)
            if faq_containers:
                for faq in faq_containers:
                    q_elements = faq.css(".question, .q, h3, h4, dt")
                    for q in q_elements:
                        question = q.get_all_text().strip()
                        if question:
                            answer = ""
                            ans_sel = q.xpath("./following-sibling::*[1]")
                            if ans_sel and len(ans_sel) > 0:
                                answer = ans_sel[0].get_all_text().strip()
                            if not answer:
                                ans_sel = faq.css(".answer, .a, dd")
                                if ans_sel:
                                    answer = ans_sel[0].get_all_text().strip()
                            faqs.append({"question": question, "answer": answer})
            if not faqs:
                for h in page.css("h1, h2, h3, h4, h5, h6, p, span", adaptive=True):
                    text = h.get_all_text().strip()
                    if text.endswith("?") and len(text) > 10 and len(text) < 150:
                        ans_sel = h.xpath("./following-sibling::p[1]")
                        if ans_sel and len(ans_sel) > 0:
                            answer = ans_sel[0].get_all_text().strip()
                            if answer:
                                faqs.append({"question": text, "answer": answer})
            print(f"[FAQS] Extracted {len(faqs)} FAQs")

            # Contact Info
            contact_info = {
                "phone": "",
                "email": "",
                "address": "",
                "social_links": []
            }
            for a in page.css("a", adaptive=True):
                href = a.css("::attr(href)").get() or ""
                if href.startswith("mailto:"):
                    contact_info["email"] = href.replace("mailto:", "").split("?")[0].strip()
                elif href.startswith("tel:"):
                    contact_info["phone"] = href.replace("tel:", "").strip()
                elif any(social in href.lower() for social in ["facebook.com", "twitter.com", "linkedin.com", "instagram.com", "github.com"]):
                    contact_info["social_links"].append(href.strip())

            address_sel = page.css("address, .address, [class*='address']", adaptive=True)
            if address_sel:
                contact_info["address"] = address_sel[0].get_all_text().strip()

            # Products / Plans
            products = []
            plan_elements = page.css(".product, .pricing, .plan, .card, [class*='product-item'], [class*='pricing-plan']", adaptive=True)
            for pe in plan_elements:
                p_name = pe.css("h3::text, h4::text, .title::text, .name::text").get()
                p_price = pe.css(".price::text, [class*='price']::text").get()
                p_desc = pe.css(".description::text, .desc::text, p::text").get()
                if p_name:
                    products.append({
                        "name": p_name.strip(),
                        "price": p_price.strip() if p_price else "",
                        "description": p_desc.strip() if p_desc else ""
                    })

            # Team members
            team_members = []
            member_elements = page.css(".team-member, .member, [class*='team-member']", adaptive=True)
            for me in member_elements:
                m_name = me.css("h3::text, h4::text, .name::text").get()
                m_role = me.css(".role::text, .title::text, .position::text").get()
                if m_name:
                    team_members.append({
                        "name": m_name.strip(),
                        "role": m_role.strip() if m_role else ""
                    })

            # Reconstruct clean text content for backward compatibility
            soup = BeautifulSoup(html_content, "html.parser")
            decompose_selectors = [
                "script", "style", "meta", "noscript", "header", "footer", "nav", "aside", "svg", "iframe",
                "div[class*='cookie']", "div[id*='cookie']", "div[class*='banner']", "div[id*='banner']",
                "div[class*='ad-']", "div[id*='ad-']", "div[class*='ads']", "div[id*='ads']",
                "div[class*='footer']", "div[id*='footer']", "div[class*='header']", "div[id*='header']",
                "div[class*='nav']", "div[id*='nav']", "div[class*='tracking']", "div[id*='tracking']"
            ]
            for selector in decompose_selectors:
                for el in soup.select(selector):
                    el.decompose()

            cleaned_text_raw = soup.get_text()
            lines = (line.strip() for line in cleaned_text_raw.splitlines())
            chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
            cleaned_text = '\n'.join(chunk for chunk in chunks if chunk)

            # Detect page type
            page_type = self.detect_page_type(url_stripped, title, headings, cleaned_text)

            # Create logical sections
            sections = self.create_sections(title, headings, paragraphs, products, testimonials, team_members, faqs, contact_info)

            parse_time = time.time() - start_parse
            total_time = time.time() - start_total

            # Cache the structured output
            structured_res = {
                "title": title,
                "headings": headings,
                "subheadings": subheadings,
                "paragraphs": paragraphs,
                "images": images,
                "tables": tables,
                "lists": lists,
                "testimonials": testimonials,
                "faqs": faqs,
                "contact_info": contact_info,
                "products": products,
                "team_members": team_members,
                "page_type": page_type,
                "content": cleaned_text,
                "sections": sections,
                "source": source
            }
            SCRAPER_CACHE[url_stripped] = structured_res

            # Time metrics
            print(f"[SCRAPLING] Fetch Time: {fetch_time:.2f}s")
            print(f"[SCRAPLING] Extraction Time: {parse_time:.2f}s")
            print(f"[SCRAPLING] Total Time: {total_time:.2f}s")
            print(f"[SCRAPLING_SUCCESS] Extracted {len(cleaned_text)} chars")

            # Final Verification Summary
            print("=========================")
            print("SCRAPER SUMMARY")
            print("===============")
            print(f"URL: {url_stripped}")
            print(f"SOURCE: {source}")
            print(f"CHARS: {len(cleaned_text)}")
            print("CACHE: MISS")
            print(f"TOTAL TIME: {total_time:.2f}s")
            print("=================")

            return structured_res

        except Exception as e:
            print(f"[EXTRACT] Parsing/extraction failed: {str(e)}")
            # Ultimate backup extraction using basic BeautifulSoup directly
            try:
                print(f"[BS4_FALLBACK] Retrying extraction using basic BeautifulSoup directly")
                source = "beautifulsoup"
                plain_text = bs_extract_text_from_url(url_stripped)
                total_time = time.time() - start_total
                
                structured_res = {
                    "title": "Webpage Summary",
                    "headings": [],
                    "subheadings": [],
                    "paragraphs": [p.strip() for p in plain_text.split("\n") if p.strip()],
                    "images": [],
                    "tables": [],
                    "lists": [],
                    "testimonials": [],
                    "faqs": [],
                    "contact_info": {"phone": "", "email": "", "address": "", "social_links": []},
                    "products": [],
                    "team_members": [],
                    "page_type": "company website",
                    "content": plain_text,
                    "sections": [{"section_name": "Core Summary", "content": [plain_text[:500]]}],
                    "source": source
                }
                SCRAPER_CACHE[url_stripped] = structured_res
                print(f"[SCRAPLING_FAILURE] Scrapling failed, fell back to BeautifulSoup")
                
                # Final Verification Summary
                print("=========================")
                print("SCRAPER SUMMARY")
                print("===============")
                print(f"URL: {url_stripped}")
                print(f"SOURCE: {source}")
                print(f"CHARS: {len(plain_text)}")
                print("CACHE: MISS")
                print(f"TOTAL TIME: {total_time:.2f}s")
                print("=================")
                
                return structured_res
            except Exception as ex:
                print(f"[SCRAPLING_FAILURE] All scraper attempts failed")
                raise ValueError(f"Failed to parse or clean content from {url_stripped}: {str(e)}")

    def extract_text_from_url(self, url: str, structured: bool = False):
        """
        Legacy text extraction wrapper. Calls extract_structured_content internally
        and returns either raw content string or full structured dictionary depending
        on structured flag.
        """
        res = self.extract_structured_content(url)
        if structured:
            return res
        return res.get("content", "")
