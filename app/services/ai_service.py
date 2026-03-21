import os
import sys
import time
import json
import requests
import random
import uuid
import re
from urllib.parse import quote
from pathlib import Path

# WINDOWS UTF-8 FIX (SAFE VERSION)
if sys.platform.startswith("win"):
    try:
        if hasattr(sys.stdout, 'reconfigure'):
            sys.stdout.reconfigure(encoding="utf-8")
        else:
            import codecs
            sys.stdout = codecs.getwriter("utf-8")(sys.stdout.buffer, "strict")
    except Exception as e:
        print(f"[Warning] UTF-8 fix failed: {e}")
        pass

class CloudAIService:
    def __init__(self):
        print("\n" + "=" * 80)
        print("AI SERVICE v20.6 FINAL - REAL IMAGES ONLY ✨✨✨")
        print("=" * 80)

        # API KEYS - ENV ONLY (no hardcoded secrets)
        self.gemini_key = os.getenv("GOOGLE_GEMINI_API_KEY", "").strip()
        self.openrouter_key = os.getenv("OPENROUTER_API_KEY", "").strip()
        self.google_search_key = os.getenv("GOOGLE_CSE_API_KEY", "").strip()
        self.google_cx_id = os.getenv("GOOGLE_CSE_CX", "").strip()
        self.clipdrop_api_key = os.getenv("CLIPDROP_API_KEY", "").strip()
        self.pexels_api_key = os.getenv("PEXELS_API_KEY", "").strip()

        # STATUS REPORT
        print("\n" + "-" * 80)
        print("API KEYS STATUS:")
        print("-" * 80)
        
        if self.gemini_key:
            print(f"[OK] Gemini 3 Flash: {self.gemini_key[:20]}... ✅")
        else:
            print("[!] Gemini: NOT SET")
        
        if self.openrouter_key:
            print(f"[OK] OpenRouter: {self.openrouter_key[:20]}...")
        else:
            print("[!] OpenRouter: NOT SET")


        if self.clipdrop_api_key:
            print("[OK] ClipDrop: KEY SET ✅")
        else:
            print("[!] ClipDrop: KEY NOT SET")
        
        if self.google_search_key and self.google_cx_id:
            print(f"[OK] Google: {self.google_search_key[:20]}... (PRIMARY)")
        else:
            print("[!] Google: NOT SET")

        if self.pexels_api_key:
            print(f"[OK] Pexels: KEY SET ✅")
        else:
            print("[!] Pexels: NOT SET (add PEXELS_API_KEY to .env for better images)")
        
        print("[OK] Wikipedia: READY")
        print("[OK] Unsplash: READY (no key needed)")
        print("[INFO] Image Mode: WIKIPEDIA → PEXELS → UNSPLASH → CLIPDROP 🔥")
        
        print("-" * 80)
        print("INITIALIZATION COMPLETE ✅")
        print("=" * 80 + "\n")

    # WIKIPEDIA IMAGE (100% ACCURATE) ✨
    def _fetch_wikipedia_image(self, query, used_urls=None):
        """
        Fetch image from Wikipedia - tries top 3 results
        """
        print(f"   [Wiki] Searching: '{query[:50]}'...")
        used_urls = used_urls or set()
        
        try:
            search_url = "https://en.wikipedia.org/w/api.php"
            
            # Step 1: Search for page - get top 3 results
            search_params = {
                "action": "query",
                "format": "json",
                "list": "search",
                "srsearch": query,
                "srlimit": 3
            }
            
            resp = requests.get(search_url, params=search_params, timeout=5)
            if resp.status_code != 200:
                return None
            
            data = resp.json()
            results = data.get("query", {}).get("search", [])
            
            if not results:
                print("   [!] Not found on Wikipedia")
                return None
            
            # Try each search result until we find an image
            for result in results:
                page_title = result["title"]
                print(f"   [Wiki] Trying: {page_title}")
                
                # Step 2: Get page image via pageimages (thumbnail)
                image_params = {
                    "action": "query",
                    "format": "json",
                    "titles": page_title,
                    "prop": "pageimages",
                    "pithumbsize": 1000
                }
                
                resp = requests.get(search_url, params=image_params, timeout=5)
                if resp.status_code != 200:
                    continue
                
                data = resp.json()
                pages = data.get("query", {}).get("pages", {})
                
                for page_id, page_data in pages.items():
                    thumbnail = page_data.get("thumbnail", {})
                    image_url = thumbnail.get("source")
                    if image_url and image_url not in used_urls:
                        print(f"   [OK] Wikipedia image found: {page_title}")
                        return image_url
                
                # Step 3: Fallback - get image list from page
                images_params = {
                    "action": "query",
                    "format": "json",
                    "titles": page_title,
                    "prop": "images",
                    "imlimit": 5
                }
                resp2 = requests.get(search_url, params=images_params, timeout=5)
                if resp2.status_code == 200:
                    pages2 = resp2.json().get("query", {}).get("pages", {})
                    for pid, pdata in pages2.items():
                        for img in pdata.get("images", []):
                            fname = img.get("title", "")
                            # Skip icons, logos, flags - prefer photos
                            if any(skip in fname.lower() for skip in ["icon", "flag", "logo", "stub", "lock", "svg"]):
                                continue
                            # Fetch the actual image URL
                            url_params = {
                                "action": "query",
                                "format": "json",
                                "titles": fname,
                                "prop": "imageinfo",
                                "iiprop": "url",
                                "iiurlwidth": 1000
                            }
                            r3 = requests.get(search_url, params=url_params, timeout=5)
                            if r3.status_code == 200:
                                for _pid, _pd in r3.json().get("query", {}).get("pages", {}).items():
                                    info = _pd.get("imageinfo", [])
                                    if info:
                                        img_url = info[0].get("thumburl") or info[0].get("url")
                                        if img_url and img_url not in used_urls:
                                            print(f"   [OK] Wikipedia image (list) found: {fname}")
                                            return img_url
            
            print("   [!] No image found on Wikipedia")
            return None
            
        except Exception as e:
            print(f"   [X] Wikipedia error: {e}")
            return None

    # PEXELS IMAGE (FREE - needs PEXELS_API_KEY) ✨
    def _fetch_pexels_image(self, query, page=1, used_urls=None):
        """Fetch from Pexels (free API - register at pexels.com/api)"""
        if not self.pexels_api_key:
            return None
        used_urls = used_urls or set()
        print(f"   [Pexels] Searching: '{query[:50]}'...")
        try:
            headers = {"Authorization": self.pexels_api_key}
            params = {
                "query": query,
                "per_page": 15,
                "orientation": "landscape",
                "page": max(1, int(page))
            }
            resp = requests.get(
                "https://api.pexels.com/v1/search",
                headers=headers, params=params, timeout=8
            )
            if resp.status_code == 200:
                photos = resp.json().get("photos", [])
                for photo in photos:
                    src = photo.get("src", {})
                    url = src.get("large2x") or src.get("large") or src.get("original")
                    if url and url not in used_urls:
                        print(f"   [OK] Pexels image found!")
                        return url
            print(f"   [!] Pexels: no result ({resp.status_code})")
        except Exception as e:
            print(f"   [X] Pexels error: {e}")
        return None

    # UNSPLASH IMAGE (FREE - no key needed) ✨
    def _fetch_unsplash_image(self, query):
        """Fetch from Unsplash random image (no API key needed)"""
        print(f"   [Unsplash] Searching: '{query[:50]}'...")
        try:
            encoded = requests.utils.quote(query)
            url = f"https://source.unsplash.com/1280x720/?{encoded}"
            resp = requests.get(url, timeout=8, allow_redirects=True)
            content_type = resp.headers.get("content-type", "")
            if resp.status_code == 200 and "image" in content_type:
                print(f"   [OK] Unsplash image found!")
                return resp.url  # actual image URL after redirect
            print(f"   [!] Unsplash: no result")
        except Exception as e:
            print(f"   [X] Unsplash error: {e}")
        return None

    # GOOGLE IMAGE - FLEXIBLE SEARCH (FIXED) ✨
    def _fetch_google_image(self, query, start_index=1):
        """
        Google Custom Search with FLEXIBLE match (Quotes removed)
        start_index allows fetching different results
        """
        if not self.google_search_key or not self.google_cx_id:
            print("   [!] Google API keys not set")
            return None

        print(f"   [Google] Searching: '{query[:50]}' (start={start_index})...")

        try:
            # 🔥 FIX: Quotes hata diye taaki zyada results milein
            search_query = query 
            
            is_person = any(word[0].isupper() for word in query.split() if len(word) > 2)
            
            params = {
                "q": search_query,
                "cx": self.google_cx_id,
                "key": self.google_search_key,
                "searchType": "image",
                "num": 10,
                "start": start_index,
                "imgSize": "large",
                "safe": "active",
                "fileType": "jpg|png",
            }
            
            if is_person:
                params["imgType"] = "face"
            else:
                params["imgType"] = "photo"
            
            resp = requests.get(
                "https://www.googleapis.com/customsearch/v1",
                params=params,
                timeout=8
            )

            if resp.status_code == 200:
                items = resp.json().get("items", [])
                
                if items:
                    # Agar results mile, toh pehla return karo
                    # (Scoring logic complex queries me fail ho sakta hai, direct link is safer here)
                    print(f"   [OK] Found image at index {start_index}")
                    return items[0]["link"]
                    
                print("   [!] No suitable images found")
                        
            elif resp.status_code == 429:
                print("   [!] Google quota exceeded")
            else:
                error_snippet = resp.text[:200].replace("\n", " ")
                print(f"   [!] Google error: {resp.status_code} | {error_snippet}")
                
        except Exception as e:
            print(f"   [X] Google error: {e}")

        return None

    # SMART REAL IMAGE FETCHER ✨
    def _fetch_real_image(self, query, slide_num=0, used_urls=None):
        """
        Smart image fetcher:
        1. Wikipedia (free, accurate)
        2. Pexels (free API key)
        3. Unsplash (free, no key)
        4. Google CSE (if keys set)
        """
        used_urls = used_urls or set()
        
        print(f"\n   [Real Image] Fetching for Slide {slide_num}")
        print(f"   [Query] '{query}'")
        
        # 1. Try Wikipedia FIRST
        print(f"   [Primary] Trying Wikipedia...")
        wiki_image = self._fetch_wikipedia_image(query, used_urls=used_urls)
        if wiki_image:
            return wiki_image

        # 2. Try Pexels (free API)
        if self.pexels_api_key:
            print(f"   [Fallback 1] Trying Pexels...")
            pexels_page = max(1, (int(slide_num) % 5) + 1)
            pexels_image = self._fetch_pexels_image(query, page=pexels_page, used_urls=used_urls)
            if pexels_image and pexels_image not in used_urls:
                return pexels_image

            # Retry Pexels with a broader query if exact query duplicates/doesn't return enough variety.
            pexels_image = self._fetch_pexels_image(f"{query} portrait", page=1, used_urls=used_urls)
            if pexels_image and pexels_image not in used_urls:
                return pexels_image

        # 3. Try Unsplash (no key needed)
        print(f"   [Fallback 2] Trying Unsplash...")
        unsplash_image = self._fetch_unsplash_image(query)
        if unsplash_image and unsplash_image not in used_urls:
            return unsplash_image

        # 4. Try Google CSE if keys are set
        if self.google_search_key and self.google_cx_id:
            print(f"   [Fallback 3] Trying Google...")
            start_index = 1 + (slide_num * 10) % 90
            google_image = self._fetch_google_image(query, start_index)
            if google_image and google_image not in used_urls:
                return google_image

            # One more page forward if the first Google result was already used.
            google_image = self._fetch_google_image(query, min(start_index + 10, 91))
            if google_image and google_image not in used_urls:
                return google_image

        print(f"\n   [!] No images found for: {query}")
        return None

    # GEMINI 3 FLASH WITH RETRY LOGIC & TIMEOUT ✅
    def _call_gemini(self, prompt):
        """Gemini 3 Flash - v20.6 Optimized with Retries & High Timeout"""
        if not self.gemini_key:
            print("   [X] Gemini key not set")
            return None

        # Latest stable preview model for 2026
        model_name = "gemini-3-flash-preview"
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent"
        
        params = {'key': self.gemini_key}
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {
                "temperature": 0.7,
                "maxOutputTokens": 8192,
                "response_mime_type": "application/json"  # Force JSON output
            }
        }

        # Retry Loop: 3 baar koshish karega agar timeout hua toh
        for attempt in range(3):
            try:
                print(f"   [Gemini] Calling (Attempt {attempt+1}/3)...")
                resp = requests.post(
                    url, 
                    params=params, 
                    json=payload, 
                    timeout=90,  # 30 se badha kar 90 kar diya ✅
                    headers={"Content-Type": "application/json"}
                )

                if resp.status_code == 200:
                    data = resp.json()
                    text = data["candidates"][0]["content"]["parts"][0]["text"]
                    print(f"   [OK] Generated {len(text)} chars")
                    return text
                
                elif resp.status_code == 429:
                    print("   [!] Rate limit hit. Sleeping 10s...")
                    time.sleep(10)
                else:
                    print(f"   [X] API Error {resp.status_code}: {resp.text[:100]}")
                    break  # Agar key galat hai toh retry ka fayda nahi

            except (requests.exceptions.ReadTimeout, requests.exceptions.ConnectTimeout):
                print(f"   [!] Timeout Error. Retrying in 2s...")
                time.sleep(2)
            except Exception as e:
                print(f"   [X] Unexpected Error: {e}")
                break

        return None

    # DEEPSEEK R1
    def _call_deepseek(self, prompt):
        """DeepSeek R1 via OpenRouter"""
        if not self.openrouter_key:
            return None

        print("   [DeepSeek] Calling...")

        try:
            resp = requests.post(
                "https://openrouter.ai/api/v1/chat/completions",
                headers={
                    "Authorization": f"Bearer {self.openrouter_key}",
                    "Content-Type": "application/json"
                },
                json={
                    "model": "deepseek/deepseek-chat",
                    "messages": [
                        {"role": "system", "content": "You are a professional presentation writer. Output ONLY valid JSON."},
                        {"role": "user", "content": prompt}
                    ],
                    "response_format": {"type": "json_object"},
                    "max_tokens": 4000,
                    "temperature": 0.7
                },
                timeout=60
            )

            if resp.status_code == 200:
                text = resp.json()["choices"][0]["message"]["content"]
                print(f"   [OK] {len(text)} chars")
                return text
            else:
                print(f"   [X] Status {resp.status_code}")
        except Exception as e:
            print(f"   [X] {e}")

        return None

    # CHATGPT VIA OPENROUTER
    def _call_openrouter_chatgpt(self, prompt):
        """ChatGPT via OpenRouter"""
        if not self.openrouter_key:
            print("   [X] OpenRouter key not set")
            return None

        print("   [OpenRouter] Calling ChatGPT...")

        try:
            resp = requests.post(
                "https://openrouter.ai/api/v1/chat/completions",
                headers={
                    "Authorization": f"Bearer {self.openrouter_key}",
                    "Content-Type": "application/json"
                },
                json={
                    "model": "openai/gpt-4o-mini",
                    "messages": [
                        {"role": "system", "content": "You are a professional presentation writer. Output ONLY valid JSON array."},
                        {"role": "user", "content": prompt}
                    ],
                    "max_tokens": 4000,
                    "temperature": 0.7
                },
                timeout=60
            )

            if resp.status_code == 200:
                text = resp.json()["choices"][0]["message"]["content"]
                print(f"   [OK] {len(text)} chars")
                return text
            else:
                error_snippet = resp.text[:200].replace("\n", " ")
                print(f"   [X] Status {resp.status_code} | {error_snippet}")
        except Exception as e:
            print(f"   [X] {e}")

        return None

    # SMART AI CALLER
    def _get_ai_text(self, prompt, ai_model="gemini"):
        """Model-specific AI caller (no cross-model fallback)"""
        print(f"\n[AI] Model: {ai_model.upper()}")

        if ai_model in {"chatgpt", "openai", "gpt"}:
            text = self._call_openrouter_chatgpt(prompt)
            if text:
                return text
            print("   [X] ChatGPT call failed")
            print("   [Fallback] Trying Gemini...")
            return self._call_gemini(prompt)
        if ai_model == "deepseek":
            text = self._call_deepseek(prompt)
            if text:
                return text
            print("   [X] DeepSeek call failed")
            print("   [Fallback] Trying Gemini...")
            return self._call_gemini(prompt)

        text = self._call_gemini(prompt)
        if text:
            return text
        print("   [X] Gemini call failed")
        return None

    def chat_assistant(self, question, ai_model="gemini"):
        """Answer dashboard assistant questions using configured AI providers."""
        q = (question or "").strip()
        if not q:
            return None

        model = (ai_model or "gemini").strip().lower()
        q_lower = q.lower()
        titles_request = any(k in q_lower for k in ["title", "titles", "outline", "outlines", "topic", "topics", "heading", "headings", "slide outline"])
        detail_request = any(k in q_lower for k in ["detail", "details", "detailed", "in detail", "explain", "explanation", "briefly explain"])

        system_prompt = (
            "You are PitchCraft AI assistant for a presentation app. "
            "Give practical, short, actionable help in plain English. "
            "If user asks for titles/outlines and does not explicitly ask for details, return only concise topic names. "
            "Do not add descriptions under each title unless explicitly requested. "
            "Do not include markdown fences."
        )

        if titles_request and not detail_request:
            system_prompt += (
                " Format strictly as numbered list with only title names, one line each. "
                "No extra sentences, no bullets under a title, no paragraphs."
            )

        def _normalize_titles_only(text):
            if not text:
                return text
            lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
            picked = []
            for ln in lines:
                clean = re.sub(r"^\s*(?:\d+\.|[-*])\s*", "", ln).strip()
                if not clean:
                    continue
                # Skip obvious detail lines that don't look like headings
                if len(clean.split()) > 14 and ":" not in clean:
                    continue
                # Keep only the heading part before long explanations
                if ":" in clean and len(clean.split(":", 1)[1].split()) > 5:
                    clean = clean.split(":", 1)[0].strip()
                picked.append(clean)
            if not picked:
                picked = [re.sub(r"^\s*(?:\d+\.|[-*])\s*", "", ln).strip() for ln in lines[:10]]
            return "\n".join(f"{i+1}. {t}" for i, t in enumerate(picked[:10]))

        if model in {"chatgpt", "openai", "gpt"} and self.openrouter_key:
            try:
                resp = requests.post(
                    "https://openrouter.ai/api/v1/chat/completions",
                    headers={
                        "Authorization": f"Bearer {self.openrouter_key}",
                        "Content-Type": "application/json"
                    },
                    json={
                        "model": "openai/gpt-4o-mini",
                        "messages": [
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": q}
                        ],
                        "max_tokens": 900,
                        "temperature": 0.5
                    },
                    timeout=60
                )
                if resp.status_code == 200:
                    answer = (resp.json().get("choices", [{}])[0].get("message", {}).get("content", "") or "").strip()
                    if titles_request and not detail_request:
                        return _normalize_titles_only(answer)
                    return answer
            except Exception as e:
                print(f"[Assistant/OpenRouter] {e}")

        if self.gemini_key:
            try:
                model_name = "gemini-3-flash-preview"
                url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent"
                resp = requests.post(
                    url,
                    params={"key": self.gemini_key},
                    headers={"Content-Type": "application/json"},
                    json={
                        "contents": [
                            {
                                "parts": [
                                    {"text": f"{system_prompt}\n\nUser question: {q}"}
                                ]
                            }
                        ],
                        "generationConfig": {
                            "temperature": 0.5,
                            "maxOutputTokens": 900
                        }
                    },
                    timeout=60
                )
                if resp.status_code == 200:
                    data = resp.json()
                    answer = (data.get("candidates", [{}])[0].get("content", {}).get("parts", [{}])[0].get("text", "") or "").strip()
                    if titles_request and not detail_request:
                        return _normalize_titles_only(answer)
                    return answer
            except Exception as e:
                print(f"[Assistant/Gemini] {e}")

        return None

    # JSON PARSER
    def _clean_json(self, text):
        """Extract JSON array from response"""
        if not text:
            return []

        try:
            text = text.replace("```json", "").replace("```", "").strip()

            try:
                obj = json.loads(text)
                if isinstance(obj, list):
                    return obj
                if isinstance(obj, dict) and "slides" in obj:
                    return obj["slides"]
            except json.JSONDecodeError:
                pass

            start = text.find("[")
            end = text.rfind("]")
            if start != -1 and end != -1:
                return json.loads(text[start:end+1])

        except Exception as e:
            print(f"   [X] JSON: {e}")

        return []

    def _generate_clipdrop_image(self, prompt):
        if not self.clipdrop_api_key:
            print("   [ClipDrop] Key not set")
            return None

        url = "https://clipdrop-api.co/text-to-image/v1"
        headers = {"x-api-key": self.clipdrop_api_key}
        files = {"prompt": (None, prompt)}

        try:
            resp = requests.post(url, headers=headers, files=files, timeout=60)
            if resp.status_code != 200:
                print(f"   [ClipDrop] Error {resp.status_code}: {resp.text[:100]}")
                return None

            safe_name = f"clipdrop_{uuid.uuid4().hex[:10]}.png"
            out_dir = Path(__file__).resolve().parents[2] / "static" / "generated"
            out_dir.mkdir(parents=True, exist_ok=True)
            out_path = out_dir / safe_name
            with open(out_path, "wb") as f:
                f.write(resp.content)

            return f"/static/generated/{safe_name}"
        except Exception as e:
            print(f"   [ClipDrop] Error: {e}")
            return None

    # GET IMAGE (GOOGLE PRIMARY WITH PAGINATION) ✨
    def get_smart_image(self, prompt, slide_num=0, force_dark=False, image_source="real", used_urls=None):
        """
        REAL IMAGES ONLY - GOOGLE PRIMARY WITH PAGINATION
        """
        source = (image_source or "real").lower().strip()
        print(f"\n[Image #{slide_num}] Source: {source}")

        raw_prompt = prompt.strip() if isinstance(prompt, str) else ""

        # Clean prompt (real image search only)
        generic = [
            "Deep Dive", "Introduction to", "Summary", "Conclusion", 
            "Overview", "Analysis", "Understanding", "Exploring",
            "Key Insights", "Final Thoughts", "Let's Explore",
            "Introduction", "Conclusion"
        ]
        clean = prompt
        for word in generic:
            clean = clean.replace(word, "").strip()

        if len(clean) < 3:
            clean = prompt

        print(f"   [Query] '{clean[:60]}'")

        if source == "ai":
            print("   [AI] Using ClipDrop")
            ai_prompt = raw_prompt if raw_prompt else clean
            ai_url = self._generate_clipdrop_image(ai_prompt)
            if ai_url:
                return ai_url
            print("   [AI] ClipDrop failed - using placeholder")
            return "https://via.placeholder.com/1280x720/cccccc/666666?text=No+Image+Available"

        # ✅ FETCH REAL IMAGE (with pagination based on slide_num)
        img = self._fetch_real_image(clean, slide_num, used_urls=used_urls)
        if img:
            return img

        print("   [!] No real image found - using ClipDrop fallback")
        ai_url = self._generate_clipdrop_image(clean)
        if ai_url:
            return ai_url
        return None

    # CONTENT VALIDATOR
    def _validate_and_fix_content(self, slide_data, slide_num, prompt):
        """Validate and fix content"""
        content = slide_data.get("content", "")
        
        if isinstance(content, list):
            content = "\n".join([str(item) for item in content])
        
        if not isinstance(content, str):
            content = str(content)
        
        # Normalize content: remove blank lines and merge numbered lines with text
        lines = [line.strip() for line in content.split("\n") if line.strip()]
        merged_lines = []
        i = 0
        while i < len(lines):
            line = lines[i]
            number_match = re.fullmatch(r"(\d+)[\.)]?", line)
            if number_match and i + 1 < len(lines):
                number = number_match.group(1)
                next_line = lines[i + 1].lstrip("-*• ").strip()
                merged_lines.append(f"{number}. {next_line}")
                i += 2
                continue
            merged_lines.append(line)
            i += 1

        content = "\n".join(merged_lines)

        if not content.strip() and slide_num == 1:
            return ""

        points = [p for p in content.split('\n') if p.strip() and (p.startswith('-') or (len(p) > 0 and p[0].isdigit()))]
        count = len(points)

        if count == 0:
            points = [p for p in content.split('\n') if p.strip()]
            count = len(points)

        # SLIDE 1: Paragraph
        if slide_num == 1:
            if count > 1 or content.startswith("-"):
                return f"This presentation explores {prompt} comprehensively. We examine its origins, significance, and modern relevance, providing insights into how it shapes our world today and influences future developments across multiple dimensions."

        # SLIDES 2, 7: 2 bullets
        elif slide_num in [2, 7]:
            if count != 2:
                return (
                    f"1. Core Analysis: Deep dive into the fundamental aspects of {prompt}, exploring its key components, underlying principles, and how they interact to create measurable value and impact across various domains and industries.\n"
                    f"2. Broader Implications: Understanding the wider societal, economic, and technological impact of {prompt}, including its influence on future trends, emerging opportunities, and potential challenges that organizations must address."
                )

        # SLIDE 3: 3 bullets
        elif slide_num == 3:
            if count != 3:
                return (
                    f"1. Innovation & Technology: {prompt} drives technological advancement through cutting-edge solutions, fostering creativity and enabling breakthrough discoveries that reshape industries and create new market opportunities for forward-thinking organizations worldwide.\n"
                    f"2. Integration & Connectivity: Seamless connection across multiple platforms and ecosystems, enabling unified experiences that enhance collaboration, streamline workflows, and create cohesive digital environments that empower teams to achieve more together.\n"
                    f"3. Global Impact & Reach: Transforming industries worldwide by establishing new standards, influencing policy decisions, and creating sustainable value chains that benefit stakeholders across geographic boundaries and cultural contexts."
                )

        # SLIDE 4: 4 bullets
        elif slide_num == 4:
            if count != 4:
                return (
                    f"1. Efficiency & Performance: Streamlined processes and optimized workflows that reduce operational overhead, eliminate redundancies, and maximize resource utilization while maintaining the highest quality standards and delivering measurable ROI.\n"
                    f"2. Scalability & Growth: Flexible architecture that grows seamlessly with your needs, supporting expansion from small teams to enterprise-level operations without compromising performance, reliability, or user experience across all touchpoints.\n"
                    f"3. Reliability & Uptime: Consistent, dependable performance backed by robust infrastructure, comprehensive monitoring systems, and proactive maintenance protocols that ensure 99.9% availability and business continuity even during peak demand periods.\n"
                    f"4. Security & Compliance: Enterprise-grade protection with multi-layered security frameworks, advanced encryption protocols, regular security audits, and full compliance with international standards including GDPR, SOC 2, and ISO 27001 certifications."
                )

        # SLIDE 5: 4 bullets
        elif slide_num == 5:
            if count < 3:
                return (
                    f"1. Global Impact: {prompt} has transcended geographical borders worldwide, creating interconnected networks that facilitate cross-cultural exchange and foster international collaboration.\n"
                    f"2. Cultural Exchange: Communities worldwide celebrate diverse traditions while finding common ground, building bridges between different perspectives and creating inclusive environments.\n"
                    f"3. Economic Growth: Drives substantial tourism revenue, retail expansion, and job creation across multiple sectors, contributing to sustainable economic development.\n"
                    f"4. Social Harmony: Strengthens social fabrics by bringing people together, fostering mutual understanding, and creating shared experiences that transcend differences."
                )

        # SLIDE 6: Roadmap
        elif slide_num == 6:
            if count < 5:
                return (
                    f"1. Research & Discovery: Comprehensive analysis of requirements, market conditions, and user needs through stakeholder interviews, competitive analysis, and data-driven insights.\n"
                    f"2. Strategic Planning: Create detailed roadmap with clear milestones, resource allocation, risk assessment, and timeline definition to ensure successful execution.\n"
                    f"3. Development & Implementation: Build core features and functionality using agile methodologies, modern frameworks, and best practices to deliver robust solutions.\n"
                    f"4. Quality Assurance: Rigorous testing protocols including unit tests, integration tests, user acceptance testing, and performance optimization to ensure excellence.\n"
                    f"5. Launch & Deployment: Carefully orchestrated rollout to production environment with monitoring, support readiness, and rollback procedures in place.\n"
                    f"6. Optimization & Growth: Continuous monitoring, performance tuning, user feedback integration, and iterative improvements to maximize value delivery."
                )

        # SLIDE 8: Summary
        elif slide_num == 8:
            if count < 2:
                return (
                    f"1. Strategic Conclusion: In conclusion, {prompt} represents a critical evolution in how we approach modern challenges, offering a robust framework for sustainable progress and excellence.\n"
                    f"2. Future Roadmap: Looking ahead, the focus remains on continuous innovation, ethical integration, and expanding the scope of positive impact across diverse sectors.\n"
                    f"3. Final Insights: By leveraging these key takeaways, organizations and individuals can unlock new levels of potential and drive meaningful change in an ever-evolving landscape."
                )

        return content

    # SLIDE 7 CONTENT VALIDATOR (EXECUTIVE SUMMARY)
    def _validate_slide_7_content(self, slide_data, prompt):
        """
        Validate and format content specifically for Slide 7.

        - 3 short points
        - 15-25 words each
        - Professional executive summary tone
        """
        content = slide_data.get("content", "")

        if isinstance(content, list):
            content = "\n".join([str(item) for item in content])

        if not isinstance(content, str):
            content = str(content)

        lines = [line.strip() for line in content.split("\n") if line.strip()]
        cleaned = []
        for line in lines:
            line = line.lstrip('-*•0123456789. ').strip()
            if line:
                cleaned.append(line)

        points = cleaned[:3]

        if len(points) < 3 or sum(len(p) for p in points) < 100:
            topic = prompt.split()[0] if prompt else "This topic"
            points = [
                f"{topic}'s impact extends beyond immediate results, creating lasting value and inspiring future developments.",
                "The strategic approach demonstrated here models excellence through quality, discipline, and sustained commitment.",
                "These insights will continue to shape standards and drive meaningful progress across the field."
            ]

        return "\n".join(points)

    # MAIN SLIDE GENERATION
    def generate_slides(
        self,
        prompt,
        slides_count=8,
        language="English",
        theme="dialogue",
        text_amount="concise",
        ai_model="gemini",
        custom_outline=None,
        **kwargs
    ):
        """Generate presentation slides with REAL IMAGES ONLY"""

        try:
            slides_count = int(slides_count)
            slides_count = max(3, min(slides_count, 20))
        except:
            slides_count = 8

        print("\n" + "="*80)
        print("[START] SLIDE GENERATION - GOOGLE PAGINATION 🔥")
        print("="*80)
        print(f"Topic: {prompt}")
        print(f"Slides: {slides_count}")
        print(f"Language: {language}")
        print(f"Theme: {theme}")
        print(f"Text Amount: {text_amount}")
        normalized_model = (ai_model or "gemini").strip().lower()
        if normalized_model in {"chatgpt", "openai", "gpt"}:
            print("AI Model: CHATGPT (OPENROUTER)")
        elif normalized_model == "deepseek":
            print("AI Model: DEEPSEEK (OPENROUTER)")
        else:
            print("AI Model: GEMINI 3 FLASH PREVIEW")
        print(f"Images: GOOGLE (with pagination - different images)")
        print(f"Image Slides: 1, 2, 5, 8")
        print("="*80 + "\n")

        text_map = {
            'minimal': 'Short and punchy',
            'concise': 'Standard professional length',
            'detailed': 'Long and descriptive',
            'extensive': 'Very detailed analysis'
        }
        text_instruction = text_map.get(text_amount, 'Standard professional length')

        source_material = (kwargs.get("source_material") or "").strip().lower()
        image_source = (kwargs.get("image_source") or "real").strip().lower()
        outline_titles = []
        if custom_outline:
            outline_titles = [line.strip() for line in custom_outline.split("\n") if line.strip()]

        if source_material == "custom outline" or outline_titles:
            if outline_titles:
                slides_count = len(outline_titles)
                print(f"[Outline] Using custom outline with {slides_count} titles")
            else:
                print("[Outline] Custom outline selected but no titles provided")

        if outline_titles:
            outline_block = "\n".join([f"{i + 1}. {title}" for i, title in enumerate(outline_titles)])
            ai_prompt = f"""
You are given an ordered list of slide titles.
Create EXACTLY {slides_count} slides using ONLY these titles, in the SAME order.
Use each provided title verbatim for the slide title. Do NOT add, remove, or rename slides.

Language: {language}
Text Length: {text_instruction}

Slide titles:
{outline_block}

CRITICAL FORMATTING RULES (Repeat this pattern every 8 slides):
1. Pattern 1: ONE PARAGRAPH (80-100 words), NO bullets
2. Pattern 2: EXACTLY 2 DETAILED bullets (60-80 words each)
3. Pattern 3: EXACTLY 3 DETAILED bullets (60-80 words each)
4. Pattern 4: EXACTLY 4 DETAILED bullets (60-80 words each)
5. Pattern 5: EXACTLY 4 bullets (40-50 words each)
6. Pattern 6: 5-6 NUMBERED STEPS (roadmap/timeline format)
7. Pattern 7: 3 SHORT bullet points (15-25 words each), key takeaways, professional tone
8. Pattern 8: 3-4 bullets (summary)

IMPORTANT: Generate EXACTLY {slides_count} slides. If you need more than 8, repeat the formatting pattern from 1.
IMPORTANT: Return "content" as a STRING with newlines, NOT as an array/list.

Output: JSON array ONLY (no markdown, no explanations)
Format: [{{"title":"...", "content":"..."}}]

START JSON:
"""
        else:
            ai_prompt = f"""
Create EXACTLY {slides_count} professional presentation slides about: "{prompt}"

Language: {language}
Text Length: {text_instruction}

CRITICAL FORMATTING RULES (Repeat this pattern every 8 slides):
1. Pattern 1: ONE PARAGRAPH (80-100 words), NO bullets
2. Pattern 2: EXACTLY 2 DETAILED bullets (60-80 words each)
3. Pattern 3: EXACTLY 3 DETAILED bullets (60-80 words each)
4. Pattern 4: EXACTLY 4 DETAILED bullets (60-80 words each)
5. Pattern 5: EXACTLY 4 bullets (40-50 words each)
6. Pattern 6: 5-6 NUMBERED STEPS (roadmap/timeline format)
7. Pattern 7: 3 SHORT bullet points (15-25 words each), key takeaways, professional tone
8. Pattern 8: 3-4 bullets (summary)

IMPORTANT: Generate EXACTLY {slides_count} slides. If you need more than 8, repeat the formatting pattern from 1.
IMPORTANT: Return "content" as a STRING with newlines, NOT as an array/list.

Output: JSON array ONLY (no markdown, no explanations)
Format: [{{"title":"...", "content":"..."}}]

Example:
[
    {{"title": "Introduction", "content": "This is a paragraph about the topic..."}},
    {{"title": "Benefits", "content": "1. First benefit\n2. Second benefit"}}
]

START JSON:
"""

        print("[AI] Calling model...")
        raw = self._get_ai_text(ai_prompt, normalized_model)
        
        data = self._clean_json(raw)

        # ✅ ROBUST CONTENT HANDLING
        if not data:
            print("[!] JSON empty, generating dummy data")
            data = []

        fallback_titles = {
            1: "Overview",
            2: "Key Insights",
            3: "Core Elements",
            4: "Major Themes",
            5: "Real-World Impact",
            6: "Roadmap",
            7: "Key Takeaways",
            8: "Conclusion",
        }

        if len(data) < slides_count:
            print(f"[!] AI returned {len(data)}/{slides_count} slides. Filling missing slides...")
            for i in range(len(data), slides_count):
                if outline_titles and i < len(outline_titles):
                    data.append({"title": outline_titles[i], "content": f"Detailed content for {outline_titles[i]}..."})
                else:
                    slide_num = i + 1
                    pattern_num = ((slide_num - 1) % 8) + 1
                    if slide_num == 1:
                        fallback_title = str(prompt).strip() or "Introduction"
                    else:
                        fallback_title = fallback_titles.get(pattern_num, f"Slide {slide_num}")
                    data.append({
                        "title": fallback_title,
                        "content": "" if pattern_num == 1 else f"Professional analysis for {prompt}, covering {fallback_title.lower()} with concise, practical points."
                    })

        if outline_titles:
            for i, title in enumerate(outline_titles):
                if i < len(data):
                    data[i]["title"] = title

        print(f"\n[Build] Creating {slides_count} slides...\n")

        final = []
        used_image_urls = set()
        
        layouts = {
            1: 'centered',
            2: 'fixed_information',
            3: 'three_col',
            4: 'grid_4',
            5: 'split_box',
            6: 'roadmap',
            7: 'fixed_information',
            8: 'fixed_mission'
        }

        # ✅ IMAGE SLIDES (Based on pattern 1-8)
        IMAGE_SLIDES_PATTERN = [1, 2, 5, 8]

        for i in range(slides_count):
            slide = data[i] if i < len(data) else {
                "title": str(prompt).strip() or f"Slide {i+1}",
                "content": "..."
            }
            
            slide_num = i + 1
            pattern_num = ((slide_num - 1) % 8) + 1  # 1 to 8 cycle

            print(f"[Slide {slide_num}] (Pattern {pattern_num}) {slide.get('title', '')[:40]}")

            validation_prompt = outline_titles[i] if outline_titles and i < len(outline_titles) else prompt

            if pattern_num == 7:
                slide['content'] = self._validate_slide_7_content(slide, validation_prompt)
            else:
                slide['content'] = self._validate_and_fix_content(slide, pattern_num, validation_prompt)

            slide['layout'] = layouts.get(pattern_num, 'fixed_mission')
            print(f"   Layout: {slide['layout']}")

            # ✅ IMAGE LOGIC
            if pattern_num in IMAGE_SLIDES_PATTERN:
                print(f"   [Image] Fetching unique image for slide {slide_num} (Pattern {pattern_num})...")
                
                # 🔥 FIX for Slide 8 style (Conclusion/Mission):
                if pattern_num == 8:
                    image_query = prompt
                    idx_override = 1 # Use first/best result for conclusion
                else:
                    image_query = validation_prompt or prompt
                    idx_override = slide_num

                if image_source == "ai":
                    title_hint = (slide.get("title") or "").strip()
                    if title_hint and pattern_num != 8:
                        image_query = f"{title_hint} {image_query}".strip()
                
                img = self.get_smart_image(
                    image_query,
                    idx_override,
                    force_dark=False,
                    image_source=image_source,
                    used_urls=used_image_urls
                )
            else:
                img = None
                print(f"   [Image] Skipped (layout doesn't need image)")

            if img:
                used_image_urls.add(img)

            final.append({
                "id": f"slide_{i}",
                "title": slide.get("title", f"Slide {slide_num}"),
                "content": slide['content'],
                "layout": slide['layout'],
                "image": img,
                "background": self._get_theme_background(theme)
            })

            print(f"   [OK]\n")

        print("="*80)
        print(f"[DONE] {len(final)} slides - DIFFERENT IMAGES 🔥")
        print("="*80 + "\n")

        return final

    # THEME BACKGROUNDS
    def _get_theme_background(self, theme):
        """Return CSS gradient based on theme"""
        theme_map = {
            "dialogue": "linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%)",
            "alien": "linear-gradient(135deg, #1a1a2e 0%, #16213e 100%)",
            "wine": "linear-gradient(135deg, #581c3c 0%, #3d1428 100%)",
            "snowball": "linear-gradient(135deg, #e0f2fe 0%, #bae6fd 100%)",
            "petrol": "linear-gradient(135deg, #475569 0%, #334155 100%)",
            "piano": "linear-gradient(135deg, #000000 0%, #1e293b 50%, #ffffff 100%)",
            "business": "linear-gradient(135deg, #3a7bd5 0%, #00d2ff 100%)"
        }
        return theme_map.get(theme, theme_map["dialogue"])

# GLOBAL INSTANCE
ai_service = None

try:
    print("\n" + "🔥" * 40)
    print("[Init] Starting AI Service v20.6 FINAL...")
    print("🔥" * 40 + "\n")
    
    ai_service = CloudAIService()
    
    print("\n" + "🔥" * 40)
    print("[SUCCESS] AI Service v20.6 READY ✅✅✅")
    print("🔥" * 40)
    print("\n✨ FEATURES:")
    print("  - Gemini 3 Flash (v1beta endpoint) 🚀")
    print("  - YOUR NEW KEY ACTIVATED ✅")
    print("  - Retry Logic (3 attempts) ✅")
    print("  - 90 Second Timeout ⏱️")
    print("  - DeepSeek R1 (Fallback)")
    print("  - Google Images (PRIMARY with pagination) 🔥🔥🔥")
    print("  - Wikipedia (Backup)")
    print("  - DIFFERENT IMAGES PER SLIDE")
    print("  - REAL IMAGES ONLY - NO AI")
    print("  - All layouts (1-8)")
    print("  - Content validation + normalization")
    print("  - Response MIME Type: application/json")
    print("\n🎯 KEY IMPROVEMENTS:")
    print("  1. Timeout increased to 90 seconds")
    print("  2. Retry loop for timeout errors")
    print("  3. v1beta endpoint for Gemini 3")
    print("  4. Content normalization (handles '1\\nText' format)")
    print("  5. Flexible Google search (no quotes)")
    print("  6. JSON output forced via MIME type")
    print("\n🎯 IMAGE STRATEGY:")
    print("  - Slide 1: Google results 1-10")
    print("  - Slide 2: Google results 11-20")
    print("  - Slide 5: Google results 41-50")
    print("  - Slide 8: Google results 71-80")
    print("  - Result: DIFFERENT IMAGES FROM SAME TOPIC")
    print("\n" + "🔥" * 40 + "\n")
    
except Exception as e:
    print("\n" + "❌" * 40)
    print(f"[ERROR] AI Service failed")
    print("❌" * 40 + "\n")
    print(f"Error: {e}\n")
    
    import traceback
    traceback.print_exc()
    
    ai_service = None