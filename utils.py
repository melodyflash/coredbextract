import os
import tempfile

def save_uploaded_file(uploaded_file):
    """
    Saves a specific uploaded file to a temporary location.
    Returns the absolute path to the saved file.
    """
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{uploaded_file.name.split('.')[-1]}") as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            return tmp_file.name
    except Exception as e:
        print(f"Error saving file: {e}")
        return None

def cleanup_temp_file(file_path):
    """Remove temporary file"""
    if file_path and os.path.exists(file_path):
        try:
            os.remove(file_path)
        except Exception:
             pass

import requests
from bs4 import BeautifulSoup

def scrape_url(url: str) -> str:
    """
    Fetch and parse text from a URL, including JSON-LD data which often contains menu structures.
    """
    try:
        # User-Agent to look like a real browser
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.9'
        }
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # 1. Extract JSON-LD (often has Schema.org Menu data)
        json_ld_text = ""
        for script in soup.find_all('script', type='application/ld+json'):
            if script.string:
                json_ld_text += f"\n--- JSON-LD DATA ---\n{script.string}\n"

        # 3. specific React/Next.js data blobs
        for script in soup.find_all('script'):
            if script.get('id') == '__NEXT_DATA__' or script.get('id') == '__Nuxt__':
                 json_ld_text += f"\n--- APP DATA ({script.get('id')}) ---\n{script.string}\n"

        # 2. Clean Text extraction
        for script in soup(["script", "style", "noscript", "iframe"]):
            script.decompose()
            
        text = soup.get_text(separator=' ')
        
        # Cleanup whitespace
        lines = (line.strip() for line in text.splitlines())
        chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
        clean_text = '\n'.join(chunk for chunk in chunks if chunk)
        
        # Check if we got useful content
        if len(clean_text) < 500 and len(json_ld_text) < 100:
             return (
                 f"WARNING: The scraped content is very short ({len(clean_text)} chars).\n"
                 "This website likely relies on JavaScript to display the menu (e.g., Toast, DoorDash, UberEats).\n"
                 "The current simple scraper cannot see this dynamic content.\n\n"
                 "*** WORKAROUND ***\n"
                 "Please open the website in your browser, Right-Click -> 'Print', select 'Save as PDF',\n"
                 "and upload the PDF to this app instead."
             )

        # Combine
        full_output = f"{clean_text}\n\n{json_ld_text}"
        return full_output

    except requests.exceptions.HTTPError as e:
        return f"HTTP Error {e.response.status_code}: {e}"
    except Exception as e:
        print(f"Scraping error: {e}")
        return f"Error scraping URL: {str(e)}"

