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
    Fetch and parse text from a URL.
    Returns the visible text content.
    """
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Remove scripts and styles
        for script in soup(["script", "style"]):
            script.decompose()
            
        text = soup.get_text()
        
        # Breakdown into lines and remove leading and trailing space on each
        lines = (line.strip() for line in text.splitlines())
        # Break multi-headlines into a line each
        chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
        # Drop blank lines
        text = '\n'.join(chunk for chunk in chunks if chunk)
        
        return text
    except Exception as e:
        print(f"Scraping error: {e}")
        return f"Error scraping URL: {str(e)}"

