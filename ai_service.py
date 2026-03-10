import google.generativeai as genai
import json

class AIService:
    """
    Handles interactions with AI providers (Gemini, OpenAI).
    """
    
    SYSTEM_PROMPT = """
    You are a specialized data extraction assistant for a POS system.
    Your task is to analyze a restaurant menu (image or text) and extract structured data.
    
    You must output strictly VALID JSON in the following format:
    {
        "items": [
            {
                "number": 100, 
                "name": "Classic Burger", 
                "price": 12.50,
                "description": "Juicy beef patty",
                "modifiers": ["Burger Toppings", "Sides"]
            }
        ],
        "submenus": [
            {
                "number": 200,
                "name": "Burgers", 
                "items": ["Classic Burger", "Cheeseburger"]
            }
        ],
        "modifier_groups": [
            {
                "number": 10000,
                "name": "Burger Toppings",
                "items": [
                    {"name": "Cheese", "price": 1.00},
                    {"name": "Bacon", "price": 2.00}
                ]
            }
        ]
    }
    
    RULES:
    1. EXPLICIT DATA ONLY. Do not infer or guess any values (like categories). Leave blank if not found.
    2. Extract all visible ITEMS with their prices.
    3. Treat menu sections (e.g. "Appetizers", "Lunch") as SUBMENUS.
    4. For MODIFIER GROUPS: Extract the Title/Name AND all individual modifier items/options within it.
    5. For MODIFIER ITEMS: Extract the Name and Helper Price (if any).
    6. Provide realistic numbers starting from 100 for items, 200 for submenus, 10000 for mod groups.
    7. Return ONLY the JSON. No markdown formatting.
    """

    def analyze_file(self, file_path: str, mime_type: str, api_key: str, provider: str = "gemini", model_name: str = "gemini-2.5-flash") -> dict:
        """
        Analyze a file (Image or PDF) and return structured data.
        """
        return self._call_gemini_file(file_path, mime_type, api_key, model_name)

    def analyze_text(self, text: str, api_key: str, provider: str = "gemini", model_name: str = "gemini-2.5-flash") -> dict:
        """
        Analyze text content (from scraping).
        """
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(model_name)
        inputs = [self.SYSTEM_PROMPT, f"Analyze the following menu text:\n\n{text}"]
        return self._generate_with_retry(model, inputs)

    def _call_gemini_file(self, file_path: str, mime_type: str, api_key: str, model_name: str) -> dict:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(model_name)
        
        uploaded_file = genai.upload_file(file_path, mime_type=mime_type)
        inputs = [self.SYSTEM_PROMPT, uploaded_file]
        return self._generate_with_retry(model, inputs)

    def _generate_with_retry(self, model, inputs, max_retries=3):
        import time
        import random
        
        for attempt in range(max_retries):
            try:
                response = model.generate_content(inputs)
                text = response.text
                
                if "```json" in text:
                    text = text.split("```json")[1].split("```")[0]
                elif "```" in text:
                    text = text.split("```")[1].split("```")[0]
                    
                return json.loads(text)
            except Exception as e:
                error_str = str(e)
                if "429" in error_str or "quota" in error_str.lower():
                    if attempt < max_retries - 1:
                        sleep_time = (2 ** attempt) + random.uniform(0, 1)
                        print(f"Rate limited. Retrying in {sleep_time:.2f}s...")
                        time.sleep(sleep_time)
                        continue
                print(f"Gemini Error (Attempt {attempt+1}): {e}")
                if attempt == max_retries - 1:
                    return {"error": f"Failed after {max_retries} attempts: {str(e)}"}
