import google.generativeai as genai
import json
from typing import List, Optional
from PIL import Image

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
                "description": "Juicy beef patty"
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
                "number": 1000,
                "name": "Burger Toppings"
            }
        ]
    }
    
    RULES:
    1. EXPLICIT DATA ONLY. Do not infer or guess any values (like categories). Leave blank if not found.
    2. Extract all visible ITEMS.
    3. Treat menu sections (e.g. "Appetizers", "Lunch") as SUBMENUS, not categories.
    4. For MODIFIER GROUPS, extract ONLY the Title/Name and a generated Number. DO NOT extract the individual modifier items (like "Lettuce", "Tomato") inside the group.
    5. Provide realistic numbers starting from 100 for items, 200 for submenus, 10000 for mod groups.
    6. Return ONLY the JSON. No markdown formatting.
    """

    def analyze_file(self, file_path: str, mime_type: str, api_key: str, provider: str = "gemini", model_name: str = "gemini-2.0-flash") -> dict:
        """
        Analyze a file (Image or PDF) and return structured data.
        """
        if provider.lower() == "gemini":
            return self._call_gemini_file(file_path, mime_type, api_key, model_name)
        elif provider.lower() == "openai":
            return self._call_openai_file(file_path, mime_type, api_key)
        elif provider.lower() == "anthropic":
             return self._call_anthropic_file(file_path, mime_type, api_key)
        elif provider.lower() == "deepseek":
             # DeepSeek vision not typically supported same way via standard chat endpoint for file upload in all utils, 
             # but we can try if they support image URL or base64. 
             # For now, let's treat DeepSeek mainly for text unless specified.
             # Actually, DeepSeek doesn't support direct image uploads in their basic chat API yet usually.
             # We will raise error for file-based deepseek for now or fallback to text if valid.
             raise NotImplementedError("DeepSeek file analysis not fully supported yet.")
        else:
            raise NotImplementedError(f"Provider {provider} not supported for file analysis")

    def analyze_text(self, text: str, api_key: str, provider: str = "gemini", model_name: str = "gemini-2.0-flash") -> dict:
        """
        Analyze text content (from scraping).
        """
        if provider.lower() == "gemini":
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(model_name)
            inputs = [self.SYSTEM_PROMPT, f"Analyze the following menu text:\n\n{text}"]
            return self._generate_with_retry(model, inputs)
        elif provider.lower() == "openai":
            return self._call_openai_text(text, api_key)
        elif provider.lower() == "anthropic":
            return self._call_anthropic_text(text, api_key)
        elif provider.lower() == "deepseek":
            return self._call_deepseek_text(text, api_key)
        else:
             raise NotImplementedError(f"Provider {provider} not supported for text analysis")

    def _call_gemini_file(self, file_path: str, mime_type: str, api_key: str, model_name: str) -> dict:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(model_name)
        
        uploaded_file = genai.upload_file(file_path, mime_type=mime_type)
        inputs = [self.SYSTEM_PROMPT, uploaded_file]
        return self._generate_with_retry(model, inputs)

    def _call_openai_text(self, text: str, api_key: str) -> dict:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)
        
        try:
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": self.SYSTEM_PROMPT},
                    {"role": "user", "content": f"Analyze the following menu text:\n\n{text}"}
                ],
                response_format={"type": "json_object"}
            )
            return json.loads(response.choices[0].message.content)
        except Exception as e:
            return {"error": str(e)}

    def _call_anthropic_text(self, text: str, api_key: str) -> dict:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
        
        try:
            message = client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=4096,
                system=self.SYSTEM_PROMPT,
                messages=[
                    {"role": "user", "content": f"Analyze the following menu text:\n\n{text}"}
                ]
            )
            # Claude return text, might need json parsing
            content = message.content[0].text
            if "```json" in content:
                content = content.split("```json")[1].split("```")[0]
            elif "```" in content:
                content = content.split("```")[1].split("```")[0]
            return json.loads(content)
        except Exception as e:
             return {"error": str(e)}

    def _call_deepseek_text(self, text: str, api_key: str) -> dict:
        # DeepSeek is OpenAI compatible
        from openai import OpenAI
        client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
        
        try:
            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": self.SYSTEM_PROMPT},
                    {"role": "user", "content": f"Analyze the following menu text:\n\n{text}"}
                ],
                response_format={"type": "json_object"}
            )
            return json.loads(response.choices[0].message.content)
        except Exception as e:
            return {"error": str(e)}

    # Stub implementations for file support on other providers (Image encoding needed)
    def _call_openai_file(self, file_path: str, mime_type: str, api_key: str) -> dict:
        # Implement base64 encoding for image
        from openai import OpenAI
        import base64
        
        def encode_image(image_path):
            with open(image_path, "rb") as image_file:
                return base64.b64encode(image_file.read()).decode('utf-8')
        
        base64_image = encode_image(file_path)
        client = OpenAI(api_key=api_key)
        
        try:
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": self.SYSTEM_PROMPT},
                    {"role": "user", "content": [
                        {"type": "text", "text": "Analyze this menu image."},
                        {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{base64_image}"}}
                    ]}
                ],
                response_format={"type": "json_object"}
            )
            return json.loads(response.choices[0].message.content)
        except Exception as e:
            return {"error": str(e)}

    def _call_anthropic_file(self, file_path: str, mime_type: str, api_key: str) -> dict:
        import anthropic
        import base64
        
        with open(file_path, "rb") as image_file:
            image_data = base64.b64encode(image_file.read()).decode('utf-8')
            
        client = anthropic.Anthropic(api_key=api_key)
        
        try:
            message = client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=4096,
                system=self.SYSTEM_PROMPT,
                messages=[
                    {
                        "role": "user", 
                        "content": [
                            {"type": "image", "source": {"type": "base64", "media_type": mime_type, "data": image_data}},
                            {"type": "text", "text": "Analyze this menu image."}
                        ]
                    }
                ]
            )
            content = message.content[0].text
            if "```json" in content:
                content = content.split("```json")[1].split("```")[0]
            elif "```" in content:
                content = content.split("```")[1].split("```")[0]
            return json.loads(content)
        except Exception as e:
            return {"error": str(e)}

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
