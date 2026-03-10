import streamlit as st
import os
from ai_service import AIService
from excel_builder import ExcelBuilder
import utils

st.set_page_config(page_title="CoreDB Extract", layout="wide")

st.title("🍽️ CoreDB Extract: Menu to Excel")
st.markdown("Extract menu data from images or PDFs and generate import-ready Excel files for Aloha.")

# Sidebar - Configuration
st.sidebar.header("Configuration")

# Provider Options
provider_options = [
    "Google Gemini 2.5 Flash",
    "Google Gemini 2.5 Pro",
    "Google Gemini 3 Flash",
    "Google Gemini 3.1 Pro",
]

provider_selection = st.sidebar.selectbox("AI Model", provider_options)
api_key = st.sidebar.text_input("API Key", type="password", help="Enter API Key for the selected provider")

# Map selection to (provider_code, model_name)
config_map = {
    "Google Gemini 2.5 Flash": ("gemini", "gemini-2.5-flash"),
    "Google Gemini 2.5 Pro": ("gemini", "gemini-2.5-pro"),
    "Google Gemini 3 Flash": ("gemini", "gemini-3-flash-preview"),
    "Google Gemini 3.1 Pro": ("gemini", "gemini-3.1-pro-preview"),
}

# Pricing Data (Per 1 Million Tokens)
# Source: Google AI Studio Pricing (Approximate public rates)
MODEL_PRICING = {
    "gemini-2.5-flash": {
        "input_price": 0.075,
        "output_price": 0.30,
        "rpm": "15 Requests Per Minute (Free) / 1000 pay-as-you-go",
        "limit": "1M Context"
    },
    "gemini-2.5-pro": {
        "input_price": 3.50,
        "output_price": 10.50,
        "rpm": "2 Requests Per Minute (Free) / 360 pay-as-you-go",
        "limit": "2M Context"
    },
    "gemini-3-flash-preview": {
        "input_price": 0.00,
        "output_price": 0.00,
        "rpm": "Preview - Rate limits apply",
        "limit": "1M Context"
    },
    "gemini-3.1-pro-preview": {
        "input_price": 0.00,
        "output_price": 0.00,
        "rpm": "Preview - Rate limits apply",
        "limit": "2M Context"
    },
}

def estimate_cost(uploaded_file, model_key):
    """
    Estimates cost based on file type and model rates.
    Returns (estimated_cost, rpm_info)
    """
    pricing_key = model_key

    if not uploaded_file or pricing_key not in MODEL_PRICING:
        return 0.0, "N/A"

    pricing = MODEL_PRICING[pricing_key]

    # 1. Estimate Input Tokens
    input_tokens = 0
    if uploaded_file.type == "application/pdf":
        try:
            # Create a BytesIO object from the uploaded file to pass to PdfReader
            from io import BytesIO
            pdf_file_obj = BytesIO(uploaded_file.getvalue())
            pdf_reader = PyPDF2.PdfReader(pdf_file_obj)
            num_pages = len(pdf_reader.pages)
            # PDF to Image conversion approx: 258 tokens per page (Gemini standard image input) + Text overhead
            # Safe estimate: 1000 tokens per page (Text + Image overhead)
            input_tokens = num_pages * 1000
            # No need to reset pointer for uploaded_file itself, as we used a BytesIO copy
        except Exception:
            input_tokens = 5000 # Fallback for unreadable PDFs or errors
    else:
        # Image (PNG, JPG, JPEG)
        input_tokens = 258 # Gemini standard image token cost for a single image

    # 2. Estimate Output Tokens (JSON)
    # Menu extraction is verbose. Expect ~1000-2000 tokens per page/image.
    # Use a heuristic: output tokens are roughly 2x input tokens, with a minimum.
    max_output_tokens = max(input_tokens * 2, 4000)

    # 3. Calculate Cost (Price is per 1M tokens)
    in_cost = (input_tokens / 1_000_000) * pricing["input_price"]
    out_cost = (max_output_tokens / 1_000_000) * pricing["output_price"]

    total_max = in_cost + out_cost
    # Round UP to safe estimate, ensure it's not zero if prices are non-zero
    total_max = max(total_max, 0.0001) if pricing["input_price"] > 0 or pricing["output_price"] > 0 else 0.0

    return total_max, pricing["rpm"]

# Model Information for UI
MODEL_INFO = {
    "Google Gemini 2.5 Flash": {
        "desc": "Fast, efficient multimodal model for general-purpose tasks.",
        "price": "Free Tier: 15 RPM, 1,500 RPD. Paid: $0.075 / 1M input tokens.",
        "strength": "Balanced speed and cost. Excellent for standard menu extraction.",
        "limit": "1M Context. Good for most menus.",
        "url": "https://ai.google.dev/gemini-api/docs/models"
    },
    "Google Gemini 2.5 Pro": {
        "desc": "High-reasoning model designed for complex tasks and large document analysis.",
        "price": "Free Tier: 2 RPM, 50 RPD. Paid: $3.50 / 1M input tokens.",
        "strength": "Superior reasoning for complex modifier logic and messy handwritten menus.",
        "limit": "Slower analysis. Lower RPM limits in Free tier. 2M Context.",
        "url": "https://ai.google.dev/gemini-api/docs/models"
    },
    "Google Gemini 3 Flash": {
        "desc": "Next-gen multimodal model with strong coding and state-of-the-art reasoning.",
        "price": "Preview - Free while in preview.",
        "strength": "Best for complex multimodal understanding and agentic tasks.",
        "limit": "Preview stability. 1M Context.",
        "url": "https://ai.google.dev/gemini-api/docs/models"
    },
    "Google Gemini 3.1 Pro": {
        "desc": "Latest reasoning-first model optimized for complex agentic workflows.",
        "price": "Preview - Free while in preview.",
        "strength": "Cutting-edge reasoning and coding capabilities.",
        "limit": "Preview stability. 2M Context.",
        "url": "https://ai.google.dev/gemini-api/docs/models"
    },
}

selected_provider_code, selected_model_name = config_map.get(provider_selection, ("gemini", "gemini-2.5-flash"))

# Display Model Info in Sidebar
st.sidebar.markdown("---")
st.sidebar.subheader("Model Details")
# Fallback info
info = MODEL_INFO.get(provider_selection, {
    "desc": "High-performance AI model.",
    "price": "Check provider official pricing page.",
    "strength": "General extraction tasks.",
    "limit": "Standard limits apply.",
    "url": "https://ai.google.dev/models"
})

st.sidebar.info(f"**{provider_selection}**\n\n"
                f"📝 **Description:** {info['desc']}\n\n"
                f"💰 **Pricing & Limits:** {info['price']}\n\n"
                f"💪 **Strengths:** {info['strength']}\n\n"
                f"⚠️ **Known Limitations:** {info['limit']}\n\n"
                f"[Official Documentation]({info['url']})")

# Main Area
st.info("ℹ️ **App Overview:** This tool populates the **Item**, **Submenu**, **SubmenuItem**, and **ModifierGroup_Items** tabs.\n\n"
        "Please strictly refer to the **Instructions** tab in the downloaded Excel file for logic details.")

# Clean Template Download
with st.expander("📄 Need a blank template?"):
    st.write("Download a clean, protected template with usage notes.")
    if st.button("Generate Empty Template"):
        try:
            builder = ExcelBuilder()
            empty_bytes = builder.build_empty_template()
            st.download_button(
                label="📥 Download Empty Template",
                data=empty_bytes,
                file_name="Aloha_Import_Template_Empty.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Could not generate template: {e}")

tab_upload, tab_url = st.tabs(["📂 File Upload", "🌐 URL (Coming Soon)"])

with tab_upload:
    uploaded_file = st.file_uploader("Upload Menu (PDF, PNG, JPG)", type=["pdf", "png", "jpg", "jpeg"])

    if uploaded_file and api_key:

        # Display Cost Estimation
        est_cost, rpm_info = estimate_cost(uploaded_file, selected_model_name)
        
        st.info(f"""
        **📊 Estimation (Safe Upper Bound)**
        - **Model**: {selected_model_name}
        - **Est. Cost**: < ${est_cost:.4f} USD
        - **RPM Limit**: {rpm_info}
        
        *Note: 1 Request = Processing 1 Uploaded File.*
        """)

        if st.button("🚀 Extract Menu Data", type="primary"):
            with st.status("Processing...", expanded=True) as status:
                try:
                    # 1. Save file
                    st.write("Saving uploaded file...")
                    file_path = utils.save_uploaded_file(uploaded_file)
                    if not file_path:
                        st.error("Failed to save file.")
                        st.stop()
                    
                    # Determine Mime Type
                    mime_type = "application/pdf" if uploaded_file.name.lower().endswith(".pdf") else "image/jpeg"
                    if uploaded_file.name.lower().endswith(".png"):
                        mime_type = "image/png"
                        
                    # 2. Call AI
                    st.write(f"Analyzing with {provider_selection}...")
                    ai_service = AIService()
                    
                    data = ai_service.analyze_file(file_path, mime_type, api_key, provider=selected_provider_code, model_name=selected_model_name)
                    
                    if "error" in data:
                        st.error(f"AI Error: {data['error']}")
                        status.update(label="Failed", state="error")
                    else:
                        st.write("Parsing data...")
                        # 3. Build Excel
                        builder = ExcelBuilder()
                        builder.add_data(data)
                        excel_data = builder.build_excel()
                        
                        st.success("Extraction Complete!")
                        status.update(label="Complete!", state="complete")
                        
                        # 4. Download
                        st.download_button(
                            label="📥 Download Excel File",
                            data=excel_data,
                            file_name="Aloha_Import_Ready.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        # Preview Data (Optional)
                        with st.expander("Preview Extracted Data"):
                            st.json(data)

                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())
                finally:
                    # Cleanup
                    if 'file_path' in locals():
                        utils.cleanup_temp_file(file_path)
    
    elif not api_key:
        st.warning("Please enter your API Key in the sidebar to proceed.")

with tab_url:
    url_input = st.text_input("Enter Menu URL", placeholder="https://example.com/menu")
    
    if url_input and api_key:
        if st.button("🌐 Scrape & Extract Menu", type="primary"):
            with st.status("Processing URL...", expanded=True) as status:
                try:
                    # 1. Scrape
                    st.write("Scraping website...")
                    scraped_text = utils.scrape_url(url_input)
                    
                    if "Error scraping URL" in scraped_text:
                        st.error(scraped_text)
                        st.stop()
                        
                    st.success("Scraping successful! (Preview first 500 chars)")
                    st.caption(scraped_text[:500] + "...")
                        
                    # 2. Call AI
                    st.write(f"Analyzing text with {provider_selection}...")
                    ai_service = AIService()
                    
                    data = ai_service.analyze_text(scraped_text, api_key, provider=selected_provider_code, model_name=selected_model_name)
                    
                    if "error" in data:
                        st.error(f"AI Error: {data['error']}")
                        status.update(label="Failed", state="error")
                    else:
                        st.write("Parsing data...")
                        builder = ExcelBuilder()
                        builder.add_data(data)
                        excel_data = builder.build_excel()
                        
                        st.success("Extraction Complete!")
                        status.update(label="Complete!", state="complete")
                        
                        st.download_button(
                            label="📥 Download Excel File",
                            data=excel_data,
                            file_name="Aloha_Import_Ready_From_URL.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        with st.expander("Preview Extracted Data"):
                            st.json(data)
                            
                except Exception as e:
                     st.error(f"An error occurred: {str(e)}")
    elif not api_key:
        st.warning("Please enter your API Key in the sidebar to proceed.")

