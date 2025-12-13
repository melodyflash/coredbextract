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

# Provider Options including specific Gemini/Gemma models
provider_options = [
    "Google Gemini 2.0 Flash",
    "Google Gemini 2.0 Flash Lite",
    "Google Gemini 2.0 Flash Exp",
    "Google Gemini 2.5 Flash",
    "Google Gemini 2.5 Flash Lite",
    "Google Gemini 2.5 Pro",
    "Google Gemini 3.0 Pro",
    "Google Gemma 3 1B",
    "Google Gemma 3 12B",
    "Google Gemma 3 27B",
    "OpenAI (GPT-4o)",
    "Anthropic (Claude 3.5 Sonnet)",
    "DeepSeek (Chat Only)"
]

provider_selection = st.sidebar.selectbox("AI Model", provider_options)
api_key = st.sidebar.text_input("API Key", type="password", help="Enter API Key for the selected provider")

# Map selection to (provider_code, model_name)
# For non-Gemini, model_name is ignored by current service implementation but we follow structure
config_map = {
    "Google Gemini 2.0 Flash": ("gemini", "gemini-2.0-flash"),
    "Google Gemini 2.0 Flash Lite": ("gemini", "gemini-2.0-flash-lite"),
    "Google Gemini 2.0 Flash Exp": ("gemini", "gemini-2.0-flash-exp"),
    "Google Gemini 2.5 Flash": ("gemini", "gemini-2.5-flash"),
    "Google Gemini 2.5 Flash Lite": ("gemini", "gemini-2.5-flash-lite"),
    "Google Gemini 2.5 Pro": ("gemini", "gemini-2.5-pro"),
    "Google Gemini 3.0 Pro": ("gemini", "gemini-3-pro"),
    "Google Gemma 3 1B": ("gemini", "gemma-3-1b"),
    "Google Gemma 3 12B": ("gemini", "gemma-3-12b"),
    "Google Gemma 3 27B": ("gemini", "gemma-3-27b"),
    "OpenAI (GPT-4o)": ("openai", "gpt-4o"),
    "Anthropic (Claude 3.5 Sonnet)": ("anthropic", "claude-3-5-sonnet-20241022"),
    "DeepSeek (Chat Only)": ("deepseek", "deepseek-chat")
}

selected_provider_code, selected_model_name = config_map.get(provider_selection, ("gemini", "gemini-2.0-flash"))

# Main Area
tab_upload, tab_url = st.tabs(["📂 File Upload", "🌐 URL (Coming Soon)"])

with tab_upload:
    uploaded_file = st.file_uploader("Upload Menu (PDF, PNG, JPG)", type=["pdf", "png", "jpg", "jpeg"])

    if uploaded_file and api_key:
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
