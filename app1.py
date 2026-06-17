import streamlit as st
import google.generativeai as genai
import openpyxl
import json
import os
import tempfile
from io import BytesIO

# --- PAGE SETUP ---
st.set_page_config(page_title="Financial Report Automator", layout="centered")
st.title("📊 Automated Financial Parser")
st.write("Upload an Annual Report PDF, and the AI will automatically extract and map the data into your INDstocks Excel template.")

# --- API KEY CONFIGURATION ---
api_key = st.text_input("Enter your Google Gemini API Key:", type="password")
if api_key:
    genai.configure(api_key=api_key)

# --- UPLOADS ---
uploaded_pdf = st.file_uploader("1. Drop the Annual Report PDF here", type="pdf")
uploaded_excel = st.file_uploader("2. Drop your Master Excel Template here", type="xlsx")

target_year = st.selectbox("3. Select Fiscal Year for this PDF", ["2022", "2023", "2024", "2025"])

# Exact Column mapping from your INDstocks file
column_map = {"2022": 3, "2023": 4, "2024": 5, "2025": 6}
target_column = column_map[target_year]

if st.button("Process Data & Generate Excel"):
    if not api_key:
        st.error("Please enter your Gemini API Key.")
    elif not uploaded_pdf or not uploaded_excel:
        st.error("Please upload both the PDF and the Excel template.")
    else:
        with st.spinner("AI is reading the financial statements. This takes about 15-30 seconds..."):
            try:
                # 1. Save PDF temporarily so Gemini can read it
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                    temp_pdf.write(uploaded_pdf.read())
                    temp_pdf_path = temp_pdf.name

                # 2. Upload to Gemini
                sample_file = genai.upload_file(path=temp_pdf_path, display_name="Annual Report")
                model = genai.GenerativeModel('gemini-2.5-flash')
                
                # 3. Prompt the AI
                prompt = f"""
                You are an expert financial analyst. Read this Annual Report.
                Extract the financial figures for the fiscal year {target_year} and output them in STRICT JSON.
                Values MUST be in crores (₹). If a value is missing, use 0. No commas in numbers.
                
                Extract exactly these keys:
                {{
                    "PnL": {{
                        "Revenue from operations": 0.0,
                        "Other income": 0.0,
                        "Employee benefit expense": 0.0,
                        "Finance costs": 0.0,
                        "Depreciation and amortisation": 0.0,
                        "Other expenses": 0.0
                    }},
                    "BalanceSheet": {{
                        "Property, plant and equipment": 0.0,
                        "Trade receivables": 0.0,
                        "Cash and cash equivalents": 0.0,
                        "Total equity": 0.0,
                        "Trade payables": 0.0
                    }}
                }}
                """
                
                response = model.generate_content([sample_file, prompt])
                
                # 4. Clean JSON (Fixed the single-line syntax)
                json_str = response.text.strip().replace('```json', '').replace('```', '')
                parsed_data = json.loads(json_str)
                
                # 5. Inject into Excel
                wb = openpyxl.load_workbook(uploaded_excel)
                
                # Exact Row mapping for your P&L sheet
                ws_pnl = wb['P&L']
                pnl_mapping = {
                    "Revenue from operations": 6, 
                    "Other income": 7, 
                    "Employee benefit expense": 10, 
                    "Finance costs": 11, 
                    "Depreciation and amortisation": 12, 
                    "Other expenses": 13
                }
                for item, row in pnl_mapping.items():
                    if item in parsed_data.get('PnL', {}):
                        ws_pnl.cell(row=row, column=target_column).value = float(parsed_data['PnL'][item])

                # Exact Row mapping for your Balance Sheet
                ws_bs = wb['Balance Sheet']
                bs_mapping = {
                    "Property, plant and equipment": 7, 
                    "Trade receivables": 16, 
                    "Cash and cash equivalents": 18, 
                    "Total equity": 29, 
                    "Trade payables": 35
                }
                for item, row in bs_mapping.items():
                    if item in parsed_data.get('BalanceSheet', {}):
                        ws_bs.cell(row=row, column=target_column).value = float(parsed_data['BalanceSheet'][item])

                # 6. Prepare file for download
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                
                # Cleanup temp file
                os.remove(temp_pdf_path)
                genai.delete_file(sample_file.name)

                st.success("✅ Success! The data has been successfully mapped.")
                st.download_button(
                    label="⬇️ Download Updated Excel File",
                    data=output,
                    file_name=f"Updated_Financials_FY{target_year}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"An error occurred: {e}")