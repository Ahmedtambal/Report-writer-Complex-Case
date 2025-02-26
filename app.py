import json
import streamlit as st
import os
from logic import (
    create_new_document,
    save_uploaded_file,
    extract_text_from_pdf,
    extract_text_from_image,
    generate_multi_risk_attitude_text,
    extract_plan_details_with_gpt,
    generate_pension_review_section,
    generate_safe_withdrawal_rate_section,
    extract_last_year_performance_text,
    extract_fund_performance_with_gpt,
    extract_dark_star_performance_with_gpt,
    extract_investment_portfolio_with_gpt,
    add_investment_holdings_tables,
    extract_sap_comparison_with_gpt,
    extract_annuity_quotes_with_gpt,
    extract_fund_comparison_with_gpt,
    generate_iht_section,
    extract_text_from_file,    
    process_plan_report,
    process_funds_for_comparison,
    generate_safe_withdrawal_rate_sections
)
import openai

# Define folders for uploaded and generated documents
UPLOAD_FOLDER = "uploaded_docs"
OUTPUT_FOLDER = "generated_docs"

# Streamlit Page Configuration
st.set_page_config(page_title="Zomi AI Persona", page_icon="💼", layout="wide")

# Inject CSS for Dark Mode Styling
st.markdown(
    """
    <style>
    /* Your CSS styles here */
    </style>
    """,
    unsafe_allow_html=True,
)

# Title Section
st.markdown('<div class="title">Zomi Wealth AI</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Generate personalized financial reports with ease.</div>', unsafe_allow_html=True)

# File Upload Section
st.markdown('<div class="upload-section">', unsafe_allow_html=True)

# Essential uploads: Template, FactFind, Risk Profiles
uploaded_template = st.file_uploader("📄 Upload Report Template (.docx)", type="docx")
uploaded_factfind = st.file_uploader("📄 Upload FactFind Document (.pdf)", type="pdf")
uploaded_risk_profiles = st.file_uploader(
    "Upload Risk Profile Image(s) or PDF(s)",
    type=["png", "jpg", "jpeg", "pdf"],
    accept_multiple_files=True
)

# Other uploads
uploaded_files = st.file_uploader(
    "Upload Plan Files (docx, pdf, png, jpg, jpeg)",
    type=["docx", "pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=True
)
uploaded_fund_fact_sheets = st.file_uploader("📄 Upload Client Fund Fact Sheets (.pdf)", type="pdf", accept_multiple_files=True)
uploaded_dark_star_fact_sheet = st.file_uploader("📄 Upload Dark Star Fact Sheet (.pdf)", type="pdf", accept_multiple_files=True)
uploaded_sap_report = st.file_uploader("📄 Upload SAP Report File (.pdf)", type="pdf", accept_multiple_files=True)
annuity_files = st.file_uploader("📤 Upload Annuity Quotes Image(s)", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

# Fund Comparison Files
st.markdown("### Fund Comparison Files")
st.markdown("Upload the necessary files for the fund comparison table.")
num_funds = st.number_input(
    "How many funds do you want to compare?",
    min_value=1,
    max_value=40,
    value=1,
    step=1,
    help="Select the number of funds (excluding P1) you wish to compare."
)
st.markdown("### Upload Files for Each Fund")
funds_uploads = []
for i in range(num_funds):
    uploaded_fund_files = st.file_uploader(
        f"Upload files for Fund {i+1} (PDF only)",
        type=["pdf"],
        accept_multiple_files=True,
        key=f"fund_{i+1}"
    )
    funds_uploads.append(uploaded_fund_files)
st.markdown("### Upload Files for P1 (Benchmark Fund)")
p1_files = st.file_uploader(
    "Upload files for P1 Fund (PDF only)",
    type=["pdf"],
    accept_multiple_files=True,
    key="p1_files"
)
st.markdown('</div>', unsafe_allow_html=True)

# ----- Processing: Only proceed if essential files are provided -----
if uploaded_template and uploaded_factfind and uploaded_risk_profiles:
    # Process Template and FactFind
    template_path = save_uploaded_file(uploaded_template, UPLOAD_FOLDER)
    factfind_path = save_uploaded_file(uploaded_factfind, UPLOAD_FOLDER)
    factfinding_text = extract_text_from_pdf(factfind_path)
    
    # Initialize variables for later use
    plan_report_data = []
    plan_report_text = ""
    fund_performance_data = []
    dark_star_performance_data = []
    plan_texts_list = []
    
    # Process Risk Profiles
    risk_texts = []
    for risk_file in uploaded_risk_profiles:
        risk_file_path = save_uploaded_file(risk_file, UPLOAD_FOLDER)
        if risk_file.name.lower().endswith(".pdf"):
            extracted_risk_text = extract_text_from_pdf(risk_file_path)
        else:
            extracted_risk_text = extract_text_from_image(risk_file_path)
        if extracted_risk_text.strip():
            risk_texts.append(extracted_risk_text)
            st.success(f"Extracted risk text from '{risk_file.name}'")
        else:
            st.warning(f"No text found in '{risk_file.name}', skipping risk parsing.")
    
    if risk_texts:
        try:
            final_attitude_text = generate_multi_risk_attitude_text(risk_texts)
        except Exception as e:
            st.error(f"Error generating final risk text: {e}")
            final_attitude_text = "No risk details provided."
    else:
        final_attitude_text = "No risk details provided."
    
    # Process Plan Reports
    plan_review_paragraphs = []
    if uploaded_files:
        for file in uploaded_files:
            file_path = save_uploaded_file(file, UPLOAD_FOLDER)
            extracted_text = extract_text_from_file(file_path)
            if extracted_text.strip():
                try:
                    plan_details = extract_plan_details_with_gpt(extracted_text)
                    plan_report_data.extend(plan_details)
                    plan_report_text += extracted_text + "\n"
                    plan_texts_list.append(extracted_text)
                    review_paragraph = generate_pension_review_section(extracted_text)
                    plan_review_paragraphs.append(review_paragraph)
                    st.success(f"Generated a pension review for '{file.name}'")
                except Exception as e:
                    st.error(f"Error processing '{file.name}': {e}")
            else:
                st.warning(f"No text found in '{file.name}', skipping review generation.")
    
    product_report_text = plan_report_text  # Modify as needed
    
    # Process Client Fund Fact Sheets (Multi-file)
    if uploaded_fund_fact_sheets:
        fund_performance_data = []
        last_year_performance_text = "No last-year performance found."
        if isinstance(uploaded_fund_fact_sheets, list):
            all_extracted_texts = []
            for file in uploaded_fund_fact_sheets:
                file_path = save_uploaded_file(file, UPLOAD_FOLDER)
                text = extract_text_from_pdf(file_path)
                if text.strip():
                    all_extracted_texts.append(text)
                else:
                    st.warning(f"No text found in {file.name}.")
            if all_extracted_texts:
                combined_text = "\n".join(all_extracted_texts)
                fund_performance_data = extract_fund_performance_with_gpt(combined_text)
                last_year_performance_text = extract_last_year_performance_text(combined_text)
            else:
                st.warning("No fund text could be extracted from the uploaded files.")
        else:
            fund_fact_sheet_path = save_uploaded_file(uploaded_fund_fact_sheets, UPLOAD_FOLDER)
            extracted_fund_text = extract_text_from_pdf(fund_fact_sheet_path)
            if extracted_fund_text.strip():
                fund_performance_data = extract_fund_performance_with_gpt(extracted_fund_text)
                last_year_performance_text = extract_last_year_performance_text(extracted_fund_text)
            else:
                st.warning("No text found in the uploaded fund fact sheet.")
    else:
        last_year_performance_text = "No last-year performance found."
    
    # Process Dark Star Fact Sheets (Multi-file)
    if uploaded_dark_star_fact_sheet:
        dark_star_performance_data = []
        if isinstance(uploaded_dark_star_fact_sheet, list):
            all_extracted_dark_star_texts = []
            for file in uploaded_dark_star_fact_sheet:
                file_path = save_uploaded_file(file, UPLOAD_FOLDER)
                text = extract_text_from_pdf(file_path)
                if text.strip():
                    all_extracted_dark_star_texts.append(text)
                else:
                    st.warning(f"No text found in {file.name}.")
            if all_extracted_dark_star_texts:
                combined_dark_star_text = "\n".join(all_extracted_dark_star_texts)
                dark_star_performance_data = extract_dark_star_performance_with_gpt(combined_dark_star_text)
            else:
                st.warning("No text extracted from the uploaded Dark Star fact sheets.")
        else:
            file_path = save_uploaded_file(uploaded_dark_star_fact_sheet, UPLOAD_FOLDER)
            text = extract_text_from_pdf(file_path)
            if text.strip():
                dark_star_performance_data = extract_dark_star_performance_with_gpt(text)
            else:
                st.warning("No text found in the uploaded Dark Star fact sheet.")
    
    # Process SAP Reports
    sap_comparison_tables = []
    if uploaded_sap_report:
        for sap_file in uploaded_sap_report:
            sap_report_path = save_uploaded_file(sap_file, UPLOAD_FOLDER)
            extracted_sap_text = extract_text_from_pdf(sap_report_path)
            if extracted_sap_text.strip():
                try:
                    comparison = extract_sap_comparison_with_gpt(extracted_sap_text)
                    sap_comparison_tables.append(comparison)
                except Exception as e:
                    st.error(f"Error processing SAP report '{sap_file.name}': {e}")
            else:
                st.warning(f"No text found in '{sap_file.name}', skipping SAP report processing.")
    
    # Process Annuity Quotes
    annuity_quotes_text = ""
    if annuity_files:
        annuity_generated = []
        for annuity_file in annuity_files:
            annuity_path = save_uploaded_file(annuity_file, UPLOAD_FOLDER)
            annuity_extracted = extract_text_from_file(annuity_path)
            if annuity_extracted.strip():
                try:
                    generated = extract_annuity_quotes_with_gpt(annuity_extracted)
                    annuity_generated.append(generated)
                except Exception as e:
                    st.error(f"Error processing annuity file '{annuity_file.name}': {e}")
                    annuity_generated.append("Error processing file.")
            else:
                annuity_generated.append("No text extracted from file.")
        annuity_quotes_text = "\n".join(annuity_generated)
    
    # Process Fund Comparisons
    fund_comparison_results = []
    try:
        fund_comparison_results = process_funds_for_comparison(funds_uploads, p1_files)
    except Exception as e:
        st.error(f"Error processing fund comparisons: {e}")
    combined_fund_comparison_text = "\n\n".join(
        [f"Fund {num}: {text}" for num, text in fund_comparison_results]
    )
    
    # Process Portfolio Extraction
    portfolio_jsons = []
    if uploaded_files:
        for file in uploaded_files:
            file_path = save_uploaded_file(file, UPLOAD_FOLDER)
            extracted_text = extract_text_from_file(file_path)
            if extracted_text.strip():
                try:
                    pj = extract_investment_portfolio_with_gpt(extracted_text)
                    portfolio_jsons.append(pj)
                    st.write(f"Portfolio JSON for {file.name}:", pj)
                except Exception as e:
                    st.error(f"Error extracting portfolio from '{file.name}': {e}")
    
    # Generate IHT Section (only if FactFind and Plan Reports were provided)
    iht_text = ""
    if factfinding_text and plan_texts_list:
        try:
            iht_text = generate_iht_section(factfinding_text, plan_texts_list)
        except Exception as e:
            st.error("Error generating IHT section: " + repr(e))
    else:
        st.warning("Please upload the FactFind and Plan Report files to extract IHT details.")
    
    # Generate Safe Withdrawal Rate Sections
    swr_sections_list = generate_safe_withdrawal_rate_sections(plan_texts_list)
    combined_swr_text = "\n\n".join(
        [f"Safe Withdrawal Rate for File {idx+1}:\n{swr}" for idx, swr in enumerate(swr_sections_list)]
    )
    st.write("Combined SWR sections:", combined_swr_text)
    
    # Create final output document
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    output_path = os.path.join(OUTPUT_FOLDER, "Generated_Report.docx")
    
    if st.button("Generate Report", key="generate_button"):
        try:
            st.markdown('<div style="text-align:center;">🛠️ Generating your personalized report...</div>', unsafe_allow_html=True)
            create_new_document(
                template_path=template_path,
                factfinding_text=factfinding_text,
                attitude_to_risk=final_attitude_text,
                table_data=plan_report_data,
                product_report_text=product_report_text,
                plan_review_texts=plan_review_paragraphs,
                plan_review_paragraphs=plan_review_paragraphs,
                plan_report_text=plan_report_text,
                fund_performance_data=fund_performance_data,
                last_year_performance_text=last_year_performance_text,
                dark_star_performance_data=dark_star_performance_data,
                sap_comparison_tables=sap_comparison_tables,
                annuity_quotes_text=annuity_quotes_text,
                fund_comparison_text=combined_fund_comparison_text,
                iht_text=iht_text,
                portfolio_json=portfolio_jsons,
                safe_withdrawal_text=combined_swr_text,
                output_path=output_path
            )
            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 Download Generated Report",
                    data=f,
                    file_name="Generated_Report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
        except Exception as e:
            st.error(f"❌ An error occurred: {e}")

else:
    st.error("Please upload the Report Template, FactFind Document, and Risk Profiles to generate a report.")

# Footer Section
st.markdown('<div class="footer">Working Hours: Monday to Friday, 9:00 AM – 5:30 PM</div>', unsafe_allow_html=True)
st.markdown(
    """
    <div class="disclaimer">
    Zomi Wealth is a trading name of Holistic Wealth Management Limited, authorized and regulated by the FCA.<br>
    Guidance provided is subject to the UK regulatory regime and is targeted at UK consumers.<br>
    Investments can go down as well as up. Past performance is not indicative of future results.
    </div>
    """,
    unsafe_allow_html=True,
)
