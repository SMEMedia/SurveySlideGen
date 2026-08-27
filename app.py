from __future__ import annotations

from pathlib import Path

import streamlit as st

from generator import PURCHASE_PROFILES, discover_metric_options, generate_presentation, metric_value, read_rows


st.set_page_config(page_title="SME Audience Slide Generator", page_icon="📊", layout="wide")

st.markdown(
    """
    <style>
    .block-container {max-width: 1120px; padding-top: 2.2rem;}
    .hero {padding: 1.4rem 1.6rem; border-radius: 16px; background: linear-gradient(120deg,#073B5C,#0B76A8); color:white; margin-bottom:1.3rem;}
    .hero h1 {margin:0; font-size:2.25rem;}
    .hero p {margin:.45rem 0 0; opacity:.9;}
    [data-testid="stMetric"] {background:#F2F8FB; border:1px solid #D5EAF4; padding:12px; border-radius:12px;}
    </style>
    <div class="hero"><h1>SME Audience Slide Generator</h1><p>Create a polished two-slide advertiser presentation from the audience-insights workbook.</p></div>
    """,
    unsafe_allow_html=True,
)

uploaded = st.file_uploader("1. Upload the audience-insights workbook", type=["xlsx"], help="Use the final frequency-report workbook.")

if uploaded is None:
    st.info("Upload an `.xlsx` research report to begin. Reach will always be created as **X** for manual HubSpot entry.")
    st.stop()

try:
    rows = read_rows(uploaded.getvalue())
    qualified_options = discover_metric_options(rows, "Which of the following best describes your job title or primary function?")
    industry_options = discover_metric_options(rows, "In which industry or markets is your organization primarily involved?")
    decision_options = discover_metric_options(rows, "What is your role in purchasing decisions?")
    fixed_decision = metric_value(decision_options, "purchase influence net", "Total")
except Exception as exc:
    st.error(f"This workbook could not be read: {exc}")
    st.stop()

st.success("Workbook loaded successfully.")

left, right = st.columns(2)
with left:
    company_name = st.text_input("2. Advertiser/company name", placeholder="e.g., Kennametal")
with right:
    report_year = st.number_input("Research report year", min_value=2024, max_value=2100, value=2026, step=1)

st.subheader("3. Define the target audience")
qualified_tab, industry_tab, purchase_tab = st.tabs(["Qualified audience", "Industry audience", "Purchase focus"])

qualified_nets = [option.label for option in qualified_options if option.label.lower() in {"leadership + purchasing", "advantive", "ssab"} or "net" in option.label.lower()]
if not qualified_nets:
    qualified_nets = [option.label for option in qualified_options]

with qualified_tab:
    q_mode = st.radio("Qualified metric source", ["Use a workbook net", "Enter a verified percentage"], horizontal=True)
    if q_mode == "Use a workbook net":
        qualified_label = st.selectbox("Workbook net", qualified_nets, index=qualified_nets.index("leadership + purchasing") if "leadership + purchasing" in qualified_nets else 0)
        qualified_column = st.selectbox("Audience column", ["Total", "Total Magazine Audience", "Media Cross-Platform", "Digtial Non-Mag"])
        qualified_value = metric_value(qualified_options, qualified_label, qualified_column)
    else:
        qualified_value = st.number_input("Verified qualified percentage", min_value=0, max_value=100, value=43, step=1) / 100
    qualified_denominator = st.selectbox("Headline ratio", [5, 10], format_func=lambda value: f"Nearest ‘in {value}’ ratio")
    qualified_description = st.text_area(
        "Qualified-audience description",
        "Hold Manufacturing Engineering, Production, or C-Suite Leadership Roles at their Organization",
        max_chars=180,
    )
    st.metric("Qualified audience", f"{qualified_value:.0%}")

industry_nets = [option.label for option in industry_options if option.label.lower() == "ssab" or "net" in option.label.lower()]
if not industry_nets:
    industry_nets = [option.label for option in industry_options]

with industry_tab:
    i_mode = st.radio("Industry metric source", ["Use a workbook net", "Enter a verified percentage"], horizontal=True)
    if i_mode == "Use a workbook net":
        industry_label = st.selectbox("Workbook industry net", industry_nets, index=0)
        industry_column = st.selectbox("Industry audience column", ["Total", "Total Magazine Audience", "Media Cross-Platform", "Digtial Non-Mag"])
        industry_value = metric_value(industry_options, industry_label, industry_column)
    else:
        industry_value = st.number_input("Verified industry percentage", min_value=0, max_value=100, value=75, step=1) / 100
    industry_description = st.text_area(
        "Industry-audience description",
        "Work across Aerospace & Defense, Automotive, Industrial Machinery & Equipment, Medical Device, & Job Shops",
        max_chars=180,
    )
    st.metric("Industry audience", f"{industry_value:.0%}")

with purchase_tab:
    purchase_profile = st.selectbox("What does this advertiser sell?", list(PURCHASE_PROFILES))
    st.caption("This selection controls the slide-two headline and platform chart. The purchase-category comparison remains consistent.")

st.subheader("4. Review fixed fields")
fixed_a, fixed_b = st.columns(2)
fixed_a.metric("Reach", "X", help="Filled manually from HubSpot after download.")
fixed_b.metric("Decision-Makers", f"{fixed_decision:.0%}", help="Always uses purchase influence net / Total.")

ready = bool(company_name.strip() and qualified_description.strip() and industry_description.strip())
if not ready:
    st.warning("Enter the company name and both audience descriptions to enable generation.")

if st.button("Generate PowerPoint", type="primary", disabled=not ready, use_container_width=True):
    try:
        template = (Path(__file__).parent / "assets" / "sme_audience_template.pptx").read_bytes()
        result = generate_presentation(
            template_bytes=template,
            rows=rows,
            company_name=company_name.strip(),
            report_year=int(report_year),
            qualified_value=qualified_value,
            qualified_description=qualified_description.strip(),
            qualified_denominator=int(qualified_denominator),
            industry_value=industry_value,
            industry_description=industry_description.strip(),
            purchase_profile_name=purchase_profile,
        )
        filename = f"SME Media Reaches {company_name.strip()} Target Audience.pptx"
        st.session_state["generated_deck"] = result
        st.session_state["generated_filename"] = filename
        st.success("Your two-slide presentation is ready.")
    except Exception as exc:
        st.error(f"The presentation could not be generated: {exc}")

if "generated_deck" in st.session_state:
    st.download_button(
        "Download PowerPoint",
        data=st.session_state["generated_deck"],
        file_name=st.session_state["generated_filename"],
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary",
        use_container_width=True,
    )

st.caption("Research metrics come from the uploaded workbook. Reach is intentionally left as X for manual HubSpot entry.")
