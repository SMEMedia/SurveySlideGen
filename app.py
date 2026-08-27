from __future__ import annotations

from pathlib import Path

import streamlit as st

from generator import discover_metric_options, generate_presentation, metric_value, read_rows


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
    <div class="hero"><h1>SME Audience Slide Generator</h1><p>Create a two-slide advertiser presentation using HubSpot and questions 10, 12, 13, and 14.</p></div>
    """,
    unsafe_allow_html=True,
)

uploaded = st.file_uploader("Upload the audience-insights workbook", type=["xlsx"], help="Use the final frequency-report workbook.")

if uploaded is None:
    st.info("Upload an `.xlsx` research report to begin.")
    st.stop()

try:
    rows = read_rows(uploaded.getvalue())
    purchase_options = discover_metric_options(rows, "Which types of products/services do you influence or purchase?")
    qualified_options = discover_metric_options(rows, "Which of the following best describes your job title or primary function?")
    industry_options = discover_metric_options(rows, "In which industry or markets is your organization primarily involved?")
    decision_options = discover_metric_options(rows, "What is your role in purchasing decisions?")
    fixed_decision = metric_value(decision_options, "purchase influence net", "Total")
except Exception as exc:
    st.error(f"This workbook could not be read: {exc}")
    st.stop()

st.success("Workbook loaded successfully.")

st.subheader("Client setup")
left, right = st.columns(2)
with left:
    company_name = st.text_input("Advertiser/company name", placeholder="e.g., Kennametal")
    reach = st.text_input(
        "HubSpot audience reach",
        placeholder="e.g., 35K+ or 12,500",
        help="Enter the audience-segment total pulled from HubSpot for slide 1.",
        max_chars=20,
    )
with right:
    report_year = st.number_input("Research report year", min_value=2024, max_value=2100, value=2026, step=1)

st.subheader("Slide 1 — Target audience")
st.caption("Choose the precomputed OR-net rows matching the advertiser's target. Slide 1 always uses the Total column.")

qualified_nets = [
    option.label
    for option in qualified_options
    if "net" in option.label.lower() or "+" in option.label or option.label.lower() in {"advantive", "ssab"}
]
if not qualified_nets:
    qualified_nets = [option.label for option in qualified_options]
industry_nets = [option.label for option in industry_options if "net" in option.label.lower() or option.label.lower() == "ssab"]
if not industry_nets:
    industry_nets = [option.label for option in industry_options]

target_left, target_right = st.columns(2)
with target_left:
    qualified_label = st.selectbox("Q14 — Job function OR-net", qualified_nets)
    qualified_value = metric_value(qualified_options, qualified_label, "Total")
    qualified_description = st.text_area(
        "Job-function description for the slide",
        "Hold Manufacturing Engineering, Production, or C-Suite Leadership Roles at their Organization",
        max_chars=180,
    )
    qualified_denominator = st.selectbox("Display the result as", [5, 10], format_func=lambda value: f"Nearest ‘in {value}’ ratio")
    st.metric("Qualified audience", f"{qualified_value:.0%}")

with target_right:
    industry_label = st.selectbox("Q13 — Industry OR-net", industry_nets)
    industry_value = metric_value(industry_options, industry_label, "Total")
    industry_description = st.text_area(
        "Industry description for the slide",
        "Work across Aerospace & Defense, Automotive, Industrial Machinery & Equipment, Medical Device, & Job Shops",
        max_chars=180,
    )
    st.metric("Industry audience", f"{industry_value:.0%}")

st.subheader("Slide 2 — Purchase influence")
purchase_labels = [option.label for option in purchase_options if option.label.lower() != "other"]
purchase_profile = st.selectbox(
    "Q12 — Product or service the advertiser sells",
    purchase_labels,
    help="This controls the slide-two headline and platform chart. The purchase-category chart uses Q12 Total percentages.",
)

st.subheader("Review")
review_a, review_b = st.columns(2)
review_a.metric("Reach", reach.strip() or "—", help="From the HubSpot audience segment.")
review_b.metric("Decision-Makers", f"{fixed_decision:.0%}", help="Fixed from Q10 purchase influence net / Total.")

ready = bool(company_name.strip() and reach.strip() and qualified_description.strip() and industry_description.strip())
if not ready:
    st.warning("Enter the company name, HubSpot reach, and both audience descriptions to enable generation.")

if st.button("Generate PowerPoint", type="primary", disabled=not ready, use_container_width=True):
    try:
        template = (Path(__file__).parent / "assets" / "sme_audience_template.pptx").read_bytes()
        result = generate_presentation(
            template_bytes=template,
            rows=rows,
            company_name=company_name.strip(),
            reach=reach.strip(),
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

st.caption("Slide 1 uses HubSpot plus Q14, Q13, and Q10. Slide 2 uses Q12.")
