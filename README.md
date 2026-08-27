# SME Audience Slide Generator

A user-friendly Streamlit application that converts the SME audience-insights frequency report into a branded two-slide advertiser presentation.

## What users configure

- Audience-insights `.xlsx` workbook
- Advertiser/company name
- Report year
- Qualified-audience net or verified percentage
- Industry-audience net or verified percentage
- Audience descriptions
- Purchase focus: Machine Tools, Manufacturing Software, Robotics/Automation, or Machine Tools + Automation

The form follows the research workflow directly:

- HubSpot supplies the Reach field.
- Q14 job functions are multi-select. Individual mutually exclusive roles are combined with OR logic; overlapping aggregate/net rows must be selected alone.
- Q13 industries are multi-select. One industry uses its Total value; multiple industries require a verified deduplicated OR-net from respondent-level data or the survey system.
- Q10 supplies the fixed Decision-Makers result.
- Q12 supplies both charts on slide 2.

The uploaded workbook is a frequency report. Because respondents can choose multiple industries, arbitrary Q13 OR-nets cannot be calculated by adding the displayed marginal percentages. The interface explicitly requests the verified union when multiple industries are selected, preventing duplicated respondents. Decision-Makers is always calculated from `purchase influence net / Total` and cannot be changed in the interface.

## Run locally

```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Push this repository to GitHub.
2. In Streamlit Community Cloud, create an app from the repository.
3. Set the entry point to `app.py`.
4. Deploy. No secrets are required.

## Data behavior

The app searches by question, response, and audience-column labels—not fixed row numbers. This makes it resilient to row movement in refreshed frequency reports. Company-specific custom nets can be selected when present; otherwise, the user can enter a verified percentage.
