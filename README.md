# SME Audience Slide Generator

This application converts an SME audience-insights frequency report into a branded two-slide advertiser presentation.

## Important links

- [Open the slide generator](https://surveyslidegenerator.streamlit.app/)
- [SMEMedia repository](https://github.com/SMEMedia/SurveySlideGen)

## Create a presentation

1. Open the slide generator.
2. Upload the audience-insights workbook in .xlsx format.
3. Enter the advertiser or company name and report year.
4. Enter the qualified-audience and industry-audience percentages requested by the form.
5. Choose the purchase focus.
6. Review the audience descriptions and calculated results.
7. Generate and download the presentation.
8. Open the downloaded slides and verify the company name, year, percentages, labels, and charts before sharing.

## How survey values are used

- **HubSpot** supplies Reach.
- **Q14 job functions:** combine mutually exclusive roles using OR logic. Do not combine an aggregate/net row with its component roles.
- **Q13 industries:** use the Total value for one industry. For multiple industries, use a verified deduplicated OR-net from respondent-level data or the survey system.
- **Q10** supplies the fixed Decision-Makers result.
- **Q12** supplies both charts on slide 2.

Adding displayed Q13 percentages together can count the same respondent more than once. Use a verified union whenever multiple industries are selected.

## Troubleshooting

### The workbook is rejected

- Confirm it is an .xlsx audience-insights frequency report.
- Make sure the file opens normally in Excel and is not password protected.
- Download a fresh copy of the report and try again.
- If the error continues, send the report name and a screenshot of the message to the research support contact. Do not attach confidential respondent-level data to an open GitHub issue.

### A question or response cannot be found

- Confirm the workbook contains Q10, Q12, Q13, and Q14.
- Check whether question or response labels were renamed in the new report.
- Confirm the correct audience column was selected.
- Escalate label changes to the technical owner so the matching rules can be reviewed.

### A percentage looks too high

- Do not add overlapping net and component rows.
- For multiple Q13 industries, confirm the value is a deduplicated OR-net.
- Verify whether the form expects a percentage or a whole-number count.

### The presentation will not generate or download

- Review the form for missing required fields.
- Refresh the page and upload the workbook again.
- Try a shorter advertiser name if text is unusually long.
- If the issue continues, capture the visible message and contact support.

### The downloaded slides need adjustment

- Confirm all inputs before regenerating.
- Minor text wrapping can be corrected in the downloaded presentation.
- Do not manually change calculated research values without confirming them against the source report.

## Ongoing maintenance

- Use the newest approved audience-insights workbook.
- Verify every generated presentation before external use.
- Keep the source workbook and any respondent-level information in approved SME storage.
- Escalate report-format changes or calculation questions to the research and technical owners.
