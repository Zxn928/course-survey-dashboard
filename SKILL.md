---
name: course-survey-dashboard
description: Generate polished, shareable HTML data dashboards from survey, questionnaire, feedback, evaluation, course review, training review, customer research, CSV, TSV, XLSX, or Excel files. Use when the user asks to summarize spreadsheet-like survey data, visualize questionnaire results, create a data board/dashboard, produce an HTML webpage for external sharing, or turn uploaded Excel/CSV feedback into charts, insights, themes, and browsable response tables.
---

# Survey Data Dashboard

## Workflow

1. Inspect the source file with a structured parser. Do not rely on line counts for CSV files because quoted multiline responses are common.
2. Identify the response count, date/time fields, rating/satisfaction fields, suggestion fields, and open-text feedback fields.
3. Generate a concise analytical summary:
   - KPI cards: valid responses, date span, positive rate, very positive rate, actionable suggestions.
   - Rating charts: distribution and comparison for each rating column.
   - Text themes: group open responses into readable topics; use keyword counts only as a starting point and name themes in human language.
   - Suggestion themes: separate blanks / "none" answers from actionable suggestions.
   - Details table: include anonymized or minimally identifying row-level feedback unless the user explicitly asks for names.
4. Build a self-contained HTML page with inline CSS and JavaScript. Avoid CDN dependencies so the file can be sent externally as an attachment.
5. Create an external-sharing copy with a friendly filename. Also create a `.zip` copy when the user says they need to send it through WeChat, email, or mobile apps.
6. Verify the page by opening it with Playwright or the in-app browser:
   - The page loads with no console errors.
   - Tabs/buttons work.
   - The detail table row count matches parsed records.
   - Mobile viewport has no page-level horizontal overflow.

## Design Standards

- Use `ui-ux-pro-max` whenever available for dashboard UX, accessibility, responsive layout, chart readability, and visual polish.
- Use `spreadsheets` whenever available for robust spreadsheet parsing, validation, and chart/data conventions.
- Make the first screen the actual dashboard, not a landing page.
- Prefer a restrained analytics style: clear hierarchy, compact KPI cards, readable charts, and a neutral palette with 2-4 functional accent colors.
- Keep cards to repeated KPI/panel units; do not nest cards inside cards.
- Include a sticky tab or segmented navigation when the dashboard has multiple sections.
- Do not use emoji as icons. Use text labels or proper icon libraries already present in the project.
- Keep all content readable on mobile. A table may scroll inside its own container, but the page itself should not horizontally overflow.

## Script

Use `scripts/build_survey_dashboard.py` for a fast first pass:

```bash
python3 /Users/mikey/.codex/skills/course-survey-dashboard/scripts/build_survey_dashboard.py INPUT_FILE --output OUTPUT_HTML --title "课程问卷数据看板"
```

The script supports `.csv`, `.tsv`, and `.xlsx` when `openpyxl` is installed. It creates a self-contained HTML dashboard with inferred metrics, charts, tabs, filters, and a row-level feedback table.

After running the script, review the generated insights and patch the output or script when the dataset needs domain-specific theme names, renamed columns, privacy adjustments, or a more tailored visual narrative.

## External Sharing

Explain clearly that `file:///...` URLs work only on the current computer. For sharing:

- Send the generated `.html` file itself as an attachment.
- Send a `.zip` containing the `.html` if the channel blocks standalone HTML.
- If the user needs a clickable public URL, upload the HTML to a static host such as GitHub Pages, company intranet, OSS/CDN, Netlify, Vercel, or an approved internal file hosting service.

## Final Response

Return the deliverable links from the current task's `outputs/` directory only. Mention the key counts and the validation performed. Keep the response short.
