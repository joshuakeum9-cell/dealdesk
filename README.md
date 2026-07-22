# DealDesk

DealDesk turns raw company financials into four consulting deliverables: a business summary, a management interview guide, an opportunity presentation, and a working Excel model with an executive email summary. Every output is tailored to a consulting practice: strategy, operations, or M&A.

**Live site:** https://joshuakeum9-cell.github.io/dealdesk/

## How it works

1. Choose your consulting practice and describe the engagement.
2. Upload financial statements (Excel or CSV), or load the built in sample company.
3. Review and generate. Download real .docx, .pptx, and .xlsx files.

Everything runs in the browser. There is no server, no API, and no data leaves your machine. The analysis is a rules engine: the app parses the financials, computes growth and margin metrics, and assembles each document from practice specific templates and threshold driven language.

## The four deliverables

| Deliverable | Format | What it contains |
| --- | --- | --- |
| Business Summary | Word | Company overview, financial summary with EBITDA margin and a projection column, recent news, key people, key products |
| Interview Guide | Word | 15 to 20 numbered questions in lettered topics with space for answers, pitched at CFO level |
| Opportunity Presentation | PowerPoint | A three slide loop: a summary slide that signposts, then two slides expanding each category in order, plus an achievability matrix appendix |
| Excel Model + Email | Excel + Word | A directional value model with labeled input cells and conservative, midpoint, and aggressive scenario formulas, plus an answer first email |

## Output branding

Documents generate in two modes. Ghost draft is the neutral firm style with DRAFT stamps. Client branded final applies the client's colors, fonts, and logo, set on the engagement form.

## Stack

Plain HTML, CSS, and JavaScript. Document generation uses SheetJS, docx, PptxGenJS, and JSZip from CDNs. Sample outputs are in the `samples/` folder.

## Author

Joshua Keum. Portfolio project.
