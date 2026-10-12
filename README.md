# TVS Lanka – Product NPS Dashboard

A private dashboard for the Product NPS Survey (Google Forms -> Google Sheets).

- `index.html` – the dashboard page (served by GitHub Pages).
- `apps-script/Code.gs` – the Google Apps Script that sits in the response sheet and returns cleaned, anonymised data when the correct access code is sent.

**Never commit the access code or any customer data to this repository.** The access code lives only in the Apps Script project (Script properties). Customer names, phone numbers and VINs are never sent to the page.

Setup: paste `Code.gs` into the sheet's Apps Script, run `setup()` to get the access code, deploy as a Web app (Execute as: Me, Access: Anyone), then put the Web app URL in `index.html` (`API_URL`).
