# MOI + PACI bulk lookup (Streamlit + Playwright)

## Run locally

```bash
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
playwright install chromium
streamlit run app.py
```

Open http://localhost:8501, upload `test.xlsx`, press **بدء الاستعلام**, then download the result.

## Excel format
Columns (exact names): `الرقم المدني`, `الداخليه`, `تجديد البطاقه`, `حاله البطاقه`.
Only the first column needs values; the other three are filled by the app.

## Notes
- Both sites use Google reCAPTCHA v3 (invisible, score-based). Running with a
  real Chromium via Playwright normally passes; if you see many
  "حدث خطأ" / timeout results, turn Headless off in the sidebar and/or raise
  the delay between requests.
- Selectors are documented at the top of `app.py` in case the sites change.
- Long runs (hundreds of rows) keep going on the server even if you close the
  tab: reopening the app re-attaches to the running job, shows progress and
  offers a partial download at any time. Only one job runs per container.
- If a site stops answering, the app pauses with back-off after 5 consecutive
  failed rows and gives up after 40 with a clear message. Download the partial
  results, then re-upload them later with "تخطي الخلايا المعبأة" enabled to
  continue from where it stopped.
