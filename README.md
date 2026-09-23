# Ontario Daily Court Lists Scraper & Search Dashboard

An automated pipeline that scrapes Ontario's Daily Court Lists ([ontariocourtdates.ca](https://www.ontariocourtdates.ca/daily-docket.aspx)) every morning, parses every court location and case appearance into Excel (`.xlsx`), CSV, and SQLite, and provides a searchable web dashboard.

---

## Architecture

- **Scraper Engine (`scraper/scraper.py`)**:
  - Handles ASP.NET WebForms session negotiation, terms agreement postbacks, and cascading dropdown queries across all 62+ Ontario municipalities and court offices.
  - Normalizes case parties, docket lines, appearance reasons, courtroom numbers, and attendance methods (Videoconference vs In-Person).
- **Data Pipeline (`data/`)**:
  - Generates daily dated Excel workbooks: `ontario_court_dockets_YYYY-MM-DD.xlsx`.
  - Generates daily CSV files and a `latest.csv` snapshot.
  - Inserts records into an indexed SQLite database (`court_dockets.db`) for high-speed queries.
- **Searchable Web Dashboard (`web/app.py`)**:
  - Streamlit dashboard with real-time text search (Party name, Case number, Charge/Event), municipality filtering, court level selection, and instant CSV export.
- **Automation (`.github/workflows/daily_scrape.yml`)**:
  - Scheduled GitHub Actions cron job running Monday through Friday morning after court dockets refresh at 8:00 AM EST.
  - Commits updated data files and uploads 90-day downloadable artifacts.

---

## Quickstart

### 1. Setup Virtual Environment
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 2. Run Scraper
To scrape specific municipalities (test run):
```bash
python scraper/scraper.py Barrie Bracebridge
```

To run the complete province-wide scrape:
```bash
python scraper/scraper.py
```

Generated outputs will be saved in `data/`:
- `data/ontario_court_dockets_YYYY-MM-DD.xlsx`
- `data/ontario_court_dockets_YYYY-MM-DD.csv`
- `data/court_dockets.db`

### 3. Launch Search Dashboard
```bash
streamlit run web/app.py
```
Open [http://localhost:8501](http://localhost:8501) in your browser.

---

## GitHub Actions Configuration

The workflow is located in `.github/workflows/daily_scrape.yml`.

To enable pushing automated commits to your repository:
1. Go to repository **Settings** -> **Actions** -> **General**.
2. Under **Workflow permissions**, select **Read and write permissions**.
3. Save.
