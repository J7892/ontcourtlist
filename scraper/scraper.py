"""
Ontario Court Daily Docket Scraper
Fetches court listings from ontariocourtdates.ca, parses entries into structured records,
and exports them to Excel (.xlsx), CSV, and SQLite.
"""

import os
import re
import sys
import time
import logging
import sqlite3
from datetime import datetime
from typing import List, Dict, Any, Optional

import requests
from bs4 import BeautifulSoup
import pandas as pd

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
logger = logging.getLogger(__name__)

BASE_URL = "https://www.ontariocourtdates.ca"
DEFAULT_PAGE = f"{BASE_URL}/Default.aspx"
DOCKET_PAGE = f"{BASE_URL}/daily-docket.aspx"

USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"


class OntarioCourtScraper:
    def __init__(self, target_date: str = "today", delay: float = 0.5):
        """
        :param target_date: "today" or "tomorrow"
        :param delay: delay in seconds between city requests to avoid rate limits
        """
        self.target_date = target_date
        self.delay = delay
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": USER_AGENT,
            "Referer": DOCKET_PAGE,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
        })

    def _extract_aspnet_tokens(self, soup: BeautifulSoup) -> Dict[str, str]:
        """Extract ASP.NET hidden state variables from the form."""
        tokens = {}
        for token_id in ["__VIEWSTATE", "__VIEWSTATEGENERATOR", "__EVENTVALIDATION"]:
            tag = soup.find("input", {"id": token_id})
            tokens[token_id] = tag.get("value", "") if tag else ""
        return tokens

    def initialize_session(self) -> BeautifulSoup:
        """Loads Default.aspx, checks the terms checkbox, and navigates to daily-docket."""
        logger.info("Initializing session and accepting terms...")
        resp = self.session.get(DEFAULT_PAGE, timeout=30)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        tokens = self._extract_aspnet_tokens(soup)

        enter_data = {
            "__VIEWSTATE": tokens["__VIEWSTATE"],
            "__VIEWSTATEGENERATOR": tokens["__VIEWSTATEGENERATOR"],
            "__EVENTVALIDATION": tokens["__EVENTVALIDATION"],
            "ctl00$MainContent$chkAgree": "on",
            "ctl00$MainContent$btnEnter": "ENTER"
        }

        resp2 = self.session.post(DEFAULT_PAGE, data=enter_data, timeout=30)
        resp2.raise_for_status()

        if self.target_date == "tomorrow":
            # Switch to tomorrow's page
            logger.info("Navigating to tomorrow dockets page...")
            resp_tmw = self.session.get(f"{BASE_URL}/tomorrow/", timeout=30)
            resp_tmw.raise_for_status()
            soup2 = BeautifulSoup(resp_tmw.text, "html.parser")
        else:
            soup2 = BeautifulSoup(resp2.text, "html.parser")

        return soup2

    def get_cities(self, soup: BeautifulSoup) -> List[str]:
        """Extract all valid municipality names from the city dropdown."""
        city_select = soup.find("select", {"id": "ctl00_MainContent_ddlCity"})
        if not city_select:
            return []
        cities = []
        for opt in city_select.find_all("option"):
            val = opt.get("value", "").strip()
            if val and not val.startswith("---"):
                cities.append(val)
        return cities

    def fetch_city_offices(self, soup: BeautifulSoup, city: str) -> (BeautifulSoup, List[Dict[str, str]]):
        """
        Triggers the postback for a specific city to obtain office locations.
        """
        tokens = self._extract_aspnet_tokens(soup)
        endpoint = f"{BASE_URL}/tomorrow/daily-docket.aspx" if self.target_date == "tomorrow" else DOCKET_PAGE

        post_data = {
            "__EVENTTARGET": "ctl00$MainContent$ddlCity",
            "__EVENTARGUMENT": "",
            "__VIEWSTATE": tokens["__VIEWSTATE"],
            "__VIEWSTATEGENERATOR": tokens["__VIEWSTATEGENERATOR"],
            "__EVENTVALIDATION": tokens["__EVENTVALIDATION"],
            "ctl00$MainContent$ddlCourt": "both",
            "ctl00$MainContent$ddlCity": city,
            "ctl00$MainContent$ddlLob": "0",
            "ctl00$MainContent$hbody": "",
            "body": ""
        }

        resp = self.session.post(endpoint, data=post_data, timeout=30)
        resp.raise_for_status()
        postback_soup = BeautifulSoup(resp.text, "html.parser")

        office_select = postback_soup.find("select", {"id": "ctl00_MainContent_listBoxCourtOffice"})
        offices = []
        if office_select:
            for opt in office_select.find_all("option"):
                code = opt.get("value", "").strip()
                name = opt.text.strip()
                if code:
                    offices.append({"code": code, "name": name})

        return postback_soup, offices

    def query_office_dockets(self, soup: BeautifulSoup, city: str, office_code: str) -> List[Dict[str, Any]]:
        """
        Submits the search form for a specific court location and parses results.
        """
        tokens = self._extract_aspnet_tokens(soup)
        endpoint = f"{BASE_URL}/tomorrow/daily-docket.aspx" if self.target_date == "tomorrow" else DOCKET_PAGE

        submit_data = {
            "__EVENTTARGET": "",
            "__EVENTARGUMENT": "",
            "__VIEWSTATE": tokens["__VIEWSTATE"],
            "__VIEWSTATEGENERATOR": tokens["__VIEWSTATEGENERATOR"],
            "__EVENTVALIDATION": tokens["__EVENTVALIDATION"],
            "ctl00$MainContent$ddlCourt": "both",
            "ctl00$MainContent$ddlCity": city,
            "ctl00$MainContent$ddlLob": "0",
            "ctl00$MainContent$listBoxCourtOffice": office_code,
            "ctl00$MainContent$hbody": "",
            "body": "",
            "ctl00$MainContent$btnSubmit": "SUBMIT"
        }

        resp = self.session.post(endpoint, data=submit_data, timeout=30)
        resp.raise_for_status()
        result_soup = BeautifulSoup(resp.text, "html.parser")

        return self._parse_docket_tables(result_soup, city)

    def _parse_docket_tables(self, soup: BeautifulSoup, city: str) -> List[Dict[str, Any]]:
        """Parses all court list tables on the result page."""
        records = []
        date_tag = soup.find(id="ctl00_MainContent_lblTodayDate")
        court_date_str = date_tag.text.strip() if date_tag else ""

        current_court = "Unknown Court"
        current_case_type = "General"

        # Walk through the content area elements sequentially
        content_area = soup.find(id="content_area")
        if not content_area:
            return records

        for element in content_area.descendants:
            if element.name in ["h2", "h3", "h4", "p"]:
                text = element.get_text(strip=True)
                if "Superior Court of Justice" in text:
                    current_court = "Superior Court of Justice"
                elif "Ontario Court of Justice" in text:
                    current_court = "Ontario Court of Justice"
                elif text in ["Criminal", "Civil", "Family", "Small Claims"]:
                    current_case_type = text

            elif element.name == "table" and "wet-boew-zebra" in element.get("class", []):
                rows = element.find_all("tr")
                if not rows:
                    continue

                headers = [th.get_text(strip=True) for th in rows[0].find_all(["th", "td"])]

                for row in rows[1:]:
                    cells = [td.get_text(strip=True) for td in row.find_all("td")]
                    if not cells or len(cells) < len(headers):
                        continue

                    row_dict = dict(zip(headers, cells))
                    record = {
                        "date_text": court_date_str,
                        "scrape_timestamp": datetime.utcnow().isoformat(),
                        "city": city,
                        "court": current_court,
                        "case_type": current_case_type,
                        "party_name": row_dict.get("Party Name", ""),
                        "case_number": row_dict.get("Case Number", ""),
                        "short_title": row_dict.get("Short Title of Proceedings", ""),
                        "time": row_dict.get("Time", ""),
                        "room": row_dict.get("Room", ""),
                        "event_type": row_dict.get("Appearance/Event Type", ""),
                        "method_of_attendance": row_dict.get("Method of Attendance", ""),
                        "docket_line": row_dict.get("Docket Line", "")
                    }
                    records.append(record)

        return records

    def scrape_all(self, cities_filter: Optional[List[str]] = None) -> List[Dict[str, Any]]:
        """
        Runs the full scraping cycle across all or filtered municipalities.
        """
        soup = self.initialize_session()
        available_cities = self.get_cities(soup)
        logger.info(f"Discovered {len(available_cities)} municipalities.")

        target_cities = cities_filter if cities_filter else available_cities
        all_records = []

        for idx, city in enumerate(target_cities, 1):
            logger.info(f"[{idx}/{len(target_cities)}] Scraping municipality: {city}...")
            try:
                # Fresh session per city to avoid ViewState desynchronization across calls
                fresh_soup = self.initialize_session()
                postback_soup, offices = self.fetch_city_offices(fresh_soup, city)

                if not offices:
                    logger.info(f"  No court offices listed for {city}.")
                    continue

                for off in offices:
                    logger.info(f"  Fetching dockets for office: {off['name']} ({off['code']})...")
                    records = self.query_office_dockets(postback_soup, city, off["code"])
                    logger.info(f"    Found {len(records)} case entries.")
                    all_records.extend(records)

                if self.delay > 0:
                    time.sleep(self.delay)

            except Exception as e:
                logger.error(f"Error scraping {city}: {e}", exc_info=True)

        logger.info(f"Finished scrape! Total entries collected: {len(all_records)}")
        return all_records


def save_data(records: List[Dict[str, Any]], output_dir: str = "data") -> Dict[str, str]:
    """
    Saves parsed records to Excel, CSV, and SQLite.
    """
    os.makedirs(output_dir, exist_ok=True)
    today_str = datetime.now().strftime("%Y-%m-%d")

    df = pd.DataFrame(records)

    excel_path = os.path.join(output_dir, f"ontario_court_dockets_{today_str}.xlsx")
    csv_path = os.path.join(output_dir, f"ontario_court_dockets_{today_str}.csv")
    latest_csv_path = os.path.join(output_dir, "latest.csv")
    db_path = os.path.join(output_dir, "court_dockets.db")

    if not df.empty:
        # Save Excel with formatting
        df.to_excel(excel_path, index=False, engine="openpyxl")
        df.to_csv(csv_path, index=False)
        df.to_csv(latest_csv_path, index=False)

        # Store in SQLite database for fast indexed searching
        conn = sqlite3.connect(db_path)
        df.to_sql("dockets", conn, if_exists="append", index=False)

        # Create search indexes
        cursor = conn.cursor()
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_party ON dockets (party_name);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_city ON dockets (city);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_date ON dockets (date_text);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_case ON dockets (case_number);")
        conn.commit()

        # Export high-compression deduplicated archive for DuckDB-Wasm in browser
        archive_df = pd.read_sql("SELECT DISTINCT * FROM dockets", conn)
        archive_parquet_path = os.path.join(output_dir, "dockets_archive.parquet")
        archive_df.to_parquet(archive_parquet_path, compression="snappy", index=False)
        conn.close()

        logger.info(f"Data saved successfully:")
        logger.info(f"  Excel:   {excel_path}")
        logger.info(f"  CSV:     {csv_path}")
        logger.info(f"  SQLite:  {db_path}")
        logger.info(f"  Parquet: {archive_parquet_path} ({len(archive_df)} records)")
    else:
        logger.warning("No records scraped; empty files not saved.")

    return {
        "excel": excel_path,
        "csv": csv_path,
        "sqlite": db_path
    }


if __name__ == "__main__":
    # Test run: can pass specific city names as CLI arguments e.g. python scraper.py Barrie Guelph
    selected_cities = sys.argv[1:] if len(sys.argv) > 1 else None
    scraper = OntarioCourtScraper(target_date="today", delay=0.5)
    data = scraper.scrape_all(cities_filter=selected_cities)
    save_data(data)
