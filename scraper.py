"""
Product Scraper -> Excel
------------------------
Scrapes every product (title, price, rating, stock, category, URL) from
https://books.toscrape.com (a public sandbox site made for scraping practice)
and exports a formatted Excel file.

Usage:
    python scraper.py                  # scrape all pages
    python scraper.py --pages 3        # scrape only the first 3 pages
    python scraper.py --out books.xlsx # custom output file
"""

import argparse
import time
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_URL = "https://books.toscrape.com/"
HEADERS = {"User-Agent": "Mozilla/5.0 (portfolio-scraper)"}
RATING_MAP = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}


def get_soup(session: requests.Session, url: str, retries: int = 3) -> BeautifulSoup:
    """Download a page and return parsed HTML, retrying on network errors."""
    for attempt in range(1, retries + 1):
        try:
            resp = session.get(url, headers=HEADERS, timeout=15)
            resp.raise_for_status()
            resp.encoding = "utf-8"
            return BeautifulSoup(resp.text, "lxml")
        except requests.RequestException as exc:
            if attempt == retries:
                raise
            print(f"  ! {exc} — retrying ({attempt}/{retries})")
            time.sleep(2 * attempt)


def parse_listing(soup: BeautifulSoup, page_url: str) -> list[dict]:
    """Extract basic product info from a listing page."""
    items = []
    for card in soup.select("article.product_pod"):
        link = card.h3.a
        rating_word = card.select_one("p.star-rating")["class"][1]
        price_text = card.select_one("p.price_color").get_text(strip=True)
        items.append(
            {
                "Title": link["title"],
                "Price (GBP)": float(price_text.replace("£", "").replace("Â", "")),
                "Rating": RATING_MAP.get(rating_word, 0),
                "In Stock": "In stock" in card.select_one("p.availability").get_text(),
                "URL": urljoin(page_url, link["href"]),
            }
        )
    return items


def parse_category(soup: BeautifulSoup) -> str:
    crumbs = soup.select("ul.breadcrumb li a")
    return crumbs[-1].get_text(strip=True) if len(crumbs) >= 3 else ""


def scrape(max_pages: int | None = None, with_category: bool = True, delay: float = 0.5) -> pd.DataFrame:
    session = requests.Session()
    url = urljoin(BASE_URL, "catalogue/page-1.html")
    products, page = [], 1

    while url and (max_pages is None or page <= max_pages):
        print(f"Page {page}: {url}")
        soup = get_soup(session, url)
        items = parse_listing(soup, url)

        if with_category:
            for item in items:
                detail = get_soup(session, item["URL"])
                item["Category"] = parse_category(detail)
                time.sleep(delay / 5)

        products.extend(items)
        nxt = soup.select_one("li.next a")
        url = urljoin(url, nxt["href"]) if nxt else None
        page += 1
        time.sleep(delay)

    return pd.DataFrame(products)


def export_excel(df: pd.DataFrame, path: str) -> None:
    """Write data + a summary sheet, with styled headers and auto column widths."""
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Products", index=False)

        if "Category" in df.columns and not df.empty:
            summary = (
                df.groupby("Category")
                .agg(Products=("Title", "count"),
                     Avg_Price=("Price (GBP)", "mean"),
                     Avg_Rating=("Rating", "mean"))
                .round(2)
                .sort_values("Products", ascending=False)
                .reset_index()
            )
            summary.to_excel(writer, sheet_name="Summary", index=False)

        header_fill = PatternFill("solid", fgColor="1F4E78")
        for ws in writer.book.worksheets:
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center")
            for i, col in enumerate(ws.columns, start=1):
                width = max(len(str(c.value or "")) for c in col)
                ws.column_dimensions[get_column_letter(i)].width = min(width + 2, 60)
            ws.freeze_panes = "A2"


def main():
    parser = argparse.ArgumentParser(description="Scrape products into Excel")
    parser.add_argument("--pages", type=int, default=None, help="max pages to scrape")
    parser.add_argument("--out", default="products.xlsx", help="output Excel file")
    parser.add_argument("--no-category", action="store_true", help="skip product pages (faster)")
    args = parser.parse_args()

    df = scrape(args.pages, with_category=not args.no_category)
    export_excel(df, args.out)
    print(f"\nDone: {len(df)} products saved to {args.out}")


if __name__ == "__main__":
    main()
