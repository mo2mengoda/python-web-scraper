# Product Scraper → Excel

A Python web scraper that collects product data (title, price, rating, stock status, category, link) from an e-commerce catalogue and exports a clean, formatted Excel report.

Demo target: [books.toscrape.com](https://books.toscrape.com), a public sandbox made for scraping practice.

## Features
- Handles pagination automatically
- Visits each product page to get its category
- Retries on network errors, polite delays between requests
- Excel output with styled headers, auto column widths, frozen header row
- Extra **Summary** sheet: product count, average price and rating per category

## Usage
```bash
pip install -r requirements.txt
python scraper.py                  # all pages
python scraper.py --pages 3        # first 3 pages only
python scraper.py --no-category    # faster, skip product pages
python scraper.py --out books.xlsx
```

## Output
| Title | Price (GBP) | Rating | In Stock | URL | Category |
|---|---|---|---|---|---|
| A Light in the Attic | 51.77 | 3 | True | … | Poetry |

## Tech
Python · requests · BeautifulSoup · pandas · openpyxl

> Need data from a different website? The same structure adapts to any site: only the parsing functions change.
