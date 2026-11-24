import re
import time
import logging
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
import phonenumbers


INPUT_FILE = "sites.xlsx"
OUTPUT_FILE = "contacts_result.xlsx"
URL_COLUMN = "URL"
TIMEOUT = 10
DELAY = 1.0
MAX_CONTACT_PAGES = 3

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

CONTACT_KEYWORDS = [
    "contact", "contacts", "contact-us", "contactus",
    "about", "about-us", "get-in-touch", "reach-us",
    "support", "feedback",
]

EMAIL_RE = re.compile(
    r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}",
    re.IGNORECASE,
)

PHONE_RE = re.compile(r"(?:\+?\d[\d\s\-\(\)]{8,}\d)")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)


def normalize_url(url: str) -> str:
    if not isinstance(url, str):
        return ""
    url = url.strip()
    if not url:
        return ""
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    return url


def fetch(url: str):
    try:
        resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        resp.raise_for_status()
        if resp.encoding is None or resp.encoding.lower() == "iso-8859-1":
            resp.encoding = resp.apparent_encoding
        return resp.text, resp.url
    except Exception as e:
        log.warning(f"Failed to load {url}: {e}")
        return None


def find_contact_links(base_url: str, html: str):
    soup = BeautifulSoup(html, "lxml")
    links = []
    seen = set()

    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        text = (a.get_text() or "").strip().lower()
        title = (a.get("title") or "").lower()
        combined = f"{text} {title} {href.lower()}"

        if any(kw in combined for kw in CONTACT_KEYWORDS):
            full = urljoin(base_url, href)
            if urlparse(full).netloc == urlparse(base_url).netloc:
                if full not in seen:
                    seen.add(full)
                    links.append(full)

    return links[:MAX_CONTACT_PAGES]


def extract_emails(text: str):
    found = set(EMAIL_RE.findall(text))
    cleaned = set()
    for e in found:
        e_low = e.lower()
        if any(bad in e_low for bad in [".png", ".jpg", ".gif", "example.", "sentry.io"]):
            continue
        cleaned.add(e_low)
    return cleaned


def extract_phones(text: str, default_region: str = "US"):
    phones = set()

    for match in phonenumbers.PhoneNumberMatcher(text, default_region):
        if phonenumbers.is_valid_number(match.number):
            phones.add(phonenumbers.format_number(
                match.number, phonenumbers.PhoneNumberFormat.INTERNATIONAL
            ))

    if not phones:
        for raw in PHONE_RE.findall(text):
            digits = re.sub(r"\D", "", raw)
            if 10 <= len(digits) <= 15:
                phones.add(raw.strip())

    return phones


def extract_contacts_from_html(html: str, region: str = "US"):
    soup = BeautifulSoup(html, "lxml")

    emails = set()
    phones = set()

    for a in soup.find_all("a", href=True):
        href = a["href"]
        if href.lower().startswith("mailto:"):
            addr = href[7:].split("?")[0].strip().lower()
            if "@" in addr:
                emails.add(addr)
        if href.lower().startswith("tel:"):
            raw = href[4:].strip()
            norm = re.sub(r"[^\d+]", "", raw)
            if len(norm) >= 10:
                phones.add(raw)

    text = soup.get_text(" ", strip=True)
    emails |= extract_emails(text)

    if phones:
        phones |= extract_phones(text, region)
    else:
        phones = extract_phones(text, region)

    return emails, phones


def process_site(url: str, region: str = "US"):
    result = {
        "URL": url,
        "Status": "",
        "Email": "",
        "Phone": "",
        "Contact Pages Found": "",
    }

    fetched = fetch(url)
    if not fetched:
        result["Status"] = "Load error"
        return result

    html, final_url = fetched
    all_emails, all_phones = extract_contacts_from_html(html, region)

    contact_urls = find_contact_links(final_url, html)
    result["Contact Pages Found"] = "; ".join(contact_urls)

    for cu in contact_urls:
        time.sleep(0.5)
        page = fetch(cu)
        if not page:
            continue
        chtml, _ = page
        e, p = extract_contacts_from_html(chtml, region)
        all_emails |= e
        all_phones |= p

    result["Email"] = "; ".join(sorted(all_emails))
    result["Phone"] = "; ".join(sorted(all_phones))
    result["Status"] = "OK" if (all_emails or all_phones) else "No contacts found"
    return result


def main():
    try:
        df = pd.read_excel(INPUT_FILE)
    except FileNotFoundError:
        log.error(f"File {INPUT_FILE} not found. Create an Excel file with a '{URL_COLUMN}' column.")
        return

    if URL_COLUMN not in df.columns:
        log.error(f"Column '{URL_COLUMN}' not found. Available columns: {list(df.columns)}")
        return

    urls = [normalize_url(u) for u in df[URL_COLUMN].tolist()]
    urls = [u for u in urls if u]

    log.info(f"Processing {len(urls)} sites...")
    results = []

    for i, url in enumerate(urls, 1):
        log.info(f"[{i}/{len(urls)}] {url}")
        try:
            results.append(process_site(url))
        except Exception as e:
            log.exception(f"Error processing {url}: {e}")
            results.append({
                "URL": url,
                "Status": f"Error: {e}",
                "Email": "",
                "Phone": "",
                "Contact Pages Found": "",
            })
        time.sleep(DELAY)

    out_df = pd.DataFrame(results)
    out_df.to_excel(OUTPUT_FILE, index=False)
    log.info(f"Done! Results saved to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
