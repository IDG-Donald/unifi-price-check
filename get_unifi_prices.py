#!/usr/bin/env python3
"""Scrape current Canadian UniFi Store pricing into an Excel-friendly CSV.

Design goals:
- Use the rendered UniFi Store rather than a hard-coded third-party mirror.
- Discover product URLs from top-level category pages.
- Prefer structured product metadata (JSON-LD/meta tags), with text fallbacks.
- Validate the new dataset before replacing the existing CSV.
- Exit non-zero on failure so GitHub Actions never commits an ERROR row over good data.
"""

from __future__ import annotations

import csv
import json
import os
import re
import sys
import tempfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import urljoin, urlparse, urlunparse

from playwright.sync_api import BrowserContext, Page, TimeoutError as PlaywrightTimeoutError, sync_playwright

STORE_ROOT = "https://ca.store.ui.com"
OUTPUT_FILE = Path(os.getenv("UNIFI_OUTPUT_FILE", "unifi_prices.csv"))
MIN_PRODUCTS = int(os.getenv("UNIFI_MIN_PRODUCTS", "40"))
HEADLESS = os.getenv("UNIFI_HEADLESS", "1").lower() not in {"0", "false", "no"}
NAV_TIMEOUT_MS = int(os.getenv("UNIFI_NAV_TIMEOUT_MS", "60000"))

# Top-level UniFi sections. These pages expose the store catalog and are much
# less brittle than maintaining a long hand-written list of individual products.
CATEGORY_PAGES = {
    "Cloud Gateways": f"{STORE_ROOT}/ca/en/category/all-unifi-cloud-gateways",
    "Switching": f"{STORE_ROOT}/ca/en/category/all-switching",
    "WiFi": f"{STORE_ROOT}/ca/en/category/all-wifi",
    "Physical Security": f"{STORE_ROOT}/ca/en/category/all-physical-security",
    "Door Access": f"{STORE_ROOT}/ca/en/category/all-door-access",
    "Integrations": f"{STORE_ROOT}/ca/en/category/all-integrations",
    "Advanced Hosting": f"{STORE_ROOT}/ca/en/category/all-advanced-hosting",
    "Accessories": f"{STORE_ROOT}/ca/en/category/all-accessory-tech",
}

CSV_FIELDS = [
    "SKU",
    "Product Name",
    "Price (CAD)",
    "Line/Category",
    "Availability",
    "Product URL",
    "Price Type",
    "Updated UTC",
]

KNOWN_SKUS = {
    "U7-Pro",
    "UDM-Pro",
    "UCG-Ultra",
    "UXG-Lite",
}

PRICE_RE = re.compile(r"(?P<from>From\s+)?\$\s*(?P<price>[0-9][0-9,]*(?:\.\d{2})?)", re.I)
SKU_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{1,49}$")

IGNORE_SKU_LINES = {
    "overview",
    "technical",
    "in the box",
    "faq",
    "add to cart",
    "sold out",
    "select",
    "compare",
    "learn more",
    "ui care",
    "new",
}


@dataclass(frozen=True)
class Product:
    sku: str
    name: str
    price_cad: float
    category: str
    availability: str
    url: str
    price_type: str
    updated_utc: str

    def csv_row(self) -> dict[str, Any]:
        return {
            "SKU": self.sku,
            "Product Name": self.name,
            "Price (CAD)": f"{self.price_cad:.2f}",
            "Line/Category": self.category,
            "Availability": self.availability,
            "Product URL": self.url,
            "Price Type": self.price_type,
            "Updated UTC": self.updated_utc,
        }


def canonical_url(url: str) -> str:
    """Normalize store URLs so duplicate links collapse to one product."""
    absolute = urljoin(STORE_ROOT, url)
    parsed = urlparse(absolute)
    return urlunparse((parsed.scheme, parsed.netloc, parsed.path.rstrip("/"), "", "", ""))


def goto_with_retry(page: Page, url: str, attempts: int = 3) -> None:
    last_error: Exception | None = None
    for attempt in range(1, attempts + 1):
        try:
            print(f"  opening ({attempt}/{attempts}): {url}")
            page.goto(url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)
            page.wait_for_timeout(1200)
            return
        except PlaywrightTimeoutError as exc:
            last_error = exc
            print(f"  navigation timeout: {exc}")
        except Exception as exc:  # browser/network errors vary by platform
            last_error = exc
            print(f"  navigation error: {exc}")

        page.wait_for_timeout(1500 * attempt)

    raise RuntimeError(f"Could not load {url} after {attempts} attempts") from last_error


def scroll_catalog(page: Page) -> None:
    """Trigger lazy-loaded catalog cards without assuming a particular DOM layout."""
    previous_height = -1
    stable_rounds = 0

    for _ in range(18):
        height = page.evaluate("document.body.scrollHeight")
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        page.wait_for_timeout(500)

        if height == previous_height:
            stable_rounds += 1
            if stable_rounds >= 2:
                break
        else:
            stable_rounds = 0
            previous_height = height

    page.evaluate("window.scrollTo(0, 0)")


def links_from_page(page: Page, fragment: str) -> set[str]:
    hrefs = page.locator(f'a[href*="{fragment}"]').evaluate_all(
        "els => els.map(el => el.href).filter(Boolean)"
    )
    return {canonical_url(href) for href in hrefs if "ca.store.ui.com" in href}


def discover_catalog(page: Page) -> tuple[dict[str, str], dict[str, str]]:
    """Return product_url -> category plus collection_url -> category."""
    products: dict[str, str] = {}
    collections: dict[str, str] = {}

    for category, url in CATEGORY_PAGES.items():
        print(f"Discovering {category}...")
        goto_with_retry(page, url)
        scroll_catalog(page)

        product_links = links_from_page(page, "/products/")
        collection_links = links_from_page(page, "/collections/")

        print(
            f"  found {len(product_links)} product links and "
            f"{len(collection_links)} collection links"
        )

        for product_url in product_links:
            products.setdefault(product_url, category)
        for collection_url in collection_links:
            collections.setdefault(collection_url, category)

    # Some store cards are "Product Collection" pages. Visit those once and
    # harvest the actual /products/ URLs they contain.
    for index, (collection_url, category) in enumerate(sorted(collections.items()), start=1):
        print(f"Expanding collection {index}/{len(collections)}: {collection_url}")
        try:
            goto_with_retry(page, collection_url, attempts=2)
            scroll_catalog(page)
            for product_url in links_from_page(page, "/products/"):
                products.setdefault(product_url, category)
        except Exception as exc:
            # A single collection should not destroy an otherwise healthy scrape.
            print(f"  warning: collection skipped: {exc}")

    return products, collections


def infer_category_from_url(url: str) -> str:
    path = urlparse(url).path.lower()
    if "door" in path or "access" in path:
        return "Door Access"
    if "camera" in path or "security" in path or "protect" in path:
        return "Physical Security"
    if "wifi" in path or "wireless" in path:
        return "WiFi"
    if "switch" in path:
        return "Switching"
    if "gateway" in path or "cloud-key" in path:
        return "Cloud Gateways"
    if "storage" in path or "power" in path or "lte" in path or "voip" in path:
        return "Integrations"
    if "accessor" in path or "mount" in path or "cable" in path:
        return "Accessories"
    return "General"


def json_ld_documents(page: Page) -> list[Any]:
    docs: list[Any] = []
    texts = page.locator('script[type="application/ld+json"]').all_text_contents()
    for text in texts:
        try:
            docs.append(json.loads(text))
        except (json.JSONDecodeError, TypeError):
            continue
    return docs


def walk_json(value: Any) -> Iterable[dict[str, Any]]:
    if isinstance(value, dict):
        yield value
        for child in value.values():
            yield from walk_json(child)
    elif isinstance(value, list):
        for child in value:
            yield from walk_json(child)


def product_json_ld(page: Page) -> dict[str, Any] | None:
    for document in json_ld_documents(page):
        for obj in walk_json(document):
            obj_type = obj.get("@type")
            types = obj_type if isinstance(obj_type, list) else [obj_type]
            if any(str(t).lower() == "product" for t in types if t):
                return obj
    return None


def first_meta_content(page: Page, selectors: list[str]) -> str | None:
    for selector in selectors:
        locator = page.locator(selector).first
        if locator.count():
            value = locator.get_attribute("content")
            if value and value.strip():
                return value.strip()
    return None


def clean_lines(text: str) -> list[str]:
    return [re.sub(r"\s+", " ", line).strip() for line in text.splitlines() if line.strip()]


def first_visible_heading(page: Page) -> str | None:
    for selector in ["main h1", "h1"]:
        locator = page.locator(selector).first
        if locator.count():
            text = locator.inner_text().strip()
            if text:
                return text
    return None


def parse_offer(offers: Any) -> tuple[float | None, str | None, str | None]:
    offer_list = offers if isinstance(offers, list) else [offers]
    for offer in offer_list:
        if not isinstance(offer, dict):
            continue

        currency = str(offer.get("priceCurrency") or "").upper()
        raw_price = offer.get("price") or offer.get("lowPrice")
        if raw_price is None:
            continue

        try:
            price = float(str(raw_price).replace(",", ""))
        except ValueError:
            continue

        availability = offer.get("availability")
        if isinstance(availability, str):
            availability = availability.rsplit("/", 1)[-1]
        else:
            availability = None

        if not currency or currency == "CAD":
            return price, currency or "CAD", availability

    return None, None, None


def fallback_sku(lines: list[str], product_name: str) -> str | None:
    try:
        start = lines.index(product_name)
    except ValueError:
        start = 0

    # On current UniFi product pages the model/SKU is immediately around the
    # product title. Keep the search window deliberately small to avoid specs.
    candidates = lines[start + 1 : start + 12]

    for line in candidates:
        normalized = line.strip()
        lower = normalized.lower()
        if lower in IGNORE_SKU_LINES or normalized == product_name:
            continue
        if "$" in normalized or len(normalized) > 50 or " " in normalized:
            continue
        if not SKU_RE.fullmatch(normalized):
            continue
        # Avoid tiny UI tokens while still permitting valid SKUs such as UX/E7.
        if len(normalized) < 2:
            continue
        return normalized

    return None


def fallback_price(lines: list[str], product_name: str) -> tuple[float | None, str]:
    try:
        start = lines.index(product_name)
    except ValueError:
        start = 0

    # Accessories and UI Care prices often appear later on the page, so only
    # inspect the product-summary area near the title.
    for line in lines[start : start + 30]:
        match = PRICE_RE.search(line)
        if match:
            return float(match.group("price").replace(",", "")), (
                "From" if match.group("from") else "Exact"
            )

    return None, "Unknown"


def normalize_availability(value: str | None, lines: list[str]) -> str:
    if value:
        key = value.lower().replace("_", "")
        if "instock" in key:
            return "In Stock"
        if "outofstock" in key or "soldout" in key:
            return "Out of Stock"
        if "preorder" in key:
            return "Preorder"

    summary = "\n".join(lines[:60]).lower()
    if "sold out" in summary or "notify me when available" in summary:
        return "Out of Stock"
    if "preorder" in summary or "pre-order" in summary:
        return "Preorder"
    if "add to cart" in summary:
        return "In Stock"
    if "select" in summary:
        return "Selectable"
    return "Unknown"


def scrape_product(page: Page, url: str, category: str, timestamp: str) -> Product | None:
    goto_with_retry(page, url, attempts=2)

    # Cloudflare/challenge pages should fail loudly instead of becoming fake data.
    title = page.title().lower()
    body_text = page.locator("body").inner_text(timeout=10_000)
    if "just a moment" in title or "verify you are human" in body_text.lower():
        raise RuntimeError("UniFi Store returned an anti-bot challenge page")

    lines = clean_lines(body_text)
    structured = product_json_ld(page) or {}

    name = str(structured.get("name") or "").strip()
    if not name:
        name = first_visible_heading(page) or ""
    if not name:
        name = first_meta_content(page, ['meta[property="og:title"]']) or ""
    name = re.sub(r"\s+-\s+Ubiquiti Store$", "", name).strip()

    sku = str(structured.get("sku") or structured.get("mpn") or "").strip()
    if not sku:
        sku = first_meta_content(page, ['meta[itemprop="sku"]']) or ""
    if not sku and name:
        sku = fallback_sku(lines, name) or ""

    price, _currency, offer_availability = parse_offer(structured.get("offers"))
    price_type = "Exact"

    if price is None:
        raw_meta_price = first_meta_content(
            page,
            [
                'meta[property="product:price:amount"]',
                'meta[itemprop="price"]',
            ],
        )
        if raw_meta_price:
            try:
                price = float(raw_meta_price.replace(",", ""))
            except ValueError:
                price = None

    if price is None and name:
        price, price_type = fallback_price(lines, name)
    else:
        # Preserve "From" if visible next to the primary displayed price.
        try:
            start = lines.index(name)
        except ValueError:
            start = 0
        nearby = " ".join(lines[start : start + 20])
        if re.search(r"From\s+\$", nearby, re.I):
            price_type = "From"

    availability = normalize_availability(offer_availability, lines)

    if not name or not sku or price is None or price <= 0:
        print(
            "  warning: skipping incomplete product "
            f"name={name!r} sku={sku!r} price={price!r} url={url}"
        )
        return None

    return Product(
        sku=sku,
        name=name,
        price_cad=round(price, 2),
        category=category,
        availability=availability,
        url=url,
        price_type=price_type,
        updated_utc=timestamp,
    )


def validate_products(products: list[Product], discovered_count: int) -> None:
    errors: list[str] = []

    if len(products) < MIN_PRODUCTS:
        errors.append(f"only {len(products)} valid products; minimum is {MIN_PRODUCTS}")

    if discovered_count and len(products) / discovered_count < 0.60:
        errors.append(
            f"only {len(products)}/{discovered_count} discovered product pages parsed successfully"
        )

    skus = {product.sku for product in products}
    if not (skus & KNOWN_SKUS):
        errors.append(
            "none of the sanity-check SKUs were found: " + ", ".join(sorted(KNOWN_SKUS))
        )

    duplicate_skus = len(products) - len(skus)
    if duplicate_skus > max(5, len(products) // 10):
        errors.append(f"too many duplicate SKUs ({duplicate_skus})")

    if errors:
        raise RuntimeError("Dataset validation failed: " + "; ".join(errors))


def deduplicate(products: list[Product]) -> list[Product]:
    """Prefer the first discovered row for each SKU, then sort for stable diffs."""
    by_sku: dict[str, Product] = {}
    for product in products:
        by_sku.setdefault(product.sku, product)
    return sorted(by_sku.values(), key=lambda p: (p.category.lower(), p.name.lower(), p.sku.lower()))


def atomic_write_csv(products: list[Product], output_file: Path) -> None:
    output_file.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.NamedTemporaryFile(
        mode="w",
        newline="",
        encoding="utf-8",
        prefix=f".{output_file.name}.",
        suffix=".tmp",
        dir=output_file.parent,
        delete=False,
    ) as tmp:
        temp_path = Path(tmp.name)
        writer = csv.DictWriter(tmp, fieldnames=CSV_FIELDS)
        writer.writeheader()
        for product in products:
            writer.writerow(product.csv_row())

    os.replace(temp_path, output_file)


def new_context(browser) -> BrowserContext:
    return browser.new_context(
        locale="en-CA",
        timezone_id="America/Toronto",
        viewport={"width": 1440, "height": 1200},
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/136.0.0.0 Safari/537.36"
        ),
        extra_http_headers={
            "Accept-Language": "en-CA,en;q=0.9",
        },
    )


def run() -> None:
    timestamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat()

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=HEADLESS)
        context = new_context(browser)
        page = context.new_page()

        try:
            discovered, _collections = discover_catalog(page)
            if not discovered:
                raise RuntimeError("No product URLs were discovered from the UniFi Store")

            print(f"Discovered {len(discovered)} unique product pages")

            products: list[Product] = []
            failures = 0

            for index, (url, category) in enumerate(sorted(discovered.items()), start=1):
                print(f"Scraping {index}/{len(discovered)}: {url}")
                try:
                    product = scrape_product(page, url, category, timestamp)
                    if product:
                        products.append(product)
                    else:
                        failures += 1
                except Exception as exc:
                    failures += 1
                    print(f"  warning: product failed: {exc}")

            products = deduplicate(products)
            validate_products(products, len(discovered))
            atomic_write_csv(products, OUTPUT_FILE)

            print(
                f"Success: wrote {len(products)} products to {OUTPUT_FILE} "
                f"({failures} product pages skipped/failed)"
            )
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    try:
        run()
    except Exception as exc:
        print(f"FATAL: {exc}", file=sys.stderr)
        print(
            "Existing CSV was not intentionally replaced unless the new dataset "
            "completed validation.",
            file=sys.stderr,
        )
        raise SystemExit(1)
