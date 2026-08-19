#!/usr/bin/env python3
"""Fetch current UniFi Canada catalog pricing into an Excel-friendly CSV.

This version uses the public JSON payloads that back the UniFi Store's Next.js
category pages. It does not launch a browser or request every product page.

Failure behavior is intentionally conservative: a new CSV is written to a
temporary file, validated, and only then atomically replaces the previous CSV.
"""

from __future__ import annotations

import csv
import os
import re
import sys
import tempfile
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import urlsplit

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# Start from the Canadian storefront directly. UniFi currently redirects
# store.ui.com/ca/en to ca.store.ui.com; keeping the JSON requests on the
# effective storefront origin avoids receiving an HTML storefront/redirect page
# where Next.js JSON is expected.
STORE_ENTRY = os.getenv("UNIFI_STORE_ENTRY", "https://ca.store.ui.com").rstrip("/")
DISPLAY_BASE = os.getenv("UNIFI_DISPLAY_BASE", "https://ca.store.ui.com").rstrip("/")
REGION_PATH = os.getenv("UNIFI_REGION_PATH", "ca/en").strip("/")
OUTPUT_FILE = Path(os.getenv("UNIFI_OUTPUT_FILE", "unifi_prices.csv"))
REQUEST_TIMEOUT = float(os.getenv("UNIFI_REQUEST_TIMEOUT", "20"))
CATEGORY_DELAY_SECONDS = float(os.getenv("UNIFI_CATEGORY_DELAY", "0.25"))
MIN_PRODUCTS = int(os.getenv("UNIFI_MIN_PRODUCTS", "100"))
MIN_PRICED_RATIO = float(os.getenv("UNIFI_MIN_PRICED_RATIO", "0.70"))

# These nine broad routes cover the current UniFi catalog. Product rows are
# deduplicated across categories by product slug before CSV generation.
CATEGORIES: dict[str, str] = {
    "Cloud Gateways": "category/all-cloud-gateways",
    "Switching": "category/all-switching",
    "WiFi": "category/all-wifi",
    "Physical Security": "category/all-cameras-nvrs",
    "Door Access": "category/all-door-access",
    "Integrations": "category/all-integrations",
    "Advanced Hosting": "category/all-advanced-hosting",
    "Accessories": "category/accessories-cables-dacs",
    "Network Storage": "category/network-storage",
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

# Validate against stable product slugs rather than SKU fields because the
# storefront's JSON shape can omit SKU/model identifiers on some product cards.
KNOWN_PRODUCT_SLUGS = {
    "u7-pro",
    "udm-pro",
    "ucg-ultra",
    "uxg-lite",
}

BUILD_ID_RE = re.compile(r'"buildId":"([^"\\]+)"')
PRICE_TEXT_RE = re.compile(r"-?[0-9]+(?:[.,][0-9]+)?")
SKUISH_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._+/-]{1,63}$")


@dataclass(frozen=True)
class CatalogRow:
    sku: str
    product_name: str
    price_cad: Decimal
    category: str
    availability: str
    product_url: str
    price_type: str
    updated_utc: str

    def as_csv_row(self) -> dict[str, str]:
        return {
            "SKU": self.sku,
            "Product Name": self.product_name,
            "Price (CAD)": f"{self.price_cad:.2f}",
            "Line/Category": self.category,
            "Availability": self.availability,
            "Product URL": self.product_url,
            "Price Type": self.price_type,
            "Updated UTC": self.updated_utc,
        }


def make_session() -> requests.Session:
    session = requests.Session()
    retry = Retry(
        total=3,
        connect=3,
        read=3,
        status=3,
        backoff_factor=0.8,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET"}),
        respect_retry_after_header=True,
    )
    session.mount("https://", HTTPAdapter(max_retries=retry))
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/136.0 Safari/537.36"
            ),
            "Accept": "application/json,text/html;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-CA,en;q=0.9",
            "Cache-Control": "no-cache",
        }
    )
    return session


def origin_from_url(url: str) -> str:
    parsed = urlsplit(url)
    if not parsed.scheme or not parsed.netloc:
        raise RuntimeError(f"Could not determine storefront origin from URL: {url}")
    return f"{parsed.scheme}://{parsed.netloc}"


def get_store_context(session: requests.Session) -> tuple[str, str]:
    """Return (build_id, effective_store_origin) for the Canadian storefront.

    requests follows redirects automatically. The important detail is to build
    the /_next/data URL on the host that actually served the homepage, not on
    the pre-redirect hostname.
    """
    url = f"{STORE_ENTRY}/{REGION_PATH}"
    response = session.get(url, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()

    match = BUILD_ID_RE.search(response.text)
    if not match:
        raise RuntimeError("Could not find the UniFi Store Next.js buildId")

    return match.group(1), origin_from_url(response.url)


def category_json_url(store_origin: str, build_id: str, route: str) -> str:
    return f"{store_origin}/_next/data/{build_id}/{REGION_PATH}/{route}.json"


def iter_products_from_page_props(page_props: dict[str, Any]) -> Iterable[dict[str, Any]]:
    """Yield products from both known category JSON layouts."""
    for subcategory in page_props.get("subCategories", []) or []:
        if not isinstance(subcategory, dict):
            continue
        for product in subcategory.get("products", []) or []:
            if isinstance(product, dict):
                yield product

    for product in page_props.get("products", []) or []:
        if isinstance(product, dict):
            yield product


def fetch_category(
    session: requests.Session,
    store_origin: str,
    build_id: str,
    category_label: str,
    route: str,
) -> tuple[list[dict[str, Any]], str, str]:
    """Fetch one category, refreshing storefront context once if needed."""
    current_build = build_id
    current_origin = store_origin

    for attempt in range(2):
        url = category_json_url(current_origin, current_build, route)
        response = session.get(url, timeout=REQUEST_TIMEOUT)

        if response.status_code == 404 and attempt == 0:
            print(f"  {category_label}: buildId may have rotated; refreshing once")
            current_build, current_origin = get_store_context(session)
            continue

        response.raise_for_status()

        content_type = response.headers.get("content-type", "")
        try:
            payload = response.json()
        except ValueError as exc:
            snippet = " ".join(response.text[:220].split())
            raise RuntimeError(
                f"{category_label} did not return JSON "
                f"(status={response.status_code}, content-type={content_type!r}, "
                f"requested={url}, final={response.url}, body={snippet!r})"
            ) from exc

        page_props = payload.get("pageProps")
        if not isinstance(page_props, dict):
            raise RuntimeError(f"{category_label} JSON has no pageProps object")

        products = list(iter_products_from_page_props(page_props))
        if not products:
            raise RuntimeError(f"{category_label} returned zero products")

        return products, current_build, current_origin

    raise RuntimeError(f"Could not fetch {category_label}")


def fetch_catalog(session: requests.Session) -> dict[str, dict[str, Any]]:
    """Fetch all categories and deduplicate products by store slug."""
    build_id, store_origin = get_store_context(session)
    print(f"Store origin: {store_origin}")
    print(f"Store buildId: {build_id}")

    by_slug: dict[str, dict[str, Any]] = {}

    for index, (category_label, route) in enumerate(CATEGORIES.items(), start=1):
        print(f"Fetching category {index}/{len(CATEGORIES)}: {category_label}")
        products, build_id, store_origin = fetch_category(
            session=session,
            store_origin=store_origin,
            build_id=build_id,
            category_label=category_label,
            route=route,
        )

        category_new = 0
        for product in products:
            slug = str(product.get("slug") or "").strip()
            if not slug:
                continue

            if slug not in by_slug:
                product = dict(product)
                product["_category"] = category_label
                by_slug[slug] = product
                category_new += 1

        print(
            f"  received {len(products)} product records; "
            f"{category_new} new unique slugs"
        )

        if CATEGORY_DELAY_SECONDS > 0 and index < len(CATEGORIES):
            time.sleep(CATEGORY_DELAY_SECONDS)

    print(f"Fetched {len(by_slug)} unique products from {len(CATEGORIES)} category requests")
    return by_slug


def first_nonempty(mapping: dict[str, Any], keys: Iterable[str]) -> str:
    for key in keys:
        value = mapping.get(key)
        if value is None:
            continue
        text = str(value).strip()
        if text:
            return text
    return ""


def money_to_decimal(value: Any) -> Decimal | None:
    """Convert known UniFi price shapes to dollars.

    Money dictionaries use integer minor units (for example amount=23900 CAD).
    Plain numbers are treated as already-denominated dollar values, matching
    how the current storefront consumer code handles them.
    """
    if value is None:
        return None

    if isinstance(value, dict):
        amount = value.get("amount")
        if amount is None:
            return None
        try:
            return (Decimal(str(amount)) / Decimal("100")).quantize(Decimal("0.01"))
        except (InvalidOperation, ValueError):
            return None

    if isinstance(value, (int, float, Decimal)):
        try:
            return Decimal(str(value)).quantize(Decimal("0.01"))
        except InvalidOperation:
            return None

    if isinstance(value, str):
        match = PRICE_TEXT_RE.search(value.replace(",", ""))
        if not match:
            return None
        try:
            return Decimal(match.group(0)).quantize(Decimal("0.01"))
        except InvalidOperation:
            return None

    return None


def variant_price(variant: dict[str, Any]) -> Decimal | None:
    # displayPrice is intentionally preferred because it is the storefront's
    # presentation price; fall back to the raw price field when necessary.
    return money_to_decimal(variant.get("displayPrice")) or money_to_decimal(variant.get("price"))


def normalize_status(value: Any) -> str:
    status = str(value or "").strip().lower().replace("_", "").replace(" ", "")
    if status == "available" or "instock" in status:
        return "In Stock"
    if status in {"soldout", "unavailable"} or "outofstock" in status:
        return "Out of Stock"
    if status == "comingsoon":
        return "Coming Soon"
    if status in {"preorder", "pre-order"}:
        return "Preorder"
    return "Unknown"


def aggregate_status(variants: list[dict[str, Any]]) -> str:
    statuses = [normalize_status(v.get("status")) for v in variants]
    if "In Stock" in statuses:
        return "In Stock"
    if "Preorder" in statuses:
        return "Preorder"
    if "Coming Soon" in statuses:
        return "Coming Soon"
    if statuses and all(status == "Out of Stock" for status in statuses):
        return "Out of Stock"
    if "Out of Stock" in statuses:
        return "Out of Stock"
    return "Unknown"


def looks_like_sku(value: str) -> bool:
    value = value.strip()
    if not value or " " in value or len(value) > 64:
        return False
    return bool(SKUISH_RE.fullmatch(value))


def extract_sku(product: dict[str, Any], variant: dict[str, Any] | None, slug: str) -> tuple[str, bool]:
    """Return (lookup_key, is_real_sku).

    The API does not consistently expose a SKU on every card, so the stable
    store slug is the final fallback. That keeps every row addressable without
    fabricating a model number.
    """
    keys = ("sku", "model", "modelNumber", "mpn", "partNumber", "part_number")

    if variant:
        candidate = first_nonempty(variant, keys)
        if candidate and looks_like_sku(candidate):
            return candidate, True

    candidate = first_nonempty(product, keys)
    if candidate and looks_like_sku(candidate):
        return candidate, True

    # Some store payloads expose the model as a short title.
    for key in ("shortTitle", "shortName"):
        candidate = str(product.get(key) or "").strip()
        if candidate and looks_like_sku(candidate):
            return candidate, True

    return slug, False


def product_name(product: dict[str, Any], slug: str) -> str:
    name = first_nonempty(product, ("title", "name", "displayName", "productName"))
    return name or slug


def variant_label(variant: dict[str, Any]) -> str:
    return first_nonempty(variant, ("title", "name", "label", "displayName"))


def product_url(slug: str) -> str:
    return f"{DISPLAY_BASE}/{REGION_PATH}/products/{slug}"


def product_to_rows(product: dict[str, Any], timestamp: str) -> list[CatalogRow]:
    slug = str(product.get("slug") or "").strip()
    if not slug:
        return []

    category = str(product.get("_category") or "Unknown")
    base_name = product_name(product, slug)
    variants = [v for v in (product.get("variants") or []) if isinstance(v, dict)]

    priced_variants = [(variant, variant_price(variant)) for variant in variants]
    priced_variants = [(variant, price) for variant, price in priced_variants if price is not None and price > 0]

    if not priced_variants:
        # A few storefront product objects may carry a direct price instead.
        direct_price = money_to_decimal(product.get("displayPrice")) or money_to_decimal(product.get("price"))
        if direct_price is None or direct_price <= 0:
            return []

        sku, _is_real_sku = extract_sku(product, None, slug)
        return [
            CatalogRow(
                sku=sku,
                product_name=base_name,
                price_cad=direct_price,
                category=category,
                availability=aggregate_status(variants) if variants else "Unknown",
                product_url=product_url(slug),
                price_type="Exact",
                updated_utc=timestamp,
            )
        ]

    # If multiple variants expose distinct real SKUs, keep one row per SKU.
    distinct_variant_rows: list[CatalogRow] = []
    seen_variant_skus: set[str] = set()
    real_variant_sku_count = 0

    for variant, price in priced_variants:
        sku, is_real_sku = extract_sku(product, variant, slug)
        if is_real_sku:
            real_variant_sku_count += 1
        if sku in seen_variant_skus:
            continue
        seen_variant_skus.add(sku)

        label = variant_label(variant)
        name = base_name
        if len(priced_variants) > 1 and label and label.lower() not in base_name.lower():
            name = f"{base_name} - {label}"

        distinct_variant_rows.append(
            CatalogRow(
                sku=sku,
                product_name=name,
                price_cad=price,
                category=category,
                availability=normalize_status(variant.get("status")),
                product_url=product_url(slug),
                price_type="Exact",
                updated_utc=timestamp,
            )
        )

    if real_variant_sku_count == len(priced_variants):
        return distinct_variant_rows

    # Otherwise, expose one product row using the lowest displayed variant
    # price. This mirrors category cards that present a "From" price.
    lowest_price = min(price for _variant, price in priced_variants)
    sku, _is_real_sku = extract_sku(product, None, slug)
    distinct_prices = {price for _variant, price in priced_variants}

    return [
        CatalogRow(
            sku=sku,
            product_name=base_name,
            price_cad=lowest_price,
            category=category,
            availability=aggregate_status(variants),
            product_url=product_url(slug),
            price_type="From" if len(distinct_prices) > 1 else "Exact",
            updated_utc=timestamp,
        )
    ]


def deduplicate_rows(rows: list[CatalogRow]) -> list[CatalogRow]:
    by_key: dict[tuple[str, str], CatalogRow] = {}
    for row in rows:
        key = (row.sku.casefold(), row.product_url.casefold())
        by_key.setdefault(key, row)
    return sorted(
        by_key.values(),
        key=lambda row: (row.category.casefold(), row.product_name.casefold(), row.sku.casefold()),
    )


def validate_catalog(products: dict[str, dict[str, Any]], rows: list[CatalogRow]) -> None:
    errors: list[str] = []

    if len(products) < MIN_PRODUCTS:
        errors.append(f"only {len(products)} unique products; minimum is {MIN_PRODUCTS}")

    if not (KNOWN_PRODUCT_SLUGS & set(products)):
        errors.append("none of the known sanity-check product slugs were found")

    if not rows:
        errors.append("no priced rows were generated")
    else:
        priced_ratio = len({row.product_url for row in rows}) / max(1, len(products))
        if priced_ratio < MIN_PRICED_RATIO:
            errors.append(
                f"only {priced_ratio:.1%} of products produced a priced row; "
                f"minimum is {MIN_PRICED_RATIO:.0%}"
            )

    if any(row.price_cad <= 0 for row in rows):
        errors.append("one or more output prices are not positive")

    if errors:
        raise RuntimeError("Dataset validation failed: " + "; ".join(errors))


def atomic_write_csv(rows: list[CatalogRow], output_file: Path) -> None:
    output_file.parent.mkdir(parents=True, exist_ok=True)

    temp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            newline="",
            encoding="utf-8",
            prefix=f".{output_file.name}.",
            suffix=".tmp",
            dir=output_file.parent,
            delete=False,
        ) as handle:
            temp_path = Path(handle.name)
            writer = csv.DictWriter(handle, fieldnames=CSV_FIELDS)
            writer.writeheader()
            for row in rows:
                writer.writerow(row.as_csv_row())

        os.replace(temp_path, output_file)
    except Exception:
        if temp_path is not None:
            temp_path.unlink(missing_ok=True)
        raise


def run() -> None:
    timestamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    session = make_session()

    products = fetch_catalog(session)

    rows: list[CatalogRow] = []
    unpriced_slugs: list[str] = []
    for slug, product in products.items():
        product_rows = product_to_rows(product, timestamp)
        if product_rows:
            rows.extend(product_rows)
        else:
            unpriced_slugs.append(slug)

    rows = deduplicate_rows(rows)
    validate_catalog(products, rows)
    atomic_write_csv(rows, OUTPUT_FILE)

    real_sku_like = sum(1 for row in rows if row.sku != row.product_url.rsplit("/", 1)[-1])
    print(f"Success: wrote {len(rows)} rows to {OUTPUT_FILE}")
    print(f"Products without a usable price: {len(unpriced_slugs)}")
    print(f"Rows with a model/SKU instead of slug fallback: {real_sku_like}/{len(rows)}")

    if unpriced_slugs:
        preview = ", ".join(sorted(unpriced_slugs)[:12])
        print(f"Unpriced examples: {preview}")


if __name__ == "__main__":
    try:
        run()
    except Exception as exc:
        print(f"FATAL: {exc}", file=sys.stderr)
        print(
            "The existing unifi_prices.csv was left untouched unless a newly fetched "
            "dataset passed validation and completed the atomic replace.",
            file=sys.stderr,
        )
        raise SystemExit(1)
