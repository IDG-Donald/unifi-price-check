#!/usr/bin/env python3
"""Fetch current UniFi Canada catalog pricing into an Excel-friendly CSV.

This version uses the public JSON payloads that back the UniFi Store's Next.js
category pages for product discovery, then fetches one normal HTML category page
per logical category to capture the displayed base / "Surcharge incl." price pair.
It does not launch a browser or request every product page.

Failure behavior is intentionally conservative: a new CSV is written to a
temporary file, validated, and only then atomically replaces the previous CSV.
"""

from __future__ import annotations

import csv
import html
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
from bs4 import BeautifulSoup
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
# Each category can have one or more candidate routes. UniFi occasionally
# renames storefront category slugs while leaving the underlying JSON format
# unchanged. The scraper tries candidates in order and uses the first one that
# returns products.
CATEGORY_ROUTES: dict[str, tuple[str, ...]] = {
    "Cloud Gateways": ("category/all-cloud-gateways",),
    "Switching": ("category/all-switching",),
    "WiFi": ("category/all-wifi",),
    # Current CA route first; legacy route retained as a fallback.
    "Physical Security": (
        "category/all-physical-security",
        "category/all-cameras-nvrs",
    ),
    "Door Access": ("category/all-door-access",),
    "Integrations": ("category/all-integrations",),
    "Advanced Hosting": ("category/all-advanced-hosting",),
    "Accessories": ("category/accessories-cables-dacs",),
    "Network Storage": ("category/network-storage",),
}

# Advanced Hosting is a storefront landing page rather than a unique hardware
# catalog. Its hardware cards substantially overlap Cloud Gateways/Integrations,
# and its JSON layout has changed independently of the normal category pages.
# A failure here is logged but should not block an otherwise complete hardware
# price refresh.
OPTIONAL_CATEGORIES = {"Advanced Hosting"}

CSV_FIELDS = [
    "SKU",
    "Product Name",
    "Store Description",
    # Backward-compatible quote price. This remains the price Excel should use.
    "Price (CAD)",
    "Base Price (CAD)",
    "Surcharge Price (CAD)",
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
HTML_TAG_RE = re.compile(r"<[^>]+>")
WHITESPACE_RE = re.compile(r"\s+")
SURCHARGE_PAIR_RE = re.compile(
    r"(?P<from>\bFrom\s+)?\$\s*(?P<base>[0-9][0-9,]*(?:\.[0-9]{2})?)"
    r"\s*\$\s*(?P<surcharge>[0-9][0-9,]*(?:\.[0-9]{2})?)"
    r"\s*Surcharge\s*incl\.?",
    re.IGNORECASE,
)
SINGLE_PRICE_RE = re.compile(
    r"(?P<from>\bFrom\s+)?\$\s*(?P<price>[0-9][0-9,]*(?:\.[0-9]{2})?)",
    re.IGNORECASE,
)

# UniFi's category payloads have used several description-ish keys over time.
# Prefer the short customer-readable summary shown on category/product cards.
DESCRIPTION_KEYS = (
    "description",
    "shortDescription",
    "short_description",
    "productDescription",
    "summary",
    "subtitle",
    "tagline",
)

# The fast Next.js catalog payload is retained as the source for product
# discovery, SKU, stock, description, and a usable current price. It does NOT
# reliably expose the Canadian memory surcharge as a separate field.
#
# The true base/surcharge pair is therefore parsed from the rendered category
# HTML, where UniFi explicitly displays e.g. "$400.00 $431.00 Surcharge incl.".
JSON_PRICE_KEYS = (
    "displayPrice",
    "price",
    "currentPrice",
    "finalPrice",
    "surchargePrice",
    "priceWithSurcharge",
    "basePrice",
    "priceBeforeSurcharge",
)


@dataclass(frozen=True)
class StorefrontPrice:
    base_price_cad: Decimal
    surcharge_price_cad: Decimal | None
    price_type: str

    @property
    def quote_price_cad(self) -> Decimal:
        return self.surcharge_price_cad or self.base_price_cad


@dataclass(frozen=True)
class CatalogRow:
    sku: str
    product_name: str
    store_description: str
    price_cad: Decimal
    base_price_cad: Decimal
    surcharge_price_cad: Decimal | None
    category: str
    availability: str
    product_url: str
    price_type: str
    updated_utc: str

    def as_csv_row(self) -> dict[str, str]:
        return {
            "SKU": self.sku,
            "Product Name": self.product_name,
            "Store Description": self.store_description,
            "Price (CAD)": f"{self.price_cad:.2f}",
            "Base Price (CAD)": f"{self.base_price_cad:.2f}",
            "Surcharge Price (CAD)": (
                f"{self.surcharge_price_cad:.2f}"
                if self.surcharge_price_cad is not None
                else ""
            ),
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


def category_page_url(store_origin: str, route: str) -> str:
    return f"{store_origin}/{REGION_PATH}/{route}"


def slug_from_product_href(href: str) -> str:
    path = urlsplit(href).path.rstrip("/")
    marker = "/products/"
    if marker not in path:
        return ""
    return path.split(marker, 1)[1].split("/", 1)[0].strip()


def decimal_from_price_text(value: str) -> Decimal | None:
    try:
        return Decimal(value.replace(",", "")).quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError):
        return None


def parse_storefront_price_text(text: str) -> StorefrontPrice | None:
    """Parse the price wording used on UniFi category cards.

    Examples:
      $400.00 $431.00 Surcharge incl.
      From $239.00 $257.00 Surcharge incl.
      $199.00
    """
    normalized = WHITESPACE_RE.sub(" ", html.unescape(text)).strip()

    pair = SURCHARGE_PAIR_RE.search(normalized)
    if pair:
        base = decimal_from_price_text(pair.group("base"))
        surcharge = decimal_from_price_text(pair.group("surcharge"))
        if base is not None and surcharge is not None and base > 0 and surcharge > 0:
            return StorefrontPrice(
                base_price_cad=base,
                surcharge_price_cad=surcharge,
                price_type="From" if pair.group("from") else "Exact",
            )

    single = SINGLE_PRICE_RE.search(normalized)
    if single:
        price = decimal_from_price_text(single.group("price"))
        if price is not None and price > 0:
            return StorefrontPrice(
                base_price_cad=price,
                surcharge_price_cad=None,
                price_type="From" if single.group("from") else "Exact",
            )

    return None


def nearest_product_card_price(anchor: Any) -> StorefrontPrice | None:
    """Walk upward from a product link until a compact price-bearing card is found."""
    node = anchor
    for _ in range(10):
        node = getattr(node, "parent", None)
        if node is None:
            break

        # Category cards are compact. Refuse very large ancestors so a match
        # cannot accidentally borrow a neighbouring product's price.
        text = " ".join(node.stripped_strings)
        if not text or len(text) > 5000:
            continue

        parsed = parse_storefront_price_text(text)
        if parsed is not None:
            return parsed

    return None


def parse_category_html_prices(
    html_text: str, known_slugs: set[str]
) -> dict[str, StorefrontPrice]:
    """Return displayed base/surcharge prices keyed by product slug."""
    soup = BeautifulSoup(html_text, "html.parser")
    prices: dict[str, StorefrontPrice] = {}

    for anchor in soup.find_all("a", href=True):
        slug = slug_from_product_href(str(anchor.get("href") or ""))
        if not slug or slug not in known_slugs or slug in prices:
            continue

        parsed = nearest_product_card_price(anchor)
        if parsed is not None:
            prices[slug] = parsed

    return prices


def fetch_category_html_prices(
    session: requests.Session,
    store_origin: str,
    route: str,
    products: list[dict[str, Any]],
) -> dict[str, StorefrontPrice]:
    """Fetch one rendered category HTML page and parse visible price pairs."""
    known_slugs = {str(p.get("slug") or "").strip() for p in products}
    known_slugs.discard("")
    if not known_slugs:
        return {}

    url = category_page_url(store_origin, route)
    response = session.get(url, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()

    content_type = response.headers.get("content-type", "").lower()
    if "html" not in content_type:
        raise RuntimeError(
            f"category page did not return HTML "
            f"(status={response.status_code}, content-type={content_type!r}, url={response.url})"
        )

    return parse_category_html_prices(response.text, known_slugs)


def looks_like_product(obj: dict[str, Any]) -> bool:
    """Return True for storefront objects that look like product cards.

    Most category pages expose products under pageProps.subCategories[].products,
    but some landing pages (notably Advanced Hosting) nest product cards deeper in
    sections/collections. Requiring a slug plus product-ish fields avoids treating
    arbitrary nested metadata as a product.
    """
    slug = str(obj.get("slug") or "").strip()
    if not slug:
        return False

    return any(
        key in obj
        for key in (
            "variants",
            "displayPrice",
            "price",
            "title",
            "name",
            "displayName",
            "productName",
        )
    )


def iter_products_from_page_props(page_props: dict[str, Any]) -> Iterable[dict[str, Any]]:
    """Recursively yield product objects from a category's pageProps.

    UniFi does not use one stable nesting shape for every category page. Walking
    pageProps recursively keeps the scraper independent of presentation sections
    while still limiting discovery to product-like dictionaries. Duplicate slugs
    are removed later by fetch_catalog().
    """
    stack: list[Any] = [page_props]
    seen_objects: set[int] = set()

    while stack:
        obj = stack.pop()

        if isinstance(obj, dict):
            object_id = id(obj)
            if object_id in seen_objects:
                continue
            seen_objects.add(object_id)

            if looks_like_product(obj):
                yield obj

            stack.extend(obj.values())

        elif isinstance(obj, list):
            stack.extend(obj)


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
    """Fetch JSON catalog plus one HTML price page per logical category."""
    build_id, store_origin = get_store_context(session)
    print(f"Store origin: {store_origin}")
    print(f"Store buildId: {build_id}")

    by_slug: dict[str, dict[str, Any]] = {}

    for index, (category_label, routes) in enumerate(CATEGORY_ROUTES.items(), start=1):
        print(f"Fetching category {index}/{len(CATEGORY_ROUTES)}: {category_label}")

        products: list[dict[str, Any]] | None = None
        chosen_route: str | None = None
        route_errors: list[str] = []

        for route_index, route in enumerate(routes, start=1):
            try:
                candidate_products, build_id, store_origin = fetch_category(
                    session=session,
                    store_origin=store_origin,
                    build_id=build_id,
                    category_label=category_label,
                    route=route,
                )
                products = candidate_products
                chosen_route = route
                if route_index > 1:
                    print(f"  using fallback route: {route}")
                break
            except RuntimeError as exc:
                route_errors.append(f"{route}: {exc}")
                if route_index < len(routes):
                    print(f"  route {route} returned no usable catalog; trying fallback")
                    continue

                if category_label in OPTIONAL_CATEGORIES:
                    print(
                        f"  WARNING: optional category {category_label} could not be parsed; "
                        "continuing because its hardware overlaps other categories"
                    )
                    print(f"  details: {' | '.join(route_errors)}")
                    products = []
                    break

                raise RuntimeError(
                    f"{category_label} failed on every known route: "
                    + " | ".join(route_errors)
                ) from exc

        if products is None:
            raise RuntimeError(f"{category_label} produced no product list")

        html_prices: dict[str, StorefrontPrice] = {}
        if products and chosen_route:
            try:
                html_prices = fetch_category_html_prices(
                    session=session,
                    store_origin=store_origin,
                    route=chosen_route,
                    products=products,
                )
                split_count = sum(
                    1 for price in html_prices.values() if price.surcharge_price_cad is not None
                )
                print(
                    f"  HTML pricing matched {len(html_prices)} product slugs; "
                    f"{split_count} include a surcharge pair"
                )
            except Exception as exc:
                # The JSON feed still supplies a usable current quote price. A
                # transient HTML/layout problem therefore leaves surcharge fields
                # blank instead of blocking the entire daily price refresh.
                print(f"  WARNING: storefront HTML pricing overlay failed: {exc}")

        category_new = 0
        for product in products:
            slug = str(product.get("slug") or "").strip()
            if not slug:
                continue

            incoming = dict(product)
            incoming["_category"] = category_label
            if slug in html_prices:
                incoming["_storefront_price"] = html_prices[slug]

            if slug not in by_slug:
                by_slug[slug] = incoming
                category_new += 1
            else:
                # A product may appear in several logical categories. Preserve
                # the first catalog record, but enrich it if a later category
                # provided the visible base/surcharge pair.
                if (
                    "_storefront_price" not in by_slug[slug]
                    and "_storefront_price" in incoming
                ):
                    by_slug[slug]["_storefront_price"] = incoming["_storefront_price"]

        print(
            f"  received {len(products)} product records; "
            f"{category_new} new unique slugs"
        )

        if CATEGORY_DELAY_SECONDS > 0 and index < len(CATEGORY_ROUTES):
            time.sleep(CATEGORY_DELAY_SECONDS)

    print(
        f"Fetched {len(by_slug)} unique products from "
        f"{len(CATEGORY_ROUTES)} logical categories"
    )
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


def first_money(mapping: dict[str, Any], keys: Iterable[str]) -> Decimal | None:
    for key in keys:
        value = money_to_decimal(mapping.get(key))
        if value is not None:
            return value
    return None


def json_quote_price(mapping: dict[str, Any]) -> Decimal | None:
    """Return the best usable current price from the Next.js JSON payload.

    This is deliberately a single price. We no longer infer base vs surcharge
    from JSON fields because the current Canadian payload can expose the same
    effective storefront price in both `price` and `displayPrice`.
    """
    return first_money(mapping, JSON_PRICE_KEYS)


def storefront_price(product: dict[str, Any]) -> StorefrontPrice | None:
    value = product.get("_storefront_price")
    return value if isinstance(value, StorefrontPrice) else None



def clean_description_text(value: Any) -> str:
    """Convert simple storefront description payloads to one-line plain text."""
    if value is None:
        return ""

    if isinstance(value, str):
        text = html.unescape(HTML_TAG_RE.sub(" ", value))
        return WHITESPACE_RE.sub(" ", text).strip()

    # Rich-text payloads commonly keep readable text under one of these keys.
    if isinstance(value, dict):
        for key in ("text", "value", "content", "plainText"):
            if key in value:
                text = clean_description_text(value.get(key))
                if text:
                    return text
        return ""

    if isinstance(value, list):
        parts = [clean_description_text(item) for item in value]
        return " ".join(part for part in parts if part).strip()

    return ""


def product_description(product: dict[str, Any]) -> str:
    # Fast path: most category cards keep the description directly on product.
    for key in DESCRIPTION_KEYS:
        text = clean_description_text(product.get(key))
        if text:
            return text

    # Some category layouts wrap card copy in a nested presentation object. Walk
    # only for known description keys rather than grabbing arbitrary text fields.
    stack: list[Any] = list(product.values())
    while stack:
        obj = stack.pop()
        if isinstance(obj, dict):
            for key in DESCRIPTION_KEYS:
                if key in obj:
                    text = clean_description_text(obj.get(key))
                    if text:
                        return text
            stack.extend(obj.values())
        elif isinstance(obj, list):
            stack.extend(obj)

    return ""


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
    description = product_description(product)
    visible_price = storefront_price(product)
    variants = [v for v in (product.get("variants") or []) if isinstance(v, dict)]

    priced_variants: list[tuple[dict[str, Any], Decimal]] = []
    for variant in variants:
        quote_price = json_quote_price(variant)
        if quote_price is None or quote_price <= 0:
            continue
        priced_variants.append((variant, quote_price))

    if not priced_variants:
        # Some products carry price directly on the product object. Prefer the
        # visible category-card price pair whenever available; otherwise retain
        # the JSON quote price and leave the surcharge field blank.
        json_price = json_quote_price(product)
        if visible_price is not None:
            base_price = visible_price.base_price_cad
            surcharge_price = visible_price.surcharge_price_cad
            quote_price = visible_price.quote_price_cad
            price_type = visible_price.price_type
        elif json_price is not None and json_price > 0:
            base_price = json_price
            surcharge_price = None
            quote_price = json_price
            price_type = "Exact"
        else:
            return []

        sku, _is_real_sku = extract_sku(product, None, slug)
        return [
            CatalogRow(
                sku=sku,
                product_name=base_name,
                store_description=description,
                price_cad=quote_price,
                base_price_cad=base_price,
                surcharge_price_cad=surcharge_price,
                category=category,
                availability=aggregate_status(variants) if variants else "Unknown",
                product_url=product_url(slug),
                price_type=price_type,
                updated_utc=timestamp,
            )
        ]

    # If multiple variants expose distinct real SKUs, keep one row per SKU.
    # A single category card cannot safely describe differently priced variants,
    # so we apply its split price only when the card is Exact and all JSON
    # variants currently agree on the same quote price.
    distinct_json_prices = {item[1] for item in priced_variants}
    safe_exact_overlay = (
        visible_price is not None
        and visible_price.price_type == "Exact"
        and len(distinct_json_prices) == 1
    )

    distinct_variant_rows: list[CatalogRow] = []
    seen_variant_skus: set[str] = set()
    real_variant_sku_count = 0

    for variant, json_price in priced_variants:
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

        if safe_exact_overlay and visible_price is not None:
            base_price = visible_price.base_price_cad
            surcharge_price = visible_price.surcharge_price_cad
            quote_price = visible_price.quote_price_cad
        else:
            # JSON remains authoritative for variant-specific quote pricing. Do
            # not invent a base/surcharge split when the category card only says
            # "From" or variants have genuinely different prices.
            base_price = json_price
            surcharge_price = None
            quote_price = json_price

        distinct_variant_rows.append(
            CatalogRow(
                sku=sku,
                product_name=name,
                store_description=description,
                price_cad=quote_price,
                base_price_cad=base_price,
                surcharge_price_cad=surcharge_price,
                category=category,
                availability=normalize_status(variant.get("status")),
                product_url=product_url(slug),
                price_type="Exact",
                updated_utc=timestamp,
            )
        )

    if real_variant_sku_count == len(priced_variants):
        return distinct_variant_rows

    # Otherwise expose one product-level row. This is exactly where the category
    # card's "From" base/surcharge pair is most useful.
    if visible_price is not None:
        base_price = visible_price.base_price_cad
        surcharge_price = visible_price.surcharge_price_cad
        quote_price = visible_price.quote_price_cad
        price_type = visible_price.price_type
    else:
        _variant, quote_price = min(priced_variants, key=lambda item: item[1])
        base_price = quote_price
        surcharge_price = None
        price_type = "From" if len(distinct_json_prices) > 1 else "Exact"

    sku, _is_real_sku = extract_sku(product, None, slug)
    return [
        CatalogRow(
            sku=sku,
            product_name=base_name,
            store_description=description,
            price_cad=quote_price,
            base_price_cad=base_price,
            surcharge_price_cad=surcharge_price,
            category=category,
            availability=aggregate_status(variants),
            product_url=product_url(slug),
            price_type=price_type,
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
        errors.append("one or more output quote prices are not positive")

    if any(row.base_price_cad <= 0 for row in rows):
        errors.append("one or more output base prices are not positive")

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
    surcharge_rows = sum(
        1
        for row in rows
        if row.surcharge_price_cad is not None
        and row.surcharge_price_cad != row.base_price_cad
    )
    descriptions = sum(1 for row in rows if row.store_description)
    html_price_products = sum(1 for product in products.values() if storefront_price(product))

    print(f"Success: wrote {len(rows)} rows to {OUTPUT_FILE}")
    print(f"Products without a usable price: {len(unpriced_slugs)}")
    print(f"Rows with a model/SKU instead of slug fallback: {real_sku_like}/{len(rows)}")
    print(f"Rows with a true base/surcharge split: {surcharge_rows}/{len(rows)}")
    print(f"Products matched to storefront HTML pricing: {html_price_products}/{len(products)}")
    print(f"Rows with a store description: {descriptions}/{len(rows)}")

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
