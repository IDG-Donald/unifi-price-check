# Automated UniFi Canada Price Dataset

This repository maintains an Excel-friendly CSV of current Canadian UniFi Store pricing.

The data flow is intentionally simple:

```text
UniFi Store Next.js category JSON
        ↓
Python + requests
        ↓
unifi_prices.csv
        ↓
GitHub raw file
        ↓
Excel / Power Query
```

## Why this version is lightweight

The scraper does **not** open hundreds of product pages in Chromium.

A run normally performs:

- 1 request to the Canadian storefront homepage to discover the current Next.js `buildId`;
- 9 category JSON requests covering the UniFi catalog;
- at most one extra homepage/category request if the `buildId` rotates during the run.

The scheduled workflow runs once per day.

## Output

`unifi_prices.csv` contains:

| Column | Purpose |
|---|---|
| `SKU` | UniFi SKU/model when exposed by the catalog; otherwise the stable product slug |
| `Product Name` | Store product/variant name |
| `Price (CAD)` | Current Canadian display price |
| `Line/Category` | High-level UniFi category |
| `Availability` | In Stock, Out of Stock, Coming Soon, Preorder, or Unknown |
| `Product URL` | Canadian store product page |
| `Price Type` | `Exact` or `From` |
| `Updated UTC` | Timestamp of the successful fetch |

The original first five columns remain first for compatibility with the Excel/Power Query design.

## Excel / Power Query URL

Use:

```text
https://raw.githubusercontent.com/IDG-Donald/unifi-price-check/main/unifi_prices.csv
```

Example Power Query:

```powerquery
let
    Source = Csv.Document(
        Web.Contents(
            "https://raw.githubusercontent.com/IDG-Donald/unifi-price-check/main/unifi_prices.csv"
        ),
        [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]
    ),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Types = Table.TransformColumnTypes(
        Headers,
        {
            {"SKU", type text},
            {"Product Name", type text},
            {"Price (CAD)", Currency.Type},
            {"Line/Category", type text},
            {"Availability", type text},
            {"Product URL", type text},
            {"Price Type", type text},
            {"Updated UTC", type datetimezone}
        }
    )
in
    Types
```

## How the fetch works

`get_unifi_prices.py`:

1. Fetches `https://ca.store.ui.com/ca/en`, extracts the current Next.js `buildId`, and records the effective storefront origin.
2. Requests the JSON backing nine broad UniFi category pages.
3. Reads products from `pageProps.subCategories[].products`, with a flat `pageProps.products` fallback.
4. Deduplicates products by their store slug.
5. Reads prices from variant `displayPrice`, falling back to `price`.
6. Reads availability from variant `status`.
7. Emits separate rows for distinct priced variants when the store provides distinct variant SKUs; otherwise it emits one product row using the lowest displayed price and marks it `From` when appropriate.
8. Validates the full dataset.
9. Atomically replaces `unifi_prices.csv` only after validation succeeds.

The current endpoint pattern and JSON structure are based on the public storefront behavior also used by the open-source `jamesccupps/UnifiStockWatcher` project.

## Failure behavior

The previous good CSV is preserved if the fetch fails.

The script refuses to publish when, for example:

- too few unique products are returned;
- none of several known UniFi product slugs are present;
- too few catalog products have usable prices;
- a broad category returns no products;
- a price is non-positive;
- the storefront stops returning the expected Next.js JSON shape.

A failed run exits non-zero so GitHub Actions shows the problem instead of committing an `ERROR` row.

## Local test

Python 3.12 is recommended.

### Windows PowerShell

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python get_unifi_prices.py
```

### Linux / macOS

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python get_unifi_prices.py
```

A healthy run should show roughly this pattern:

```text
Store buildId: ...
Fetching category 1/9: Cloud Gateways
  received ... product records; ... new unique slugs
...
Fetched ... unique products from 9 category requests
Success: wrote ... rows to unifi_prices.csv
```

It should finish vastly faster than the browser-based version because it is no longer loading every product page.

## Configuration

Optional environment variables:

| Variable | Default | Purpose |
|---|---:|---|
| `UNIFI_OUTPUT_FILE` | `unifi_prices.csv` | Output CSV path |
| `UNIFI_MIN_PRODUCTS` | `100` | Minimum unique catalog products required |
| `UNIFI_MIN_PRICED_RATIO` | `0.70` | Minimum fraction of catalog products that must yield a price |
| `UNIFI_REQUEST_TIMEOUT` | `20` | HTTP timeout in seconds |
| `UNIFI_CATEGORY_DELAY` | `0.25` | Delay between category calls |
| `UNIFI_STORE_ENTRY` | `https://ca.store.ui.com` | Canadian storefront used to discover the build ID and effective JSON host |
| `UNIFI_DISPLAY_BASE` | `https://ca.store.ui.com` | Host used in generated product links |
| `UNIFI_REGION_PATH` | `ca/en` | Store locale/region path |

## GitHub Actions

`.github/workflows/unifi-tracker.yml` runs once daily at 06:00 UTC and can also be run manually.

Because `Updated UTC` records each successful refresh, a normal successful daily run will produce a fresh CSV commit even when prices themselves did not change.

## Notes

This project relies on an undocumented public storefront data shape, not a documented retail-pricing API. Ubiquiti can change the Next.js routes or JSON schema, so validation is deliberately strict enough to fail visibly rather than silently publishing a bad price list.
