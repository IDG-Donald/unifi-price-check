# Automated UniFi Canada Price Dataset

This repository maintains an Excel-friendly CSV containing current hardware pricing from the official Canadian Ubiquiti UniFi Store.

The intended data flow is:

```text
Canadian UniFi Store
        ↓
Playwright scraper
        ↓
unifi_prices.csv
        ↓
GitHub raw file
        ↓
Excel / Power Query
```

## Output

The scraper writes `unifi_prices.csv` with these columns:

| Column | Purpose |
|---|---|
| `SKU` | UniFi model/SKU used as the lookup key in Excel |
| `Product Name` | Store product name |
| `Price (CAD)` | Current displayed Canadian price |
| `Line/Category` | High-level UniFi product category |
| `Availability` | In Stock, Out of Stock, Preorder, Selectable, or Unknown |
| `Product URL` | Source product page |
| `Price Type` | `Exact` or `From` |
| `Updated UTC` | Timestamp of the successful scrape |

The original first five columns are retained so existing Excel/Power Query mappings can continue to use them.

## Excel / Power Query URL

Use the raw GitHub CSV URL:

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

In Excel, load the query as a table and use `SKU` as the stable lookup key from estimating/quotation sheets.

## How the scraper works

`get_unifi_prices.py`:

1. Opens the major Canadian UniFi Store category pages in Chromium using Playwright.
2. Discovers product and product-collection links from the rendered store.
3. Expands collection pages to discover additional individual products.
4. Visits each product page.
5. Prefers structured JSON-LD/meta product data and falls back to visible page text when required.
6. Deduplicates products by SKU.
7. Validates the result before publishing it.
8. Replaces `unifi_prices.csv` atomically only after validation succeeds.

A scraping failure exits with a non-zero status. It does **not** deliberately overwrite the previous good CSV with an `ERROR` row.

## GitHub Actions

`.github/workflows/unifi-tracker.yml` runs every day at 06:00 UTC and can also be started manually from the Actions tab.

On a successful scrape it commits `unifi_prices.csv` only when the file changed.

## Local test

Python 3.12 is recommended.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m playwright install chromium
python get_unifi_prices.py
```

On Windows PowerShell:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m playwright install chromium
python get_unifi_prices.py
```

For troubleshooting with a visible browser:

```powershell
$env:UNIFI_HEADLESS="0"
python get_unifi_prices.py
```

## If GitHub-hosted runners are blocked

Ubiquiti may challenge or block traffic from cloud/datacenter IP ranges. If the workflow consistently receives an anti-bot page while the script works from your own network, use a GitHub self-hosted runner.

Change:

```yaml
runs-on: ubuntu-latest
```

to:

```yaml
runs-on: [self-hosted, linux, x64]
```

The rest of the workflow can remain the same.

## Safety / failure behavior

The scraper deliberately refuses to publish obviously bad datasets. Current checks include:

- minimum product count;
- minimum successful parse ratio;
- presence of at least one known UniFi SKU;
- excessive duplicate-SKU detection;
- positive product prices.

If UniFi changes its site structure, the workflow should fail visibly while leaving the last successfully generated CSV in place.

## Configuration

Optional environment variables:

| Variable | Default | Purpose |
|---|---:|---|
| `UNIFI_OUTPUT_FILE` | `unifi_prices.csv` | Output path |
| `UNIFI_MIN_PRODUCTS` | `40` | Minimum valid rows required |
| `UNIFI_HEADLESS` | `1` | Set to `0` for a visible browser |
| `UNIFI_NAV_TIMEOUT_MS` | `60000` | Navigation timeout per page |

## Important note

This project depends on the current public presentation of the Ubiquiti Store. It does not use a documented public Ubiquiti retail-pricing API. Site changes or anti-bot controls can therefore require maintenance.
