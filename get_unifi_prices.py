import csv
import json
import requests

# URL for the UI store's internal client-side product query catalog
STORE_API_URL = "https://ui.com"
OUTPUT_FILE = "unifi_prices.csv"

# Browser simulation headers to bypass basic data-center blocking blocks
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "application/json",
    "Referer": "https://store.ui.com/"
}

def fetch_and_parse_catalog():
    print(f"Connecting to Ubiquiti Store data engine...")
    try:
        response = requests.get(STORE_API_URL, headers=HEADERS, timeout=15)
        response.raise_for_status()
        data = response.json()
    except Exception as e:
        print(f"Error fetching data: {e}")
        return

    # Extract products array from response payload
    products = data.get("products", [])
    if not products:
        print("No products found in the response object.")
        return

    print(f"Successfully located {len(products)} total catalog items. Processing tables...")

    # Write data cleanly to a structural CSV
    with open(OUTPUT_FILE, mode="w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        # Header Row for Excel Power Query
        writer.writerow(["SKU", "Product Name", "Price (USD)", "Line/Category", "Availability"])

        for item in products:
            sku = item.get("sku", "N/A")
            name = item.get("title", "Unknown Product")
            
            # Extract standard pricing (converts cents to decimal format)
            price_raw = item.get("price", 0)
            price_usd = round(float(price_raw) / 100, 2) if price_raw else 0.00
            
            category = item.get("product_line", "General")
            
            # Simple stock translation flag
            in_stock = "In Stock" if item.get("is_in_stock", False) else "Out of Stock"

            writer.writerow([sku, name, price_usd, category, in_stock])

    print(f"File successfully created: {OUTPUT_FILE}")

if __name__ == "__main__":
    fetch_and_parse_catalog()
