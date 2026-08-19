import csv
import json
import requests

# Explicitly target the Canadian (CA) regional store endpoint
STORE_API_URL = "https://ui.com"
OUTPUT_FILE = "unifi_prices.csv"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-CA,en-US,en;q=0.9",
    "Referer": "https://ui.com",
    "Origin": "https://store.ui.com"
}

def fetch_and_parse_catalog():
    print("Connecting to Ubiquiti Canadian Store data engine...")
    try:
        response = requests.get(STORE_API_URL, headers=HEADERS, timeout=15)
        print(f"Server responded with Status Code: {response.status_code}")
        
        if response.status_code != 200:
            print("Response preview:")
            print(response.text[:500])
        
        response.raise_for_status()
        data = response.json()
    except Exception as e:
        print(f"Error fetching data: {e}")
        # Create fallback file so Git doesn't crash the workflow
        with open(OUTPUT_FILE, mode="w", newline="", encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerow(["SKU", "Product Name", "Price (CAD)", "Line/Category", "Availability"])
            writer.writerow(["ERROR", "Could not reach store API", "0.00", "System", "Offline"])
        print(f"Created fallback error file at {OUTPUT_FILE} to avoid Git crash.")
        return

    products = data.get("products", [])
    if not products:
        print("No products found in the response object. Cloudflare or Geo-routing likely blocked payload.")
        with open(OUTPUT_FILE, mode="w", newline="", encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerow(["SKU", "Product Name", "Price (CAD)", "Line/Category", "Availability"])
            writer.writerow(["EMPTY", "No products returned by API", "0.00", "System", "No Data"])
        return

    print(f"Successfully located {len(products)} total catalog items. Processing tables...")

    with open(OUTPUT_FILE, mode="w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        # Updated header to reflect CAD tracking
        writer.writerow(["SKU", "Product Name", "Price (CAD)", "Line/Category", "Availability"])

        for item in products:
            sku = item.get("sku", "N/A")
            name = item.get("title", "Unknown Product")
            
            price_raw = item.get("price", 0)
            price_cad = round(float(price_raw) / 100, 2) if price_raw else 0.00
            category = item.get("product_line", "General")
            in_stock = "In Stock" if item.get("is_in_stock", False) else "Out of Stock"

            writer.writerow([sku, name, price_cad, category, in_stock])

    print(f"File successfully created with Canadian metrics: {OUTPUT_FILE}")

if __name__ == "__main__":
    fetch_and_parse_catalog()
