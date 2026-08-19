import csv
import json
import requests

# Targets a Cloudflare-bypassed third-party tracking index scraping the Canadian (CA) UI storefront
STORE_API_URL = "https://trackalacker.com"
OUTPUT_FILE = "unifi_prices.csv"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept": "application/json"
}

def fetch_and_parse_catalog():
    print("Connecting to Cloudflare-Bypassed UniFi CA Data Feed...")
    try:
        response = requests.get(STORE_API_URL, headers=HEADERS, timeout=25)
        print(f"Server Response Status: {response.status_code}")
        response.raise_for_status()
        data = response.json()
    except Exception as e:
        print(f"Network error reading proxy: {e}")
        write_fallback_file("ERROR", "Failed to reach bypass database feed")
        return

    # Unpack TrackaLacker's product catalog array mapping
    products = data if isinstance(data, list) else data.get("products", [])
    if not products:
        print("Data parsing error: JSON index returned blank schema.")
        write_fallback_file("EMPTY", "Data stream index returned no objects")
        return

    print(f"Successfully processed {len(products)} live items. Transcribing Canadian tables...")

    try:
        with open(OUTPUT_FILE, mode="w", newline="", encoding="utf-8") as file:
            writer = csv.writer(file)
            # Retain your original template header structure for seamless Excel Power Query compatibility
            writer.writerow(["SKU", "Product Name", "Price (CAD)", "Line/Category", "Availability"])

            for item in products:
                # Handle variants and extra nested dictionaries inside TrackaLacker objects
                sku = item.get("sku") or item.get("model", "N/A")
                name = item.get("name") or item.get("title", "Unknown UniFi Device")
                category = item.get("category") or item.get("product_line", "General")
                
                # Fetch pricing (TrackaLacker serves raw float integers already set to CAD when using locale=ca)
                price_cad = item.get("price") or item.get("msrp", 0.00)
                price_cad = round(float(price_cad), 2)

                # Check structural stock indicators
                is_available = item.get("in_stock") or item.get("available", False)
                stock_status = "In Stock" if is_available else "Out of Stock"

                writer.writerow([sku, name, price_cad, category, stock_status])
                
        print(f"File successfully created: {OUTPUT_FILE}")
        
    except Exception as parse_err:
        print(f"File write failure: {parse_err}")
        write_fallback_file("CRASH", "Internal writing pipeline crashed")

def write_fallback_file(status_flag, message_detail):
    with open(OUTPUT_FILE, mode="w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(["SKU", "Product Name", "Price (CAD)", "Line/Category", "Availability"])
        writer.writerow([status_flag, message_detail, "0.00", "System", "Offline"])
    print(f"Safety fallback created at {OUTPUT_FILE} to prevent Action engine crash.")

if __name__ == "__main__":
    fetch_and_parse_catalog()
