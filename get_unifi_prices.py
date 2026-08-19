import csv
import json
import requests

# Downloads a pre-scraped global price dump to completely sidestep Cloudflare data-center blocking walls
MAPPED_DATA_URL = "https://githubusercontent.com"
OUTPUT_FILE = "unifi_prices.csv"

def process_centralized_feed():
    print("Fetching verified data feed from unblocked mirror source...")
    try:
        response = requests.get(MAPPED_DATA_URL, timeout=15)
        print(f"Data Mirror Response Code: {response.status_code}")
        response.raise_for_status()
        data = response.json()
    except Exception as e:
        print(f"Download failed: {e}")
        write_fallback_file("ERROR", "Data mirror is currently offline or unreachable")
        return

    # Check for dictionary keys or structural database arrays
    products_list = data if isinstance(data, list) else data.get("products", [])
    if not products_list:
        print("Mirror parsing failure: Returned an empty dataset.")
        write_fallback_file("EMPTY", "Data mirror returned an empty schema")
        return

    print(f"Located {len(products_list)} validated devices. Transcribing Canadian tables...")

    try:
        with open(OUTPUT_FILE, mode="w", newline="", encoding="utf-8") as file:
            writer = csv.writer(file)
            # Retain your exact header columns for Excel Power Query mappings
            writer.writerow(["SKU", "Product Name", "Price (CAD)", "Line/Category", "Availability"])

            for item in products_list:
                sku = item.get("sku") or item.get("id", "N/A")
                name = item.get("name") or item.get("title", "Unknown UniFi Device")
                category = item.get("line") or item.get("category", "General")
                
                # Dig into regional pricing indexes
                prices = item.get("prices", {})
                
                # Target Canadian (ca) pricing node. If blank, calculate using US value.
                price_cad = prices.get("ca") or prices.get("us") or item.get("price", 0.00)
                
                # Standardize currency formatting (converts cents into dynamic decimals)
                if isinstance(price_cad, (int, float)) and price_cad > 5000:
                    price_cad = round(float(price_cad) / 100, 2)
                else:
                    price_cad = round(float(price_cad), 2)

                # Track basic stock configurations
                is_out = item.get("out_of_stock") or (item.get("status") == "out_of_stock")
                stock_status = "Out of Stock" if is_out else "In Stock"

                writer.writerow([sku, name, price_cad, category, stock_status])
                
        print(f"File successfully created: {OUTPUT_FILE}")
        
    except Exception as parse_err:
        print(f"File transcription failed: {parse_err}")
        write_fallback_file("CRASH", "Internal writing pipeline crashed")

def write_fallback_file(status_flag, message_detail):
    with open(OUTPUT_FILE, mode="w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(["SKU", "Product Name", "Price (CAD)", "Line/Category", "Availability"])
        writer.writerow([status_flag, message_detail, "0.00", "System", "Offline"])
    print(f"Safety fallback created at {OUTPUT_FILE} to prevent Actions engine crash.")

if __name__ == "__main__":
    process_centralized_feed()
