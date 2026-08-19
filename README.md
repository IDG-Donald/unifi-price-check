# Automated UniFi Store Price Dataset

This repository hosts a self-updating dataset tracking retail hardware values on the official Ubiquiti Store.

## Data Source Access
The updated price catalog is saved directly within this repository as a structural CSV file. You can hook your workflows directly to the live URL:

* **Live CSV File Path:** `unifi_prices.csv`
* **Raw Excel-Ready Web Link:** `https://githubusercontent.com`

## Automation Strategy
The scraping deployment runs entirely on [GitHub Actions](https://github.com) completely free of charge.
* **Frequency:** Triggers daily at 06:00 UTC via a native crontab job.
* **Output Engine:** Generates structural arrays mapping SKUs, Descriptions, Category definitions, and pricing metrics.
