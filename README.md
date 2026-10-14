# ecommerce-automation
# E-Commerce Data Automation & Web Scraping Suite

A robust set of Python automation scripts designed to parse, synchronize, and enrich large-scale product inventories across inventory management systems (BenimPOS) and structured Excel spreadsheets without layout corruption.

## Overview
This repository contains production-ready automation scripts built to eliminate manual data entry errors and streamline inventory management for e-commerce operations. The suite handles barcode normalization, dynamic pricing & VAT updates, and automated high-resolution image scraping directly embedded into Excel cells.

---

## Key Modules & Features

### 1. KDV & Price Synchronization (`kdv_aktar.py`)
- Cross-references external inventory exports with target operational spreadsheets using intelligent barcode matching.
- Automatically normalizes barcodes (stripping hidden characters, trailing spaces, and leading zeros).
- Synchronizes and updates **VAT rates (KDV)** and pricing columns dynamically while generating discrepancy reports (`eslesmeyen_barkodlar.xlsx`) for unmatched items.

### 2. Secondary Price Injection (`fiyat2.py`)
- Extracts secondary pricing data (`Fiyat 2`) from management systems and injects it into designated spreadsheet columns.
- Preserves layout integrity by appending new columns safely without shifting embedded images or styling.

### 3. Automated Web Scraping & Image Embedding (`resim_bul.py`)
- Utilizes **Selenium WebDriver** (with optimized eager page load strategies) to query Google Images dynamically based on product names.
- **Smart Image Filtering:** Applies pixel-size filters (excluding elements under 60x60px) to automatically filter out UI logos, banners, and irrelevant icons.
- Processes raw image binaries via **Pillow (PIL)** and embeds clean, high-resolution product images directly into spreadsheet cells using **OpenPyXL**, adjusting row heights dynamically.

---

## Tech Stack
- **Languages:** Python
- **Libraries & Frameworks:** Pandas, OpenPyXL, Selenium, Requests, Pillow (PIL), WebDriver Manager
- **Automation Tools:** Chrome WebDriver

---

## Project Structure
```text
ecommerce-automation/
│
├── kdv_aktar.py       # Synchronizes VAT rates and primary prices via barcode matching
├── fiyat2.py          # Extracts and appends secondary pricing data
├── resim_bul.py       # Scrapes web images, filters quality, and embeds them into Excel
└── README.md          # Project documentation
