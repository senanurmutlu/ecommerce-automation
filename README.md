# E-Commerce Data Automation & Web Scraping

A Python-based automation project that simplifies e-commerce product data management by processing Excel spreadsheets, matching product barcodes, updating VAT and pricing information, and collecting product images from Google Images.

The project consists of three independent scripts that automate repetitive spreadsheet operations and help reduce manual data entry.

## Features

### 1. VAT and Price Processing — `kdv_aktar.py`

- Reads product data exported from BenimPOS.
- Matches products using normalized barcodes.
- Transfers VAT (KDV) rates and primary prices to the target Excel workbook.
- Creates a separate output file to preserve the original target workbook.
- Generates a report for unmatched product barcodes.

### 2. Product Image Collection — `resim_bul.py`

- Searches Google Images using product names.
- Uses Selenium to automate browser interactions.
- Downloads or processes image data using Requests and Pillow.
- Embeds product images directly into Excel worksheets.
- Adjusts row heights to accommodate embedded images.

### 3. Secondary Price Processing — `fiyat2.py`

- Reads secondary price information (`Fiyat 2`) from BenimPOS exports.
- Matches products using their barcodes.
- Adds secondary prices to the final spreadsheet.
- Renames the existing `Fiyat` column to `Fiyat 1`.
- Preserves the images by copying the image-containing workbook before processing.

## Tech Stack

- **Language:** Python
- **Data Processing:** Pandas
- **Excel Automation:** OpenPyXL
- **Browser Automation:** Selenium, WebDriver Manager
- **HTTP Requests:** Requests
- **Image Processing:** Pillow (PIL)

## Project Structure

```text
ecommerce-automation/
├── kdv_aktar.py
├── resim_bul.py
├── fiyat2.py
└── README.md
```

## Requirements

- Python 3.10 or a compatible version
- Google Chrome
- Visual Studio Code (recommended)
- The required Excel input files

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/senanurmutlu/ecommerce-automation.git
cd ecommerce-automation
```

Alternatively, download the repository as a ZIP file and open the extracted folder in VS Code.

### 2. Create a virtual environment

```bash
python -m venv .venv
```

### 3. Activate the virtual environment on Windows

For PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

For Command Prompt:

```cmd
.venv\Scripts\activate.bat
```

### 4. Install dependencies

```bash
python -m pip install pandas openpyxl selenium requests Pillow webdriver-manager
```

## Input Files and Configuration

The scripts expect specific Excel filenames, worksheet names, and column headers. Place the required input workbooks in the project directory, or update the filenames in the configuration section of each Python script.

| File | Purpose |
|---|---|
| `Ürünler_Dışa_Aktar_-__BenimPOS.xlsx` | BenimPOS product export containing barcode, VAT, and pricing data |
| `HARUN_CALISMA__1_.xlsx` | Target workbook for VAT and primary price processing |

The scripts also expect the worksheet names and column headers defined in their source code, including `Sheet1`, `Ürünlerinizi Burada Listeleyin`, `Ürün barkodu`, `KDV`, `Fiyat 1`, `Fiyat 2`, and `Barkod`.

The image collection script expects the output of the VAT and price processing step as its input. The secondary price script then uses the image-containing workbook.

**Important:** The example filenames above are the current values configured in the scripts. Replace them with your own filenames if necessary.

## Usage

Run the scripts individually from the VS Code terminal, in the following order.

### Step 1: Process VAT and primary prices

```bash
python kdv_aktar.py
```

**Output:**
- `HARUN_CALISMA_KDV_FIYAT_GUNCEL.xlsx`
- `eslesmeyen_barkodlar.xlsx` if unmatched barcodes are found

### Step 2: Collect and embed product images

```bash
python resim_bul.py
```

**Output:**
- `HARUN_CALISMA_RESIMLI_SON.xlsx`

This step requires Google Chrome and an internet connection.

### Step 3: Add secondary prices

```bash
python fiyat2.py
```

**Output:**
- `HARUN_CALISMA_EN_SON_FINAL.xlsx`

The final workbook contains the processed pricing information and embedded product images, provided the input data and matching barcodes are valid.

## Workflow

```text
BenimPOS Excel Export
        |
        v
VAT & Primary Price Processing
        |
        v
Updated Excel Workbook
        |
        v
Automated Product Image Collection
        |
        v
Workbook with Embedded Images
        |
        v
Secondary Price Processing
        |
        v
Final Excel Workbook
```

## Notes

- Keep the required input files in the expected directory.
- Verify worksheet names and column headers before running the scripts.
- Run the scripts in the documented order because later steps depend on earlier output files.
- Review the output workbooks and unmatched barcode report after processing.
- Google Images results may not always correspond to the exact product. Verify images before using the generated workbook commercially.
- This project was developed to explore Python scripting, Excel automation, data processing, web scraping, and browser automation.

## Author

**Sena Nur Mutlu**

Software Engineering Student  
Karadeniz Technical University

GitHub: [@senanurmutlu](https://github.com/senanurmutlu)
