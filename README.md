# Excel Product Data Cleaner

A Node.js tool that cleans messy product data from Excel sheets.  
It extracts attributes and values from raw strings like:

```
color=Blue,mb_hover_img=/b/l/blue_1.png,ram=8GB,rom=256GB,processor=Snapdragon 685<br />CPU: Octa-core<br />GPU: Adreno 610
```

and produces a **clean Excel file** with structured columns.

---

## ✨ Features

- Reads Excel `.xlsx` input files.
- Parses `key=value` pairs from messy strings.
- Strips **all HTML tags** and decodes entities (e.g., `&nbsp; → space`).
- Splits specifications into neat **Attribute → Value** fields.
- Supports three output formats:
  - **Wide format**: one product per row, attributes as columns.
  - **Long format**: stacked table with Attribute/Value rows.
  - **SKU-Attributes format**: per-SKU table `SKU | Attribute | Value` (and optional `ParentSKU`).
- Detects variation SKUs and links them to parent SKUs using suffix rules.
- Configurable input column and sheet.

---

## 🚀 Installation

```bash
git clone https://github.com/tahamahmood004/excel-product-cleaner.git
cd excel-product-cleaner
npm install
```

---

## 🛠 Usage

### Basic

```bash
node clean_excel.js input.xlsx output.xlsx
```

### Options

- `--col=N` → input column number (default `1`, i.e. column A).
- `--format=wide|long|sku-attrs` → output format.
- `--sheet=SheetName` → choose sheet by name.

#### SKU-Attributes mode specific:

- `--skuCol=HeaderName` → column header containing SKUs (default `sku`).
- `--blobCol=HeaderName` → column header containing the raw key=value blob (optional, auto-detects otherwise).
- `--includeParent=1` → also output a `ParentSKU` column.
- `--suffixes=A,B,C,...` → list of suffix tokens to strip when deriving parent SKUs.

---

## 🔎 Examples

```bash
# Process first column, wide format
node clean_excel.js input.xlsx output.xlsx --col=1 --format=wide

# Process column 2, long format
node clean_excel.js input.xlsx cleaned.xlsx --col=2 --format=long

# Extract SKU-wise attributes with ParentSKU detection
node clean_excel.js products.xlsx out.xlsx   --format=sku-attrs   --skuCol=sku   --blobCol=additional_attributes   --includeParent=1   --suffixes=BLACK,BLUE,GREEN,RED,WHITE,SILVER,GOLD,GRAY,GREY
```

---

## 📂 Output

**Wide format:**

| Row | color | ram | rom | processor | CPU | GPU |
|-----|-------|-----|-----|-----------|-----|-----|
| 2   | Blue  | 8GB | 256GB | Snapdragon 685 | Octa-core CPU | Adreno 610 |

**Long format:**

| Row | Attribute | Value           |
|-----|-----------|-----------------|
| 2   | color     | Blue            |
| 2   | ram       | 8GB             |
| 2   | rom       | 256GB           |

**SKU-Attributes format:**

| SKU          | Attribute    | Value             | ParentSKU |
|--------------|--------------|-------------------|-----------|
| 2306EPN60G   | color        | Sunrise Gold      |           |
| 2306EPN60GBL | color        | Black             | 2306EPN60G |
| 2306EPN60GBL | ram          | 8GB               | 2306EPN60G |

---

## ⚡ Notes

- Handles thousands of rows easily.
- For **very large files**, increase Node memory:
  ```bash
  node --max-old-space-size=4096 clean_excel.js big.xlsx out.xlsx --format=sku-attrs --skuCol=sku --includeParent=1
  ```
- Requires **Node.js 14+** (tested with Node.js 22).

---

## 📜 License

MIT
