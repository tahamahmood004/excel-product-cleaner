# Excel Product Data Cleaner

A Node.js tool that cleans messy product data from Excel sheets.\
It extracts attributes and values from raw strings like:

    color=Blue,mb_hover_img=/b/l/blue_1.png,specifications=<p>Dimension: 241*156*7.5mm<br />Weight: 484g</p>

and produces a **clean Excel file** with structured columns.

------------------------------------------------------------------------

## ✨ Features

-   Reads Excel `.xlsx` input files.
-   Parses `key=value` pairs from messy strings.
-   Strips **all HTML tags** and decodes entities (e.g.,
    `&nbsp; → space`).
-   Splits specifications into neat **Attribute → Value** fields.
-   Supports two output formats:
    -   **Wide format**: one product per row, attributes as columns.
    -   **Long format**: stacked table with Attribute/Value rows.
-   Configurable input column and sheet.

------------------------------------------------------------------------

## 🚀 Installation

``` bash
git clone https://github.com/tahamahmood004/excel-product-cleaner.git
cd excel-product-cleaner
npm install
```

------------------------------------------------------------------------

## 🛠 Usage

### Basic

``` bash
node clean_excel.js input.xlsx output.xlsx
```

### Options

-   `--col=N` → input column number (default `1`, i.e. column A).
-   `--format=wide|long` → output format.
-   `--sheet=SheetName` → choose sheet by name.

### Examples

``` bash
# Process first column, wide format
node clean_excel.js input.xlsx output.xlsx --col=1 --format=wide

# Process column 2, long format
node clean_excel.js input.xlsx cleaned.xlsx --col=2 --format=long
```

------------------------------------------------------------------------

## 📂 Output

**Wide format (one product per row):**

  Row   color   mb_hover_img      Dimension       Weight   OS
  ----- ------- ----------------- --------------- -------- ---------------
  2     Blue    /b/l/blue_1.png   241*156*7.5mm   484g     Doke OS_P 3.0

**Long format (stacked attributes):**

  Row   Attribute      Value
  ----- -------------- -----------------
  2     color          Blue
  2     mb_hover_img   /b/l/blue_1.png
  2     Dimension      241*156*7.5mm
  2     Weight         484g

------------------------------------------------------------------------

## ⚡ Notes

-   Handles thousands of rows easily (tested with 50k+).
-   If files are very large, consider using CSV output for performance.
-   Requires Node.js 14+.

------------------------------------------------------------------------

## 📜 License

MIT
