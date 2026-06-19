# Python Di Report Tool

A Python tool for generating ASM-wise PG Business reports with merchant counts.

## Features

- Processes RazorOrders and Payments data from Excel files
- Generates formatted Excel reports with ASM-wise breakdowns
- Supports PG-2 and PG-5 payment gateway data
- Includes Head-wise subtotals and Grand Total rows
- Styled Excel output with color-coded headers and borders

## Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/israrahmad341/Python_Di_Report_Tool.git
   ```

2. Install dependencies:
   ```bash
   pip install pandas openpyxl
   ```

3. Configure environment:
   - Copy `.env` and fill in your database connection details.

## Usage

This script is designed to be imported as a module. Run `master_tool.py` to generate reports.

To use `Razor_Report_Tool.py` as a module:

```python
from Razor_Report_Tool import generate_report

generate_report(
    razor_orders_file="path/to/RazorOrders.xlsx",
    payments_report_file="path/to/Payments_Report.xlsx",
    asm_file="path/to/asm.xlsx",
    output_dir="path/to/output/"
)
```

## Input Files

| File | Required Columns |
|------|-----------------|
| RazorOrders Excel | `RequestNo`, `code`, `Retailer`, `status`, `PGType`, `amount` |
| Payments_Report Excel | `Order Id`, `Status`, `Total Amount` |
| ASM Excel | `code`, `head`, `asmname`, `parentcode` |

## Output

- **`ASM_WISE_PG_BUSINESS_WITH_MERCHANT_COUNT.xlsx`** — Styled Excel report with head-wise and grand total rows.

## Requirements

- Python 3.7+
- `pandas`
- `openpyxl`
