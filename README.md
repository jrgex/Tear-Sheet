# BEATS Tear Sheet Generator

Automatically generates a styled Excel tear sheet from the BEATS monthly returns data.

## Workflow

1. Add new monthly return rows to `data/Beats data .xlsx`  
   *(column B = date, column C = decimal return, e.g. `0.0154` = 1.54%)*
2. Run the generator:
   ```bash
   python update_tear_sheet.py
   ```
3. Open the output file from `output/Tear_Sheet_MMMYYYY.xlsx`

Each run **replaces** all data in the output file with whatever is currently in the data spreadsheet.

## Setup

```bash
pip install -r requirements.txt
```

## Output

The Excel file contains:

| Section | Contents |
|---------|----------|
| Header | Fund name + reporting month |
| Key Highlights | Strategy bullet points |
| Performance Statistics | 3-Month ROR, YTD, Total Return, 6-Month ROR |
| General Information | Min investment, fees, liquidity |
| Statistics | Total return, Sharpe ratio, winning months %, alpha |
| Performance (VAMI) chart | Cumulative value chart starting at 1,000 |
| Monthly Returns chart | Bar chart of monthly % returns |
| Monthly Performance table | Year × month grid with annual totals |

## Data file format

`data/Beats data .xlsx` — Sheet1 with columns:

| (empty) | Month | Beats |
|---------|-------|-------|
| | 2023-01-01 | 0.0101 |
| | 2023-02-01 | 0.0108 |
| | … | … |

## Fund details

| Field | Value |
|-------|-------|
| Manager | Green Exchange |
| AFSL | 559619 |
| Min Investment | 50,000 AUD |
| Liquidity | Quarterly |
| Management Fee | 2.00% |
| Performance Fee | 15.00% |
| Highwater Mark | Yes |
