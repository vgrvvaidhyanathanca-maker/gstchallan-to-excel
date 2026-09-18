# GST Challan Extractor

Reads GST challan PDFs downloaded from the GST portal and tabulates them in one
Excel sheet — one row per CPIN.

Handles both portal layouts:

| Layout | When you get it | Extra fields |
|---|---|---|
| **Form GST PMT-06** | Generated before payment | Expiry date, ticked mode of payment |
| **GST Payment Receipt** | After payment | CIN, bank, BRN/UTR, deposit date |

## Columns

File · Type · Status · CPIN · Challan Date · Expiry Date · GSTIN · Legal Name ·
State · Reason · Financial Year · Period ·
CGST / IGST / CESS / SGST × (Tax, Interest, Penalty, Fee, Others, Total) ·
Total Tax / Interest / Penalty / Fee / Others · Grand Total ·
Mode of Payment · Bank · BRN / UTR · CIN · Payment Date

UTGST rows are mapped to the SGST columns; the **State** column says which.

## Setup

```bash
py -m pip install -r requirements.txt
```

## Usage

Put the challan PDFs in a folder and run:

```bash
py gst_challan_extract.py "C:\path\to\challans"
```

`GST_Challans.xlsx` is written into that folder. Optional second argument is a
different output path.

Re-running is safe: the sheet is rebuilt from existing rows + new PDFs, keyed
on CPIN, sorted by challan date. If a CPIN exists both as a PMT-06 form and a
Payment Receipt, the receipt wins and Status becomes *Paid*. PDFs that are not
GST challans (income-tax challans, returns, notices) are ignored.

## How the tick-box is read

On the PMT-06 the three payment modes are all printed; the chosen one is marked
by a tick *image*, not text. The script renders the small square beside each
option and picks the one with ink in its centre (an empty box only has ink on
its border).

## Notes

- `*.pdf` and `*.xlsx` are git-ignored — client data stays local.
- Needs Python 3.10+, [PyMuPDF](https://pymupdf.readthedocs.io/) and openpyxl.
