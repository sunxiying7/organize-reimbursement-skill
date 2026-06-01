---
name: organize-reimbursement
description: Organize Chinese reimbursement folders into an Excel reimbursement sheet from bill screenshots, order screenshots, invoice PDFs, and an Excel template. Use when the user asks to整理报销, generate or update a reimbursement Excel, match bill/order/invoice items, rename invoice files by reimbursement reason, mark unmatched order-only items red, or mark invoice-backed items yellow.
---

# Organize Reimbursement

Use this skill to process a reimbursement workspace that usually contains:

- `账单截图/`: bill summary screenshots, the primary source when non-empty.
- `订单截图/`: order screenshots used as evidence images and as supplemental items.
- `发票/`: invoice PDFs to match, rename, and mark in the sheet.
- `费用报销单模板.xlsx`: Excel template.

## Workflow

1. Inspect the workspace files.
   - Use `Get-ChildItem -File -Recurse` to list screenshots, invoices, and templates.
   - Use `openpyxl` to inspect workbook sheet names, header rows, merged cells, row heights, and column positions.

2. Extract reimbursement items.
   - If `账单截图/` has readable bill items, make it the source of truth.
   - Sort by date ascending.
   - Use reason format `MM.DD事项`, for example `05.20水果和花`.
   - Keep the bill category when visible, usually `办公` or `餐饮`.
   - If an order screenshot exists but the bill screenshots do not contain that item, append it and mark it red.
   - If a bill item has no corresponding order screenshot, leave the screenshot cell empty.
   - If `账单截图/` is empty or has no readable content, derive items from `订单截图/`.

3. Match order screenshots.
   - Match primarily by date, amount, and item description.
   - Embed only matching order screenshots in the screenshot column.
   - Treat duplicate screenshots of the same order as one item unless the amounts/order details differ.
   - Note any uncertain mismatch in `备注` instead of inventing certainty.

4. Rename matched order screenshots.
   - Rename each confidently matched order screenshot to `<MM.DD事项>.<original extension>`.
   - Use the renamed screenshot path when embedding images in Excel.
   - If two screenshots map to the same reason, keep one only when they are duplicate evidence for the same order; otherwise append a suffix such as `_2`.
   - Do not rename unmatched or ambiguous screenshots.

5. Match invoices.
   - Render PDF pages to images when PDF text extraction is unreliable.
   - Match by amount, invoice item, date, seller, and reimbursement reason.
   - Rename confidently matched invoice PDFs to `<MM.DD事项>.pdf`.
   - Mark matched rows yellow and set `是否增值税发票` to `是`.
   - Leave unmatched invoices with their original filenames and mention them in the final response.

6. Generate the workbook.
   - Prefer the current template over old generated files.
   - Preserve the template header and signature area.
   - Insert enough rows for all items.
   - When extending beyond the template's original detail rows, unmerge detail/signature rows if needed to avoid merged-cell write errors.
   - Set the total formula to sum the actual amount column over the generated detail rows.

7. Validate.
   - Reopen the output workbook with `openpyxl`.
   - Verify row count, image count, amount sum, total formula, red supplemental rows, and yellow invoice rows.
   - Report unresolved invoices or ambiguous matches.

## Script

Use `scripts/build_reimbursement.py` after extracting item data. It expects a JSON config with:

```json
{
  "template": "费用报销单模板.xlsx",
  "output": "报销单_整理完成.xlsx",
  "sheet": "没发票",
  "date": "2026-06-01",
  "columns": {
    "category": "B",
    "reason": "C",
    "screenshot": "D",
    "amount": "E",
    "has_invoice": "F",
    "invoice_no": "G",
    "remark": "H",
    "total": "G"
  },
  "items": [
    {
      "category": "办公",
      "reason": "05.20水果和花",
      "amount": 72.10,
      "order_image": "",
      "has_invoice": true,
      "supplemental": false,
      "remark": ""
    }
  ],
  "rename_order_images": true,
  "invoice_renames": {
    "old.pdf": "05.20水果和花.pdf"
  }
}
```

Run:

```powershell
python path\to\organize-reimbursement\scripts\build_reimbursement.py config.json
```

The script does not OCR screenshots. Codex must first inspect screenshots/PDF previews and prepare the item JSON.

## Practical Notes

- Current project outputs may become invalid if copied from temporary generated files; always validate by reopening with `openpyxl`.
- Keep order screenshot filenames aligned to the final `reason` text whenever the match is confident.
- Do not mark an invoice row yellow unless the invoice is confidently matched.
- Use light yellow fill `FFF2CC` for invoice-backed rows.
- Use red font/fill for order-only supplemental rows.
- If a PDF cannot be parsed as text, render pages using Windows PDF APIs, browser tooling, or another available local renderer, then inspect the page image.
