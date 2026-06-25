# organize-reimbursement-skill

Codex skill for organizing Chinese reimbursement folders into an Excel reimbursement sheet.

## Contents

```text
organize-reimbursement/
  SKILL.md
  agents/
    openai.yaml
  assets/
    费用报销单模板.xlsx
  scripts/
    build_reimbursement.py
```

## Install

Copy the `organize-reimbursement` folder into your Codex skills directory:

```powershell
Copy-Item -Recurse .\organize-reimbursement "$env:USERPROFILE\.codex\skills\organize-reimbursement"
```

Then ask Codex to use `$organize-reimbursement` or say something like:

```text
请整理当前文件夹中的账单截图、订单截图、发票和费用报销单模板。
```

## Expected Workspace

```text
账单截图/
订单截图/
发票/
费用报销单模板.xlsx  # optional project-specific template
```

The skill treats bill screenshots as the primary source when they contain readable items, renames and embeds matching order screenshots by reason, appends order-only items in red, renames matched invoice PDFs by reason, and marks invoice-backed rows yellow.

Template lookup order:

1. Current project `费用报销单模板.xlsx`
2. Bundled `organize-reimbursement/assets/费用报销单模板.xlsx`
3. Auto-created standard reimbursement workbook

## Privacy

Do not commit real reimbursement screenshots, invoices, or generated reimbursement sheets. The bundled template should be a sanitized template that is safe to share.
