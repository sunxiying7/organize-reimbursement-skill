# organize-reimbursement-skill

Codex skill for organizing Chinese reimbursement folders into an Excel reimbursement sheet.

## Contents

```text
organize-reimbursement/
  SKILL.md
  agents/
    openai.yaml
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
费用报销单模板.xlsx
```

The skill treats bill screenshots as the primary source when they contain readable items, renames and embeds matching order screenshots by reason, appends order-only items in red, renames matched invoice PDFs by reason, and marks invoice-backed rows yellow.

## Privacy

Do not commit real reimbursement screenshots, invoices, Excel templates, or generated reimbursement sheets. The included `.gitignore` excludes common private artifacts.
