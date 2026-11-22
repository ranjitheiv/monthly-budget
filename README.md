# Monthly Budget

A dashboard for the Monthly Budget Google Sheet. It shows each month and year against budget, lets you investigate overruns and flag why they happened, and finds patterns in spending.

## How it fits together

- **The page** (`index.html`) lives here and is hosted on GitHub Pages. It holds no financial data.
- **The API** (`apps-script/Code.gs`) lives inside the Google Sheet. The page asks it for the data and sends it any flags you save.
- **The secret key** is what stops anyone else reading the data. It is stored once on each device. Treat it like a password.

The script never edits Transactions or Map. It changes Budgets only when you press Reallocate or Publish in the Plan tab, matching rows exactly on month, line, section and person, and it leaves formula cells alone. Everything else it writes goes to tabs it creates:

| Tab | What it holds |
|---|---|
| Merchant | Every bank description and the merchant name it rolls up to. Fix a wrong name in column B and set column C to `manual`. |
| Flags | Every merchant you flag in investigate mode, with the reason and note. |
| Markers | Optional timeline labels: a date in column A, a short label in column B. |
| Plan Draft | Saved drafts of next year's budget, one row per person, month and line. |
| Change Log | Every reallocation and publish the app makes, with before and after values. |
| Budgets backup … | A copy of Budgets taken before each publish. The last five are kept. |

## Setting up the Apps Script

1. Open the **Monthly Budget** sheet, then **Extensions > Apps Script**.
2. Delete what is in `Code.gs`, paste in the contents of `apps-script/Code.gs`, and save.
3. Choose `setup` in the function list at the top and press **Run**. Approve the permissions when asked.
4. Open **Execution log**. Copy the secret key, and check that the row counts look right.
5. **Deploy > New deployment**. Type: **Web app**. Execute as: **Me**. Who has access: **Anyone**. Press **Deploy** and copy the web app URL (it ends in `/exec`).

After changing `Code.gs` later: **Deploy > Manage deployments > Edit > Version: New version > Deploy**. The URL stays the same.

## Connecting a device

Open the GitHub Pages link, paste the web app URL and the secret key, then **Connect**. To add a second device, open **Settings > Copy setup link for another device** and send the link to yourself privately.

On iPhone: open the link in Safari, then **Share > Add to Home Screen**.
