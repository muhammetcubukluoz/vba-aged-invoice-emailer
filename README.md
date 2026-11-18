# Aged Invoice Report Automation (Excel VBA + Outlook)

A VBA macro that automates the daily **Aged Invoice Summary** report: it checks that the source data is fresh, refreshes an Excel template, archives a dated copy, and emails it to a configurable recipient list via Outlook. If anything goes wrong, it sends a failure notification instead of failing silently.

## Features

- **Data freshness check** – aborts if the source CSV was not refreshed today.
- **Automated refresh** – opens the template, refreshes all connections and pivots, and waits for queries to finish.
- **Dated archive copy** – saves each day's report as `AgedInvoice_yyyymmdd.xlsx`.
- **Standalone output** – removes Power Query queries and data connections from the archived copy, so recipients can open it without refresh errors.
- **Controlled distribution** – recipients and an on/off trigger are managed from an Excel sheet, so no code changes are needed to update the mailing list.
- **Error notifications** – sends an email with the failure reason at every critical step.

## How It Works

```
CSV fresh today?  ──No──►  Error email ─► stop
       │ Yes
       ▼
Refresh template & save dated copy  ──Fail──►  Error email ─► stop
       │ OK
       ▼
Read recipients from EmailList.xlsx  ──No trigger / no address──►  Error email ─► stop
       │ OK
       ▼
Send report via Outlook  ──Fail──►  Error email
```

## Project Structure

```
D:\Excel\data\sendDailyEmail\
├── SendAgedInvEmail.bat              # Launcher: opens the macro workbook, which starts the macros
├── Aged_Invoice_Template.xlsm        # Macro-enabled workbook that contains all the VBA code
├── Aged_Invoice_Template.xlsx        # Report template (with queries/pivots)
├── EmailList.xlsx                    # Recipient list and send trigger
├── Data\
│   └── Aged_Invoice_Details.csv      # Source data (must be refreshed daily)
└── Archive\
    └── AgedInvoice_yyyymmdd.xlsx     # Generated daily reports
```

## Requirements

- Windows with Microsoft Excel and Microsoft Outlook (desktop) installed
- VBA reference enabled: **Microsoft Outlook 16.0 Object Library**
  (VBA Editor → *Tools* → *References*)
- Macros enabled for `Aged_Invoice_Template.xlsm`, the workbook that contains all the VBA code

## Setup

1. Create the folder structure shown above, or change the `sFolderPath` constant in the code to match your own path.
2. Place `Aged_Invoice_Template.xlsx` in the root folder and make sure the `Data` and `Archive` subfolders exist.
3. Create `EmailList.xlsx` with two sheets:

   | Sheet | Content |
   |---|---|
   | `AgedInvoice_Em` | Header in `A1`, one recipient email per row starting from `A2` |
   | `SendFile` | Cell `B2` is the send trigger: any value other than `No` or blank enables sending (e.g. `Yes`) |

4. In `SendErrorEmail`, replace the placeholder with the address that should receive failure notifications.
5. Paste the VBA code into a module of `Aged_Invoice_Template.xlsm`, add the Outlook reference (see Requirements), and save the file as a macro-enabled workbook.
6. Create `SendAgedInvEmail.bat` in the root folder with the following content (adjust the path if needed):

   ```bat
   @Echo off
   Pushd %~dp0
   start Excel.exe "D:\Excel\data\sendDailyEmail\Aged_Invoice_Template.xlsm"
   Popd
   ```

## Usage

### Automatic run (recommended)

Double-click `SendAgedInvEmail.bat`. It opens `Aged_Invoice_Template.xlsm` in Excel, and the macros run automatically through the `Workbook_Open` event in `ThisWorkbook`, producing and sending the report without any further action.

### Manual run

Run the macro from the VBA editor or the *Macros* dialog:

```
CreateAgedInvoiceReport
```

### Daily scheduling

Point Windows Task Scheduler to `SendAgedInvEmail.bat` to send the report every morning. Note that the `.bat` file contains a hardcoded path to the workbook, so update it if the folder is moved.

## Error Handling

| Condition | Error email reason |
|---|---|
| CSV file not modified today | `csv file hasn't refreshed` |
| Template refresh / copy failed | `Couldn't create a copy` |
| Trigger is `No`/blank, or no email address | `Email Aborted due to No Trigger in Email List or No Email Id` |
| Outlook send failed | `Couldn't send email` |

## Code Overview

| Procedure | Purpose |
|---|---|
| `Workbook_Open` (in `ThisWorkbook`) | Event that starts the macro flow automatically when the workbook is opened |
| `CreateAgedInvoiceReport` | Main procedure that runs the whole flow |
| `CreateInvReportCopy` | Refreshes the template, saves the dated copy, strips queries/connections |
| `GetEmail` | Reads recipients and the trigger from `EmailList.xlsx` |
| `Send_Email_Single_Attach` | Sends the report as an Outlook attachment |
| `SendErrorEmail` | Sends the failure notification |

## Notes

- Email addresses in `EmailList.xlsx` and any personal addresses in the code should not be committed to a public repository.
- The macro calls `.Send` on the Outlook mail item, so emails are sent immediately without manual review.

## Tech Stack

Excel VBA · Outlook Object Library · Power Query · Pivot Tables
