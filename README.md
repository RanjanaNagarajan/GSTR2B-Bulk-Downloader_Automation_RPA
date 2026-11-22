# GSTR-2B Excel Downloader

A Python automation utility that helps download **GSTR-2B Details (Excel)** from the GST Portal for multiple GSTINs and selected tax periods using Selenium.

> **Important:** This project automates browser navigation and file downloading. It does not bypass CAPTCHA, OTP, portal security controls, or access restrictions. The user must complete CAPTCHA and any additional verification manually.

## Project overview

The script is designed to reduce repetitive manual downloading of GSTR-2B Excel files for multiple GST registrations.

### What it does

- Reads GSTIN login details from an Excel workbook.
- Supports multiple GSTINs in one credentials workbook.
- Uses the Trade Name for output folder and file naming (if supplied).
- Opens the GST Portal login page using Chrome.
- Fills the username and password from the selected credentials workbook.
- Displays a small popup for the user to enter the CAPTCHA manually.
- Allows the user to complete OTP or other verification manually if the portal requests it.
- Navigates to the Returns Dashboard using the portal interface.
- Selects the financial year, quarter, and month.
- Opens the GSTR-2B summary using the **VIEW** action.
- Looks for the **DOWNLOAD GSTR-2B DETAILS (EXCEL)** action on the summary page.
- Waits for the Excel file to download and renames it.
- Creates an Excel download log with the GSTIN, Trade Name, financial year, month, status, file name, file path, remarks, and date/time.
- Saves screenshots and diagnostic HTML when certain page elements cannot be located.

## Month selection

The month numbers follow the Indian financial year (April to March):

| Input | Month |
|---:|---|
| 1 | April |
| 2 | May |
| 3 | June |
| 4 | July |
| 5 | August |
| 6 | September |
| 7 | October |
| 8 | November |
| 9 | December |
| 10 | January |
| 11 | February |
| 12 | March |
| 13 | All months (April to March) |

You can enter one month, a range, or a combination:

- `1` — April
- `1-3` — April, May, June
- `4-6` — July, August, September
- `1-6` — April to September
- `1-3,7-9` — April to June and October to December
- `1,3,5` — April, June and August
- `13` — all 12 months

The script prints the selected months before processing.

## Sample credentials Excel workbook

Use the included **[Sample_GSTIN_Credentials.xlsx](Sample_GSTIN_Credentials.xlsx)** as the structure for your input file.

The `GSTIN Credentials` sheet must contain these column headings in the first row:

| Column | Required? | Description |
|---|---|---|
| GSTIN | Yes | GST registration number |
| Trade Name | Recommended | Used for naming output files and folders |
| Username | Yes | GST Portal username |
| Password | Yes | GST Portal password |

`Trade Name` is optional in the script. If it is not supplied, the GSTIN is used as the name.

### Example layout

| GSTIN | Trade Name | Username | Password |
|---|---|---|---|
| YOUR_GSTIN_HERE | YOUR_TRADE_NAME_HERE | YOUR_GST_USERNAME_HERE | YOUR_GST_PASSWORD_HERE |

The values above are placeholders. Replace them with your own details in a **local copy** of the workbook. Add one GSTIN per row for multiple registrations.

## Requirements

- Windows is recommended (the script uses Tkinter popups and the default Downloads folder).
- Python 3.10 or later is recommended.
- Google Chrome installed.
- Internet access and a valid GST Portal account for each GSTIN.
- The GST Portal must be accessible and permit the user's session.

### Python packages

Install the required packages:

```bash
python -m pip install --upgrade selenium openpyxl
```

`tkinter` is included with most standard Windows Python installations. If your Python installation does not include Tkinter, install/repair Python with the Tcl/Tk support option enabled.

Selenium Manager may automatically manage the compatible ChromeDriver. If your environment blocks this, configure a compatible ChromeDriver separately.

## How to use

### 1. Download the project

Clone the repository:

```bash
git clone https://github.com/YOUR_GITHUB_USERNAME/YOUR_REPOSITORY_NAME.git
cd YOUR_REPOSITORY_NAME
```

Alternatively, download the repository as a ZIP file from GitHub and extract it.

### 2. Prepare the credentials workbook

1. Open `Sample_GSTIN_Credentials.xlsx`.
2. Save a separate local copy.
3. Replace the yellow placeholder row with your actual GSTIN, Trade Name, Username, and Password.
4. Add one row for each additional GSTIN.
5. Save the workbook as `.xlsx`.

**Never commit or upload your real credentials workbook to GitHub.**

### 3. Run the script

Open Command Prompt in the project folder and run:

```bash
python GSTR2B_Excel_Automation.py
```

The program asks for:

1. The full path of your local credentials Excel workbook.
2. The financial year, for example `2025-26`.
3. The month or month range, for example `1-3`.

Use quotes around the Excel path if it contains spaces, for example:

```text
C:\Users\YourName\Documents\GST Credentials.xlsx
```

### 4. Complete portal verification

For each GSTIN:

1. The script opens the GST Portal login page.
2. It fills the username and password.
3. Enter the CAPTCHA in the popup.
4. If OTP or any additional verification is requested, complete it manually in Chrome.
5. Allow the script to continue after successful login.

The script does not solve or bypass CAPTCHA or OTP.

### 5. Download and locate the files

The main output folder is:

```text
Downloads/
└── GSTR2B_DOWNLOADS/
```

Output behaviour:

- If one month is selected, the Excel file is saved directly in `GSTR2B_DOWNLOADS` using the Trade Name as the file name.
- If multiple months are selected, a folder is created using the Trade Name, and the monthly files are saved inside it.

Example for multiple months:

```text
GSTR2B_DOWNLOADS/
├── Example Traders/
│   ├── GSTR2B_Example Traders_2025-26_April.xlsx
│   ├── GSTR2B_Example Traders_2025-26_May.xlsx
│   └── GSTR2B_Example Traders_2025-26_June.xlsx
├── GSTR2B_Download_Log.xlsx
└── GSTR2B_Download_Log_Backup.xlsx
```

The exact downloaded extension depends on the file supplied by the GST Portal.

### 6. Check the download log

The script creates:

- `GSTR2B_Download_Log.xlsx` — main processing log.
- `GSTR2B_Download_Log_Backup.xlsx` — fallback log if the main log cannot be updated.

The log records the GSTIN, Trade Name, financial year, month, status, file name, file path, remarks, and date/time.

## Important notes and limitations

- The GST Portal may change its page layout, labels, or security flow. Selenium selectors may need maintenance if the portal changes.
- The automation depends on successful login and portal availability.
- CAPTCHA, OTP, and other verification must be completed by the user.
- The script does not bypass Access Denied responses or other portal security restrictions.
- A syntax check does not guarantee successful operation against every live GST Portal account or page layout. Test with one GSTIN and one month before processing multiple GSTINs.
- Keep Chrome open if the script reports an unexpected error and asks you to inspect the browser.
- Avoid running multiple instances of the script against the same download folder and log at the same time.
- Review downloaded files and the log after each run.

## Security and privacy

The credentials workbook contains usernames and passwords in readable Excel cells. Treat it as confidential:

- Keep the real credentials workbook on your own computer.
- Do not upload it to GitHub, email it publicly, or commit it to source control.
- Use the included sample workbook only as a template.
- Do not share screenshots, logs, or diagnostic HTML if they expose GSTINs, usernames, session information, or other confidential data.
- Use the automation only for GST registrations and accounts you are authorized to access.

## Repository contents

```text
.
├── GSTR2B_Excel_Automation.py
├── Sample_GSTIN_Credentials.xlsx
├── README.md
├── LICENSE
└── .gitignore
```

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file.

## Disclaimer

This is an independent automation project and is not affiliated with, endorsed by, or officially connected with the Goods and Services Tax Network (GSTN), the GST Portal, or the Government of India. GST Portal names and marks belong to their respective owners. Users are responsible for ensuring that their use complies with applicable laws, portal terms, and their authorization.
