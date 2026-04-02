# Privacy Policy — Canvas SpeedGrader Helper

**Last updated:** 2 April 2026

## What this extension does

Canvas SpeedGrader Helper is a browser extension that helps instructors paste feedback comments from a CSV file into Canvas LMS SpeedGrader. It reads student names and anonymous IDs from the SpeedGrader page to match them against rows in your CSV.

## Data collection

This extension **does not collect, transmit, or share any data** with external servers. All processing happens locally in your browser.

Specifically:

- **CSV data** you upload is stored in your browser's local extension storage (`chrome.storage.local`) so it persists while you grade. It is never sent anywhere. You can clear it at any time using the "Clear" button.
- **Student names and IDs** are read from the Canvas SpeedGrader page solely to match against your CSV. They are held in memory only during use and are not stored or transmitted.
- **No analytics, tracking, or telemetry** of any kind is included.
- **No data is sent to any third party.**
- **No cookies** are set by this extension.

## Permissions used

| Permission | Why it is needed |
|---|---|
| `activeTab` | To identify the SpeedGrader tab when you click the extension icon |
| `scripting` | To inject scripts into SpeedGrader that detect the current student and fill the comment editor |
| `storage` | To keep your uploaded CSV data available while you grade |
| `sidePanel` | To display the extension UI as a side panel alongside SpeedGrader |
| Host access to `canvas.imperial.ac.uk` | Canvas LMS at Imperial College London is hosted on this domain; the extension needs page access to read student info and fill comments |

## Data retention

CSV data remains in local browser storage until you click "Clear" or uninstall the extension. No data is retained elsewhere.

## Contact

If you have questions about this privacy policy, please open an issue on the [GitHub repository](https://github.com/peterjatkinson/canvas-speedgrader-helper).
