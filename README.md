# Hachette Email to Excel Importer

A web-based tool for converting Hachette order email data into formatted Excel files.

## Overview

This application allows you to quickly convert order data received via email into properly formatted Excel spreadsheets. It handles both standard comma-separated data and run-together text formats that often occur when copying from email clients.

## Features

- **Smart Parsing**: Automatically detects and handles data with or without line breaks
- **Publisher Recognition**: Intelligently splits data by recognizing publisher names (HODDER, JOHN MURRAY, ORION, LITTLE BROWN, etc.)
- **Live Preview**: View parsed data in a table before downloading
- **Excel Export**: Generates properly formatted .xlsx files with appropriate column widths
- **No Server Required**: Runs entirely in your browser - all data stays on your computer
- **Responsive Design**: Works on desktop, tablet, and mobile devices

## Installation

1. Download the `hachette_importer.html` file
2. Save it anywhere on your computer
3. Double-click to open in your web browser

That's it! No installation, no dependencies, no configuration needed.

## Usage

### Basic Workflow

1. **Copy Order Data**: Copy the order data from your Hachette email
2. **Paste**: Paste the data into the text area
3. **Parse**: Click the "Parse Data" button (or press Ctrl/Cmd + Enter)
4. **Preview**: Review the parsed data in the table
5. **Download**: Click "Download Excel File" to save the .xlsx file

### Expected Data Format

The application expects data in the following column order:

1. Publisher
2. PO
3. ISBN
4. Title
5. Full Parcel Qty
6. Full Parcel Size
7. Part Parcel
8. Total Qty
9. Special Instructions

### Example Input

```
Publisher,PO,ISBN,Title,Full Parcel Qty,Full Parcel Size,Part Parcel,Total Qty,Special Instructions
HODDER [SHORT RUN],HA00458786,9780340897508,ASR DEATH OF A RED HEROINE,2,40,0,80,Actual
JOHN MURRAY,HA00459912,9780340995105,ASR GENEROUS JUSTICE,2,80,0,160,Actual
```

**Note**: The application can handle data that's all run together without line breaks, which commonly happens when copying from emails.

## Keyboard Shortcuts

- **Ctrl/Cmd + Enter**: Parse data (when focused in the text area)

## Output

The application generates Excel files with:

- Timestamp in filename (e.g., `Hachette_Orders_2026-01-06.xlsx`)
- Properly sized columns for optimal readability
- Single worksheet named "Hachette Orders"
- All data preserved exactly as parsed

## Technical Details

### Browser Compatibility

Works in all modern browsers:
- Chrome/Edge (recommended)
- Firefox
- Safari
- Opera

### Technologies Used

- HTML5
- CSS3 (Bootstrap 5.3.2)
- JavaScript (ES6+)
- SheetJS (xlsx library) for Excel generation

### Data Privacy

All data processing happens locally in your browser. No data is sent to any server or third party. The application works completely offline after the initial page load.

## Troubleshooting

### Data Not Parsing Correctly

**Problem**: Data appears in wrong columns or not split correctly

**Solutions**:
- Ensure the data includes the header row (Publisher, PO, ISBN, etc.)
- Check that publisher names match expected patterns (HODDER, JOHN MURRAY, etc.)
- Try adding line breaks between rows manually if the smart parser fails

### Excel File Not Downloading

**Problem**: Nothing happens when clicking "Download Excel File"

**Solutions**:
- Check your browser's download settings/permissions
- Try a different browser
- Ensure popup blockers aren't interfering

### Preview Shows "No Data Loaded"

**Problem**: After parsing, preview remains empty

**Solutions**:
- Verify data is pasted in the text area
- Check that data has at least a header row and one data row
- Look for error messages in the status alert

## Limitations

- Preview limited to first 50 records (all records included in Excel download)
- Maximum practical limit of ~10,000 records (browser memory dependent)
- Requires JavaScript enabled in browser

## Version History

### Version 1.0
- Initial release
- Smart parsing for run-together email data
- Publisher pattern recognition
- Bootstrap 5 styling
- Excel export with proper formatting

## Support

For issues or questions, refer to the inline instructions in the application or review the example data format provided.

## License

This tool is provided as-is for internal use.

---

**Quick Start**: Simply open `hachette_importer.html` in your browser and start pasting order data!