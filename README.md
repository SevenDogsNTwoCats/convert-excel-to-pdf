# excel-to-pdf-converter

Convert Excel files (.xlsx) to PDF easily with style preservation.

## Installation

```bash
npm install excel-to-pdf-converter
```

## Usage

```javascript
import { convertExcelToPdf } from 'excel-to-pdf-converter';

await convertExcelToPdf('path/to/file.xlsx', 'path/to/output.pdf');
```

## Features

- Preserves cell styles (fonts, colors, borders)
- Supports merged cells
- Maintains text alignment
- Handles background colors
- Supports custom cell borders

## API

### convertExcelToPdf(inputFilePath, outputFilePath)
Converts an Excel file (.xlsx) to PDF.
- `inputFilePath`: Path to the input Excel file
- `outputFilePath`: Path and name for the output PDF file

## Examples

```javascript
// Basic usage
await convertExcelToPdf('input.xlsx', 'output.pdf');

// With custom path
await convertExcelToPdf('./documents/spreadsheet.xlsx', './exports/report.pdf');
```

## Requirements
- Node.js 12 or higher
- NPM or Yarn

## License
ISC
