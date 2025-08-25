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

## New 

Support for:
- **Formula evaluation**: Automatically resolves Excel formulas (COUNTIF, SUM, etc.) and displays calculated results
- **Page size limits**: Configure minimum and maximum page dimensions for better PDF output control
- **Decimal formatting**: Preserves original decimal formatting (700.00 stays as 700.00, not 700)
- **Rich text support**: Handles Excel rich text formatting and hyperlinks
- **Enhanced pagination**: Automatic pagination when content exceeds maximum dimensions

## Advanced Usage

```javascript
import { convertExcelToPdf } from 'excel-to-pdf-converter';

// Basic usage
await convertExcelToPdf({
  inputFilePath: 'input.xlsx',
  outputFilePath: 'output.pdf'
});

// Advanced configuration
await convertExcelToPdf({
  inputFilePath: 'input.xlsx',
  outputFilePath: 'output.pdf',
  enablePagination: true,
  MAX_WIDTH_SIZE: 14400,    // Maximum page width in points
  MAX_HEIGHT_SIZE: 14400,   // Maximum page height in points
  MIN_WIDTH_SIZE: 612,      // Minimum page width (letter size)
  MIN_HEIGHT_SIZE: 792,     // Minimum page height (letter size)
  useMinLimit: true,        // Enforce minimum size limits
  fixedAt: 2               // Decimal places for numeric formatting
});
```

## Configuration Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `inputFilePath` | string | required | Path to the input Excel file |
| `outputFilePath` | string | required | Path for the output PDF file |
| `enablePagination` | boolean | `false` | Enable automatic pagination |
| `MAX_WIDTH_SIZE` | number | `14400` | Maximum page width in points |
| `MAX_HEIGHT_SIZE` | number | `14400` | Maximum page height in points |
| `MIN_WIDTH_SIZE` | number | `612` | Minimum page width in points |
| `MIN_HEIGHT_SIZE` | number | `792` | Minimum page height in points |
| `useMinLimit` | boolean | `false` | Enforce minimum size limits |
| `fixedAt` | number | `2` | Decimal places for numbers |

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
