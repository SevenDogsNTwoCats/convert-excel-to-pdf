declare module "excel-to-pdf-converter" {
  /**
   * Converts an Excel file to a PDF document.
   * @param {string} inputFilePath Path to the input Excel file.
   * @param {string} outputFilePath Name of the output PDF file.
   * @param {boolean} enablePagination Whether to enable pagination (default: false)
   * @param {number} MAX_WIDTH_SIZE Maximum width size for the PDF (default: 14400, max jsPDF limit)
   * @param {number} MAX_HEIGHT_SIZE Maximum height size for the PDF (default: 14400, max jsPDF limit)
   * @param {number} MIN_WIDTH_SIZE Minimum width size for the PDF (default: 612, letter width in points)
   * @param {number} MIN_HEIGHT_SIZE Minimum height size for the PDF (default: 792, letter height in points)
   * @param {boolean} useMinLimit Whether to enforce minimum size limits (default: false)
   * @param {number} fixedAt Number of decimal places for numeric formatting (default: 2)
   */
  export function convertExcelToPdf({
    inputFilePath,
    outputFilePath,
    enablePagination,
    MAX_WIDTH_SIZE,
    MAX_HEIGHT_SIZE,
    MIN_WIDTH_SIZE,
    MIN_HEIGHT_SIZE,
    useMinLimit,
    fixedAt,
  }: {
    inputFilePath: string;
    outputFilePath: string;
    enablePagination?: boolean;
    MAX_WIDTH_SIZE?: number;
    MAX_HEIGHT_SIZE?: number;
    MIN_WIDTH_SIZE?: number;
    MIN_HEIGHT_SIZE?: number;
    useMinLimit?: boolean;
    fixedAt?: number;
  }): Promise<void>;
}
