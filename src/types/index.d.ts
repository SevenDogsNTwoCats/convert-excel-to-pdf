declare module 'excel-to-pdf-converter' {
  /**
   * Converts an Excel file to a PDF document
   * @param inputFilePath Path to the input Excel file
   * @param outputFileName Path for the output PDF file
   * @param enablePagination Whether to enable pagination (default: false)
   */
  export function convertExcelToPdf(
    inputFilePath: string,
    outputFileName: string,
    enablePagination?: boolean
  ): Promise<void>;
}