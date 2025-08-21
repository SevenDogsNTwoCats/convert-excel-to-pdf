declare module 'excel-to-pdf-converter' {
  interface ExcelToPDFOptions {
    margins?: {
      top?: number;
      bottom?: number;
      left?: number;
      right?: number;
    };
    pageSize?: string;
    landscape?: boolean;
  }

  interface ExcelToPDF {
    convert(excelPath: string, pdfPath: string, options?: ExcelToPDFOptions): Promise<void>;
  }

  const excelToPDF: ExcelToPDF;
  export default excelToPDF;
}
