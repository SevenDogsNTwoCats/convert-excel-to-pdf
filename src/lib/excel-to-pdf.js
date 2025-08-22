// src/lib/excel-to-pdf.js

import ExcelJS from "exceljs";
import fs from "fs";
import PDFDocument from "pdfkit";
import { drawBorders } from "./utils/draw-borders.js";
import { encodeCell } from "./utils/encodeCell.js";
import { decodeCell } from "./utils/decodeCell.js";
import { join, dirname } from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const fonts = {
  OpenSans: {
    normal: join(__dirname, "./fonts/OpenSans-Regular.ttf"),
    bold: join(__dirname, "./fonts/OpenSans-Bold.ttf"),
    italic: join(__dirname, "./fonts/OpenSans-Italic.ttf"),
    bolditalic: join(__dirname, "./fonts/OpenSans-BoldItalic.ttf"),
  },
};

/**
 * Converts an Excel file to a PDF document.
 * @param {string} inputFilePath Path to the input Excel file.
 * @param {string} outputFileName Name of the output PDF file.
 * @param {boolean} enablePagination Whether to enable pagination (default: false)
 */
export async function convertExcelToPdf(
  inputFilePath,
  outputFileName,
  enablePagination = false
) {
  // Check if file exists
  if (!fs.existsSync(inputFilePath)) {
    throw new Error(`Error: The file "${inputFilePath}" was not found.`);
  }

  // Create Workbook instance
  const workbook = new ExcelJS.Workbook();

  try {
    // Read Excel file asynchronously
    await workbook.xlsx.readFile(inputFilePath);

    const worksheet = workbook.getWorksheet(1);

    // Process data as before
    const styledRows = [];
    const totalRows = worksheet.rowCount;
    const totalCols = worksheet.columnCount;
    // Build merge map: { 'row-col': { range, startRow, startCol, endRow, endCol } }
    const mergeMap = {};
    if (worksheet._merges) {
      Object.entries(worksheet._merges).forEach(([key, merge]) => {
        const { top, left, bottom, right } = merge.model;
        for (let r = top; r <= bottom; r++) {
          for (let c = left; c <= right; c++) {
            mergeMap[`${r}-${c}`] = {
              range: `${top}-${left}:${bottom}-${right}`,
              startRow: top,
              startCol: left,
              endRow: bottom,
              endCol: right,
            };
          }
        }
      });
    }
    // Only add main cell of merge and skip secondary cells
    for (let rowNumber = 1; rowNumber <= totalRows; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      const cells = [];
      for (let colNumber = 1; colNumber <= totalCols; colNumber++) {
        const key = `${rowNumber}-${colNumber}`;
        // If cell is part of a merge and not the main one, skip it
        if (
          mergeMap[key] &&
          (rowNumber !== mergeMap[key].startRow ||
            colNumber !== mergeMap[key].startCol)
        ) {
          // Ensure secondary merge cells are empty
          cells.push({
            text: "",
            style: {},
            mergeInfo: null,
            isSecondaryMergeCell: true,
          });
          continue;
        }
        // Only process the main cell of the merge
        let cell = row.getCell(colNumber);
        let text =
          cell.value != null
            ? typeof cell.value === "object"
              ? ""
              : String(cell.value)
            : "";
        let style = cell.style || {};
        let mergeInfo = null;
        if (mergeMap[key]) {
          mergeInfo = mergeMap[key];
        }
        cells.push({ text, style, mergeInfo });
      }
      styledRows.push(cells);
    }

    // Adjust references before calling decodeCell
    Object.entries(mergeMap).forEach(([key, mergeInfo]) => {
      const startCell = encodeCell(mergeInfo.startRow, mergeInfo.startCol);
      const endCell = encodeCell(mergeInfo.endRow, mergeInfo.endCol);
      mergeInfo.range = `${startCell}:${endCell}`;
    });

    // Calculate dynamic table and page dimensions
    const defaultFontSize = 11;
    const rowHeight = 25;
    const tempDoc = new PDFDocument({ margin: 0 });
    const padding = 10;
    const extraSpace = 10;
    const colWidths = Array(totalCols).fill(padding);

    tempDoc.registerFont("Helvetica", fonts.OpenSans.normal);
    tempDoc.registerFont("Helvetica-Bold", fonts.OpenSans.bold);
    tempDoc.registerFont("Helvetica-Oblique", fonts.OpenSans.italic);
    tempDoc.registerFont("Helvetica-BoldOblique", fonts.OpenSans.bolditalic);

    // Calculate column widths considering all rows, including the header
    styledRows.forEach((row) => {
      row.forEach((cell, idx) => {
        if (cell.isSecondaryMergeCell) {
          return; // Ignore secondary merge cells
        }

        const text = cell.text || "";
        const size = cell.style.font?.size || defaultFontSize;
        let font = "Helvetica";
        if (cell.style.font?.bold) font = "Helvetica-Bold";
        else if (cell.style.font?.italic) font = "Helvetica-Oblique";
        tempDoc.font(font).fontSize(size);
        const textWidth = tempDoc.widthOfString(text) + padding + extraSpace;

        if (cell.mergeInfo) {
          // If cell is part of a merge, calculate total width of merged columns
          const { startCol, endCol } = cell.mergeInfo;
          const mergedWidth = colWidths
            .slice(startCol - 1, endCol)
            .reduce((sum, w) => sum + w, 0);

          // If total width of merged columns is less than required for the text, adjust it
          if (mergedWidth < textWidth) {
            const extraWidth = textWidth - mergedWidth;
            const numCols = endCol - startCol + 1;
            const additionalWidthPerCol = extraWidth / numCols;

            for (let i = startCol - 1; i < endCol; i++) {
              colWidths[i] += additionalWidthPerCol;
            }
          }
        } else {
          // If not a merged cell, adjust the width of the individual column
          if (textWidth > colWidths[idx]) {
            colWidths[idx] = textWidth;
          }
        }
      });
    });
    const tableWidth = colWidths.reduce((sum, w) => sum + w, 0);
    const tableHeight = totalRows * rowHeight + 10;
    const margin = 50;
    const pageWidth = tableWidth + margin * 2;
    const pageHeight = tableHeight + margin * 2 + 40;

    // Extract images from worksheet
    const images = [];
    worksheet.getImages().forEach((img) => {
      const range = img.range;
      const ext = range.ext;
      const tl = range?.tl;
      if (!tl || tl.col == null || tl.row == null) return;
      const media = workbook.getImage(img.imageId);
      if (!media?.buffer) return;
      images.push({ buffer: media.buffer, ext, tl });
    });

    // Generate PDF with dynamic size
    const doc = new PDFDocument({
      size: enablePagination ? "letter" : [pageWidth, pageHeight],
      margin,
    });

    doc.registerFont("Helvetica", fonts.OpenSans.normal);
    doc.registerFont("Helvetica-Bold", fonts.OpenSans.bold);
    doc.registerFont("Helvetica-Oblique", fonts.OpenSans.italic);
    doc.registerFont("Helvetica-BoldOblique", fonts.OpenSans.bolditalic);

    doc.pipe(fs.createWriteStream(outputFileName));

    // Draw images first
    images.forEach(({ tl, ext, buffer }) => {
      const colIndex = tl.nativeCol || 0;
      const rowIndex = tl.nativeRow || 0;
      const imgX =
        margin + colWidths.slice(0, colIndex).reduce((sum, w) => sum + w, 0);
      const imgY = margin + rowIndex * rowHeight;
      const ptsPerPx = 0.75;
      const imgWidthPts = (ext?.width || 0) * ptsPerPx;
      const imgHeightPts = (ext?.height || 0) * ptsPerPx;
      doc.image(buffer, imgX, imgY, {
        width: imgWidthPts,
        height: imgHeightPts,
      });
    });

    let y = margin;
    const startX = margin;

    // Process all rows
    styledRows.forEach((row, rowIdx) => {
      let x = startX; // Ensure x is initialized before use
      // Check if content exceeds page height
      if (enablePagination && y + rowHeight > pageHeight - margin) {
        doc.addPage(); // Add new page
        y = margin; // Reset vertical position
      }

      row.forEach((cell, i) => {
        // Check if cell is part of a merge
        let isMerged = false;
        let isMainMergeCell = false;
        let mergeCols = 1;
        let mergeRows = 1;

        if (cell.mergeInfo) {
          // Calculate merge range
          const [start, end] = cell.mergeInfo.range.split(":");
          const startCell = decodeCell(start);
          const endCell = decodeCell(end);
          mergeCols = endCell.col - startCell.col + 1;
          mergeRows = endCell.row - startCell.row + 1;
          isMerged = true;

          // Only draw if we are in the main cell of the merge (top-left)
          if (rowIdx + 1 === startCell.row && i + 1 === startCell.col) {
            isMainMergeCell = true;
          }
        }

        // If it's a merged cell but not the main one, skip drawing
        if (isMerged && !isMainMergeCell) {
          x += colWidths[i] || 10 * 6 + padding;
          return;
        }

        // Draw cell (merged or normal)
        if (isMerged && isMainMergeCell) {
          // Merged cell - calculate combined dimensions
          const mergedWidth = colWidths
            .slice(i, i + mergeCols)
            .reduce((sum, w) => sum + w, 0);
          // const mergedWidth = colWidths[i] || 10 * 6 + padding;
          const mergedHeight = rowHeight * mergeRows;

          // Background
          if (
            cell.style.fill &&
            cell.style.fill.fgColor &&
            cell.style.fill.fgColor.argb
          ) {
            const hex = cell.style.fill.fgColor.argb.slice(2);
            doc.rect(x, y, mergedWidth, mergedHeight).fill(`#${hex}`);
          }

          // Font and text
          if (cell.style.font) {
            const { size, color, bold, italic } = cell.style.font;
            if (size) doc.fontSize(size);
            if (bold) doc.font("Helvetica-Bold");
            if (italic) doc.font("Helvetica-Oblique");
            if (color && color.argb) doc.fillColor(`#${color.argb.slice(2)}`);
            else doc.fillColor("black");
          } else {
            doc.font("Helvetica").fontSize(11).fillColor("black");
          }
          const fontSize = cell.style.font?.size || 11;
          const dynamicYOffset = (mergedHeight - fontSize - 10) / 2;

          // Adjust vertical position of text using dynamicYOffset
          doc.text(cell.text, x + 2, y + dynamicYOffset, {
            width: mergedWidth - 4,
            align: cell.style.alignment?.horizontal || "left",
            // ellipsis: true,
          });

          // Borders
          const borders = cell.style.border || {};
          drawBorders(doc, x, y, mergedWidth, mergedHeight, borders);

          // Update x correctly so cells align
          x += colWidths[i] || 10 * 6 + padding; // Increment x by total width of merged cell
        } else {
          // Normal cell (not merged)
          const cellWidth = colWidths[i] || 10 * 6 + padding;

          // Background
          if (
            cell.style.fill &&
            cell.style.fill.fgColor &&
            cell.style.fill.fgColor.argb
          ) {
            const hex = cell.style.fill.fgColor.argb.slice(2);
            doc.rect(x, y, cellWidth, rowHeight).fill(`#${hex}`);
          }

          // Font and text
          if (cell.style.font) {
            const { size, color, bold, italic } = cell.style.font;
            if (size) doc.fontSize(size);
            if (bold) doc.font("Helvetica-Bold");
            if (italic) doc.font("Helvetica-Oblique");
            if (color && color.argb) doc.fillColor(`#${color.argb.slice(2)}`);
            else doc.fillColor("black");
          } else {
            doc.font("Helvetica").fontSize(11).fillColor("black");
          }
          const fontSize = cell.style.font?.size || 11;
          const dynamicYOffset = (fontSize * 1) / 2;

          // Adjust vertical position of text using dynamicYOffset
          doc.text(cell.text, x + 2, y + dynamicYOffset, {
            width: cellWidth - 4,
            align: cell.style.alignment?.horizontal || "left",
            ellipsis: true,
          });

          // Borders
          const borders = cell.style.border || {};
          drawBorders(doc, x, y, cellWidth, rowHeight, borders);

          // Update x correctly so cells align
          x += cellWidth; // Increment x by total width of normal cell
        }
      });
      y += rowHeight;
    });

    doc.end();
  } catch (error) {
    throw new Error(`Error processing Excel file: ${error.message}`);
  }
}
