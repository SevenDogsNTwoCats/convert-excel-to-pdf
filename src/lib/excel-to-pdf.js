// src/lib/excel-to-pdf.js

import ExcelJS from "exceljs";
import fs from "fs";
import { jsPDF } from "jspdf";
import { drawBorders } from "./utils/draw-borders.js";
import { encodeCell } from "./utils/encodeCell.js";
import { decodeCell } from "./utils/decodeCell.js";
import { extractCellText } from "./utils/extractCellText.js";

/**
 * Converts an Excel file to a PDF document.
 * @param {string} inputFilePath Path to the input Excel file.
 * @param {string} outputFileName Name of the output PDF file.
 * @param {boolean} enablePagination Whether to enable pagination (default: false)
 */
export async function convertExcelToPdf(
  inputFilePath,
  outputFileName,
  enablePagination = false,
  fixedAt = 2
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

    // Try to calculate formulas (this might help with some formula results)
    try {
      await workbook.calcNow();
    } catch (calcError) {
      console.warn('Could not calculate formulas:', calcError.message);
    }

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
        let text = extractCellText(cell, fixedAt);
        
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
    const defaultFontSize = 12;
    let rowHeight = 20; // Make it let so we can scale it later
    // Create temporary jsPDF instance for text width calculation
    const tempDoc = new jsPDF({
      unit: 'pt',
      format: 'a4'
    });
    const padding = 10;
    const extraSpace = 10;
    const colWidths = Array(totalCols).fill(padding);

    // Calculate column widths considering all rows, including the header
    styledRows.forEach((row) => {
      row.forEach((cell, idx) => {
        if (cell.isSecondaryMergeCell) {
          return; // Ignore secondary merge cells
        }

        const text = cell.text || "";
        const size = cell.style.font?.size || defaultFontSize;
        let fontStyle = "normal";
        if (cell.style.font?.bold && cell.style.font?.italic) {
          fontStyle = "bolditalic";
        } else if (cell.style.font?.bold) {
          fontStyle = "bold";
        } else if (cell.style.font?.italic) {
          fontStyle = "italic";
        }
        
        tempDoc.setFont("helvetica", fontStyle);
        tempDoc.setFontSize(size);
        const textWidth = tempDoc.getTextWidth(text) + padding + extraSpace;

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
    let pageWidth = tableWidth + margin * 2;
    let pageHeight = tableHeight + margin * 2 + 40;
    
    // jsPDF limit is 14400 units - handle large documents intelligently
    const MAX_SIZE = 14400;
    let scaleFactor = 1;
    
    if (pageWidth > MAX_SIZE || pageHeight > MAX_SIZE) {
      if (pageHeight > MAX_SIZE && !enablePagination) {
        // If height exceeds limit, enable pagination and use ideal table width
        console.warn('Document height too large, enabling pagination automatically');
        enablePagination = true;
        
        // Keep the ideal table width (up to MAX_SIZE) for better readability
        if (pageWidth > MAX_SIZE) {
          pageWidth = MAX_SIZE;
          console.log(`Table width capped at ${MAX_SIZE} units for jsPDF compatibility`);
        }
        
        // Use standard letter height for pagination
        pageHeight = 792; // Letter height in points
        
      } else if (pageWidth > MAX_SIZE && pageHeight <= MAX_SIZE && !enablePagination) {
        // If only width exceeds limit, cap it at MAX_SIZE
        pageWidth = MAX_SIZE;
        console.log(`Table width capped at ${MAX_SIZE} units for jsPDF compatibility`);
        
      } else if (enablePagination) {
        // If pagination is already enabled, use standard letter size for height
        // but preserve table width up to MAX_SIZE
        if (pageWidth > MAX_SIZE) {
          pageWidth = MAX_SIZE;
          console.log(`Table width capped at ${MAX_SIZE} units for jsPDF compatibility`);
        }
        pageHeight = 792; // Letter height in points
      }
      
      console.log(`Final page dimensions: ${pageWidth}x${pageHeight}`);
    }

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
    const doc = new jsPDF({
      orientation: pageWidth > pageHeight ? 'landscape' : 'portrait',
      unit: 'pt',
      format: [pageWidth, pageHeight]
    });

    // Draw images first (jsPDF has limited image support)
    images.forEach(({ tl, ext, buffer }) => {
      try {
        const colIndex = tl.nativeCol || 0;
        const rowIndex = tl.nativeRow || 0;
        const imgX =
          margin + colWidths.slice(0, colIndex).reduce((sum, w) => sum + w, 0);
        const imgY = margin + rowIndex * rowHeight;
        const ptsPerPx = 0.75;
        const imgWidthPts = (ext?.width || 0) * ptsPerPx;
        const imgHeightPts = (ext?.height || 0) * ptsPerPx;
        
        // Convert buffer to base64 for jsPDF
        const base64String = buffer.toString('base64');
        const dataURL = `data:image/png;base64,${base64String}`;
        
        doc.addImage(dataURL, 'PNG', imgX, imgY, imgWidthPts, imgHeightPts);
      } catch (imageError) {
        console.warn('Could not add image:', imageError.message);
      }
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

      // For very wide tables, we'll just use the original logic but with smaller cells if needed
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
          const mergedHeight = rowHeight * mergeRows;

          // Background
          if (
            cell.style.fill &&
            cell.style.fill.fgColor &&
            cell.style.fill.fgColor.argb
          ) {
            const hex = cell.style.fill.fgColor.argb.slice(2);
            doc.setFillColor(`#${hex}`);
            doc.rect(x, y, mergedWidth, mergedHeight, 'F');
          }

          // Font and text configuration
          const fontSize = cell.style.font?.size || 12;
          const isBold = cell.style.font?.bold;
          const isItalic = cell.style.font?.italic;
          
          let fontStyle = 'normal';
          if (isBold && isItalic) fontStyle = 'bolditalic';
          else if (isBold) fontStyle = 'bold';
          else if (isItalic) fontStyle = 'italic';
          
          doc.setFont('helvetica', fontStyle);
          doc.setFontSize(fontSize);

          // Text color
          const textColor = cell.style.font?.color?.argb 
            ? `#${cell.style.font.color.argb.slice(2)}` 
            : '#000000';
          doc.setTextColor(textColor);

          // Text positioning and alignment
          const text = cell.text || '';
          const align = cell.style.alignment?.horizontal || 'left';
          const textY = y + mergedHeight / 2 + fontSize / 3; // Centered vertically

          if (align === 'center') {
            doc.text(text, x + mergedWidth / 2, textY, { align: 'center' });
          } else if (align === 'right') {
            doc.text(text, x + mergedWidth - 2, textY, { align: 'right' });
          } else {
            doc.text(text, x + 2, textY);
          }

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
            doc.setFillColor(`#${hex}`);
            doc.rect(x, y, cellWidth, rowHeight, 'F');
          }

          // Font and text configuration
          const fontSize = cell.style.font?.size || 12;
          const isBold = cell.style.font?.bold;
          const isItalic = cell.style.font?.italic;
          
          let fontStyle = 'normal';
          if (isBold && isItalic) fontStyle = 'bolditalic';
          else if (isBold) fontStyle = 'bold';
          else if (isItalic) fontStyle = 'italic';
          
          doc.setFont('helvetica', fontStyle);
          doc.setFontSize(fontSize);

          // Text color
          const textColor = cell.style.font?.color?.argb 
            ? `#${cell.style.font.color.argb.slice(2)}` 
            : '#000000';
          doc.setTextColor(textColor);

          // Text positioning and alignment
          const text = cell.text || '';
          const align = cell.style.alignment?.horizontal || 'left';
          const textY = y + rowHeight / 2 + fontSize / 3; // Centered vertically

          if (align === 'center') {
            doc.text(text, x + cellWidth / 2, textY, { align: 'center' });
          } else if (align === 'right') {
            doc.text(text, x + cellWidth - 2, textY, { align: 'right' });
          } else {
            doc.text(text, x + 2, textY);
          }

          // Borders
          const borders = cell.style.border || {};
          drawBorders(doc, x, y, cellWidth, rowHeight, borders);

          // Update x correctly so cells align
          x += cellWidth; // Increment x by total width of normal cell
        }
      });
      y += rowHeight;
    });

    // Save PDF
    doc.save(outputFileName);
  } catch (error) {
    throw new Error(`Error processing Excel file: ${error.message}`);
  }
}
