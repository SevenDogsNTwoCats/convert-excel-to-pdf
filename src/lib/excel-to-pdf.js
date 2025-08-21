// src/lib/excel-to-pdf.js

import ExcelJS from "exceljs";
import fs from "fs";
import PDFDocument from "pdfkit";
import { drawBorders } from "./utils/draw-borders.js";

/**
 * Convierte un archivo de Excel a un documento PDF.
 * @param {string} inputFilePath La ruta al archivo Excel de entrada.
 * @param {string} outputFileName El nombre del archivo PDF de salida.
 */
export async function convertExcelToPdf(inputFilePath, outputFileName) {
  // Verifica si el archivo existe
  if (!fs.existsSync(inputFilePath)) {
    throw new Error(`Error: El archivo "${inputFilePath}" no se encontró.`);
  }

  // Crea una instancia de Workbook
  const workbook = new ExcelJS.Workbook();

  try {
    // Lee el archivo de Excel de forma asíncrona
    await workbook.xlsx.readFile(inputFilePath);

    // Obtiene la primera hoja de trabajo por su índice (es 1-based, no 0-based)
    const worksheet = workbook.getWorksheet(1);

    // Función para decodificar referencias tipo 'A1' a { row, col }
    function decodeCell(ref) {
      const match = ref.match(/^([A-Z]+)(\d+)$/);
      if (!match) {
        throw new Error(`Referencia de celda inválida: ${ref}`);
      }
      const colLetters = match[1];
      const row = parseInt(match[2], 10);
      // Convierte letras de columna a número
      let col = 0;
      for (let i = 0; i < colLetters.length; i++) {
        col *= 26;
        col += colLetters.charCodeAt(i) - 64;
      }
      return { row, col };
    }

    function encodeCell(row, col) {
      let colLetters = "";
      while (col > 0) {
        const remainder = (col - 1) % 26;
        colLetters = String.fromCharCode(65 + remainder) + colLetters;
        col = Math.floor((col - 1) / 26);
      }
      return `${colLetters}${row}`;
    }

    // Procesa los merges para marcar solo la celda principal y dejar las demás vacías
    const styledRows = [];
    const totalRows = worksheet.rowCount;
    const totalCols = worksheet.columnCount;
    // Construye un mapa de merges: { 'row-col': { range, startRow, startCol, endRow, endCol } }
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
    // Solo agregar la celda principal del merge y saltar las secundarias
    for (let rowNumber = 1; rowNumber <= totalRows; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      const cells = [];
      for (let colNumber = 1; colNumber <= totalCols; colNumber++) {
        const key = `${rowNumber}-${colNumber}`;
        // Si la celda es parte de un merge y no es la principal, saltar
        if (
          mergeMap[key] &&
          (rowNumber !== mergeMap[key].startRow ||
            colNumber !== mergeMap[key].startCol)
        ) {
          // Asegurarse de que las celdas secundarias de un merge estén vacías
          cells.push({
            text: "",
            style: {},
            mergeInfo: null,
            isSecondaryMergeCell: true,
          });
          continue;
        }
        // Solo procesar la celda principal del merge
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

    // Ajustar referencias antes de llamar a decodeCell
    Object.entries(mergeMap).forEach(([key, mergeInfo]) => {
      const startCell = encodeCell(mergeInfo.startRow, mergeInfo.startCol);
      const endCell = encodeCell(mergeInfo.endRow, mergeInfo.endCol);
      mergeInfo.range = `${startCell}:${endCell}`;
    });

    // Calcula dimensiones dinámicas de la tabla y la página
    const defaultFontSize = 12;
    const rowHeight = 20;
    const tempDoc = new PDFDocument({ margin: 0 });
    const padding = 10;
    const extraSpace = 10;
    const colWidths = Array(totalCols).fill(padding);

    // Calcula el ancho de columnas considerando todas las filas, incluyendo el header
    styledRows.forEach((row) => {
      row.forEach((cell, idx) => {
        if (cell.isSecondaryMergeCell) {
          return; // Ignorar celdas secundarias de un merge
        }

        const text = cell.text || "";
        const size = cell.style.font?.size || defaultFontSize;
        let font = "Helvetica";
        if (cell.style.font?.bold) font = "Helvetica-Bold";
        else if (cell.style.font?.italic) font = "Helvetica-Oblique";
        tempDoc.font(font).fontSize(size);
        const textWidth = tempDoc.widthOfString(text) + padding + extraSpace;

        if (cell.mergeInfo) {
          // Si la celda es parte de un merge, calcular el ancho total de las columnas mergeadas
          const { startCol, endCol } = cell.mergeInfo;
          const mergedWidth = colWidths
            .slice(startCol - 1, endCol)
            .reduce((sum, w) => sum + w, 0);

          // Si el ancho total de las columnas mergeadas es menor que el ancho necesario para el texto, ajustar
          if (mergedWidth < textWidth) {
            const extraWidth = textWidth - mergedWidth;
            const numCols = endCol - startCol + 1;
            const additionalWidthPerCol = extraWidth / numCols;

            for (let i = startCol - 1; i < endCol; i++) {
              colWidths[i] += additionalWidthPerCol;
            }
          }
        } else {
          // Si no es una celda mergeada, ajustar el ancho de la columna individual
          if (textWidth > colWidths[idx]) {
            colWidths[idx] = textWidth;
          }
        }
      });
    });
    const tableWidth = colWidths.reduce((sum, w) => sum + w, 0);
    const tableHeight = totalRows * rowHeight;
    const margin = 50;
    const pageWidth = tableWidth + margin * 2;
    const pageHeight = tableHeight + margin * 2;

    // Extrae imágenes del worksheet
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

    // Genera el PDF con tamaño dinámico
    const doc = new PDFDocument({ size: [pageWidth, pageHeight], margin });
    doc.pipe(fs.createWriteStream(outputFileName));

    // Dibuja imágenes primero
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

    styledRows.forEach((row, rowIdx) => {
      let x = startX; // Asegurar que x esté inicializado antes de su uso
      row.forEach((cell, i) => {
        // Verificar si la celda es parte de un merge
        let isMerged = false;
        let isMainMergeCell = false;
        let mergeCols = 1;
        let mergeRows = 1;

        if (cell.mergeInfo) {
          // Calcular el rango del merge
          const [start, end] = cell.mergeInfo.range.split(":");
          const startCell = decodeCell(start);
          const endCell = decodeCell(end);
          mergeCols = endCell.col - startCell.col + 1;
          mergeRows = endCell.row - startCell.row + 1;
          isMerged = true;

          // Solo dibujar si estamos en la celda principal del merge (top-left)
          if (rowIdx + 1 === startCell.row && i + 1 === startCell.col) {
            isMainMergeCell = true;
          }
        }

        // Si es una celda mergeada pero no es la principal, saltar sin dibujar
        if (isMerged && !isMainMergeCell) {
          x += colWidths[i] || 10 * 6 + padding;
          return;
        }

        // Dibujar la celda (mergeada o normal)
        if (isMerged && isMainMergeCell) {
          // Celda mergeada - calcular dimensiones combinadas
          const mergedWidth = colWidths
            .slice(i, i + mergeCols)
            .reduce((sum, w) => sum + w, 0);
          // const mergedWidth = colWidths[i] || 10 * 6 + padding;
          const mergedHeight = rowHeight * mergeRows;

          // Fondo
          if (
            cell.style.fill &&
            cell.style.fill.fgColor &&
            cell.style.fill.fgColor.argb
          ) {
            const hex = cell.style.fill.fgColor.argb.slice(2);
            doc.rect(x, y, mergedWidth, mergedHeight).fill(`#${hex}`);
          }

          // Fuente y texto
          if (cell.style.font) {
            const { size, color, bold, italic } = cell.style.font;
            if (size) doc.fontSize(size);
            if (bold) doc.font("Helvetica-Bold");
            if (italic) doc.font("Helvetica-Oblique");
            if (color && color.argb) doc.fillColor(`#${color.argb.slice(2)}`);
            else doc.fillColor("black");
          } else {
            doc.font("Helvetica").fontSize(12).fillColor("black");
          }
          const fontSize = cell.style.font?.size || 12;
          const dynamicYOffset = (mergedHeight - fontSize) / 2;

          // Ajusta la posición vertical del texto usando dynamicYOffset
          doc.text(cell.text, x + 2, y + dynamicYOffset, {
            width: mergedWidth - 4,
            align: cell.style.alignment?.horizontal || "left",
            // ellipsis: true,
          });

          // Bordes
          const borders = cell.style.border || {};
          drawBorders(doc, x, y, mergedWidth, mergedHeight, borders);

          // Actualiza x correctamente para que las celdas colinden
          x += colWidths[i] || 10 * 6 + padding; // Incrementa x por el ancho total de la celda mergeada
        } else {
          // Celda normal (no mergeada)
          const cellWidth = colWidths[i] || 10 * 6 + padding;

          // Fondo
          if (
            cell.style.fill &&
            cell.style.fill.fgColor &&
            cell.style.fill.fgColor.argb
          ) {
            const hex = cell.style.fill.fgColor.argb.slice(2);
            doc.rect(x, y, cellWidth, rowHeight).fill(`#${hex}`);
          }

          // Fuente y texto
          if (cell.style.font) {
            const { size, color, bold, italic } = cell.style.font;
            if (size) doc.fontSize(size);
            if (bold) doc.font("Helvetica-Bold");
            if (italic) doc.font("Helvetica-Oblique");
            if (color && color.argb) doc.fillColor(`#${color.argb.slice(2)}`);
            else doc.fillColor("black");
          } else {
            doc.font("Helvetica").fontSize(12).fillColor("black");
          }
          const fontSize = cell.style.font?.size || 12;
          const dynamicYOffset = (fontSize * 1) / 2;

          // Ajusta la posición vertical del texto usando dynamicYOffset
          doc.text(cell.text, x + 2, y + dynamicYOffset, {
            width: cellWidth - 4,
            align: cell.style.alignment?.horizontal || "left",
            ellipsis: true,
          });

          // Bordes
          const borders = cell.style.border || {};
          drawBorders(doc, x, y, cellWidth, rowHeight, borders);

          // Actualiza x correctamente para que las celdas colinden
          x += cellWidth; // Incrementa x por el ancho total de la celda normal
        }
      });
      y += rowHeight;
    });

    doc.end();
  } catch (error) {
    throw new Error(`Error al procesar el archivo Excel: ${error.message}`);
  }
}
