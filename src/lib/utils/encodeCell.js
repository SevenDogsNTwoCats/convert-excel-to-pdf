/**
 * Converts row and column numbers to Excel-style cell references
 * @param {number} row - The row number (1-based)
 * @param {number} col - The column number (1-based)
 * @returns {string} Excel cell reference (e.g. 'A1', 'B2', 'AA10', etc.)
 * @example
 * encodeCell(1, 1) // returns 'A1'
 * encodeCell(2, 27) // returns 'AA2'
 */
export function encodeCell(row, col) {
  let colLetters = "";
  while (col > 0) {
    const remainder = (col - 1) % 26;
    colLetters = String.fromCharCode(65 + remainder) + colLetters;
    col = Math.floor((col - 1) / 26);
  }
  return `${colLetters}${row}`;
}
