/**
 * Decodes Excel-style cell references (e.g. 'A1') to row and column numbers
 * @param {string} ref - The cell reference (e.g. 'A1', 'B2', etc.)
 * @returns {Object} An object containing row and column numbers
 * @throws {Error} If the cell reference is invalid
 */
export function decodeCell(ref) {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(`Invalid cell reference: ${ref}`);
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
