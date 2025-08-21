export function encodeCell(row, col) {
  let colLetters = "";
  while (col > 0) {
    const remainder = (col - 1) % 26;
    colLetters = String.fromCharCode(65 + remainder) + colLetters;
    col = Math.floor((col - 1) / 26);
  }
  return `${colLetters}${row}`;
}
