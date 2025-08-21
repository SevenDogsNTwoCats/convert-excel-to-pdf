// Función para decodificar referencias tipo 'A1' a { row, col }
export function decodeCell(ref) {
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
