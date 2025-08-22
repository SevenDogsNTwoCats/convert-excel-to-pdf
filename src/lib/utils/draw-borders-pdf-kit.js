/**
 * Draws borders around a cell in a PDF document using jsPDF
 * @param {jsPDF} doc - The jsPDF document instance
 * @param {number} x - The x coordinate of the cell
 * @param {number} y - The y coordinate of the cell
 * @param {number} width - The width of the cell
 * @param {number} height - The height of the cell
 * @param {Object} borders - The border styles object containing top, left, bottom, right properties
 */
export function drawBorders(doc, x, y, width, height, borders) {
  if (!borders) return;

  const sides = ["top", "left", "bottom", "right"];
  
  sides.forEach((side) => {
    const border = borders[side];
    if (border && border.style && border.style.toLowerCase() !== "none") {
      const color = border.color?.argb 
        ? `#${border.color.argb.slice(2)}` 
        : "#000000";
      
      doc.setDrawColor(color);
      doc.setLineWidth(1);

      switch (side) {
        case "top":
          doc.line(x, y, x + width, y);
          break;
        case "left":
          doc.line(x, y, x, y + height);
          break;
        case "bottom":
          doc.line(x, y + height, x + width, y + height);
          break;
        case "right":
          doc.line(x + width, y, x + width, y + height);
          break;
      }
    }
  });
}