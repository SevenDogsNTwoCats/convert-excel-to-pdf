// Refine drawBorders function to ensure accurate border rendering
export function drawBorders(doc, x, y, width, height, borders) {
  if (!borders) return;

  const sides = ["top", "left", "bottom", "right"];
  sides.forEach((side) => {
    const border = borders[side];
    if (border && border.style && border.style.toLowerCase() !== "none") {
      const color = border.color?.argb
        ? `#${border.color.argb.slice(2)}`
        : "black";
      doc.strokeColor(color);

      switch (side) {
        case "top":
          doc
            .moveTo(x, y)
            .lineTo(x + width, y)
            .stroke();
          break;
        case "left":
          doc
            .moveTo(x, y)
            .lineTo(x, y + height)
            .stroke();
          break;
        case "bottom":
          doc
            .moveTo(x, y + height)
            .lineTo(x + width, y + height)
            .stroke();
          break;
        case "right":
          doc
            .moveTo(x + width, y)
            .lineTo(x + width, y + height)
            .stroke();
          break;
      }
    }
  });
}