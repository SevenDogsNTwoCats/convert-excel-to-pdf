/**
 * Extracts text from an Excel cell, handling different value types and formatting
 * @param {Object} cell - The Excel cell object
 * @param {number} fixedAt - Number of decimal places for numeric formatting (default: 2)
 * @returns {string} The formatted text value
 */
export function extractCellText(cell, fixedAt = 2) {
  if (cell.value == null) {
    return "";
  }

  // Handle object values (richText, formulas, etc.)
  if (typeof cell.value === "object") {
    // Handle formulas with calculated results
    if (cell.value.result !== undefined) {
      // If the formula has a calculated result, use it
      const result = cell.value.result;
      
      // Format the result based on its type
      if (typeof result === "number") {
        const format = cell.numFmt || "";
        
        if (format.includes('0.00') || format.includes('#.##') || format.includes('0.0')) {
          const decimalMatch = format.match(/\.0+/);
          const decimalPlaces = decimalMatch ? decimalMatch[0].length - 1 : fixedAt;
          return result.toFixed(decimalPlaces);
        }
        
        // For numbers with natural decimals, preserve them
        if (result % 1 !== 0) {
          return result.toString();
        }
        
        return result.toString();
      }
      
      return String(result);
    }
    
    // Handle richText
    if (cell.value.richText && Array.isArray(cell.value.richText)) {
      return cell.value.richText
        .map(part => part.text || "")
        .join("");
    }
    
    // Handle hyperlinks
    if (cell.value.hyperlink && cell.value.text) {
      return cell.value.text;
    }
    
    // Handle other object types with text property
    if (cell.value.text) {
      return cell.value.text;
    }
    
    // For formulas without results, try different approaches
    if (cell.value.formula) {
      // Check if there's a cached result or error value
      if (cell.value.error) {
        return `#${cell.value.error.toUpperCase()}`;
      }
      
      // Use cell.text if available for formulas (often contains calculated value)
      if (cell.text && cell.text !== "") {
        return cell.text;
      }
      
      // If no result available, show the formula
      return `=${cell.value.formula}`;
    }
    
    // Handle shared formulas without results
    if (cell.value.sharedFormula) {
      // Use cell.text if available
      if (cell.text && cell.text !== "") {
        return cell.text;
      }
      // Try to show a placeholder or the formula reference
      return "0"; // Default value for unresolved shared formulas
    }
    
    // If it's an empty object or unrecognized structure, return empty string
    if (Object.keys(cell.value).length === 0) {
      return "";
    }
    
    // Fallback for other objects (but avoid showing complex formula objects)
    return "";
  }

  // Handle numeric values
  if (typeof cell.value === "number") {
    try {
      const format = cell.numFmt || "";
      
      // Check if the format explicitly defines decimal places
      if (format.includes('0.00') || format.includes('#.##') || format.includes('0.0')) {
        const decimalMatch = format.match(/\.0+/);
        const decimalPlaces = decimalMatch ? decimalMatch[0].length - 1 : fixedAt;
        return cell.value.toFixed(decimalPlaces);
      }
      
      // Check if cell.text shows decimals but preserve our decimal formatting logic
      if (cell.text && cell.text.includes('.')) {
        // Use cell.text but only if it's reasonable (not a formula display)
        if (!cell.text.startsWith('=') && !cell.text.includes('(')) {
          return cell.text;
        }
      }
      
      // For numbers with natural decimals, preserve them
      if (cell.value % 1 !== 0) {
        return cell.value.toString();
      }
      
      // For whole numbers, keep as integer unless format suggests decimals
      return cell.value.toString();
    } catch (e) {
      return String(cell.value);
    }
  }

  // Check cell.text for other types (like dates, strings with special formatting)
  if (cell.text && cell.text !== "") {
    return cell.text;
  }

  // Handle all other types (strings, booleans, dates, etc.)
  if (typeof cell.value === "string") {
    // Handle ISO date strings
    if (cell.value.match(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/)) {
      try {
        const date = new Date(cell.value);
        // Return formatted date (you can customize this format)
        return date.toLocaleDateString();
      } catch (e) {
        return cell.value;
      }
    }
  }
  
  return String(cell.value);
}
