# excel-to-pdf-converter

Convierte archivos Excel (.xlsx) a PDF fácilmente.

## Instalación

```
npm install excel-to-pdf-converter
```

## Uso

```js
import { convertExcelToPdf } from 'excel-to-pdf-converter';

await convertExcelToPdf('ruta/al/archivo.xlsx', 'ruta/salida.pdf');
```

## API

### convertExcelToPdf(inputFilePath, outputFileName)
Convierte un archivo Excel (.xlsx) en un PDF.
- `inputFilePath`: Ruta al archivo Excel de entrada.
- `outputFileName`: Ruta y nombre del archivo PDF de salida.

## Licencia
ISC
