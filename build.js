import esbuild from 'esbuild';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const buildConfig = {
  entryPoints: ['src/lib/excel-to-pdf.js'],
  bundle: true,
  platform: 'node',
  target: 'node16',
  format: 'esm',
  outfile: 'dist/excel-to-pdf.js',
  external: [
    'exceljs',
    'jspdf', 
    'fs'
  ],
  minify: process.env.NODE_ENV === 'production',
  sourcemap: true,
  resolveExtensions: ['.js'],
  loader: {
    '.js': 'js'
  }
};

async function build() {
  try {
    console.log('Building excel-to-pdf-converter...');
    
    await esbuild.build(buildConfig);
    
    console.log('✅ Build completed successfully!');
    console.log('📦 Output: dist/excel-to-pdf.js');
    
  } catch (error) {
    console.error('❌ Build failed:', error);
    process.exit(1);
  }
}

build();