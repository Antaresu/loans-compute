import fs from 'fs';
import path from 'path';
import * as XLSX from 'xlsx';

// CLI helper to read any Excel file row-by-row and column-by-column
function run() {
  const args = process.argv.slice(2);
  let filePath = args[0] || './sample_loans.xlsx';

  if (!fs.existsSync(filePath)) {
    console.error(`❌ Hindi makita ang file: ${filePath}`);
    console.log(`\nGamit: node cli-reader.js <path-to-excel.xlsx> [--json] [--row <number>] [--sheet <name>]`);
    process.exit(1);
  }

  const fileBuffer = fs.readFileSync(filePath);
  const workbook = XLSX.read(fileBuffer, { type: 'buffer', cellDates: true });

  console.log(`\n======================================================`);
  console.log(`📗 SYSGEN RPO EXCEL READER`);
  console.log(`📁 File: ${path.basename(filePath)} (${(fileBuffer.length / 1024).toFixed(1)} KB)`);
  console.log(`📑 Sheet Count: ${workbook.SheetNames.length}`);
  console.log(`📋 Worksheets: [ ${workbook.SheetNames.join(', ')} ]`);
  console.log(`======================================================\n`);

  let targetSheet = workbook.SheetNames[0];
  const sheetArgIdx = args.indexOf('--sheet');
  if (sheetArgIdx !== -1 && args[sheetArgIdx + 1]) {
    targetSheet = args[sheetArgIdx + 1];
  }

  const ws = workbook.Sheets[targetSheet];
  if (!ws) {
    console.error(`❌ Hindi makita ang worksheet na "${targetSheet}"`);
    process.exit(1);
  }

  const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: false });
  console.log(`🔹 Sheet: "${targetSheet}" | Total Rows: ${rawRows.length}\n`);

  // Row Inspection Option
  const rowArgIdx = args.indexOf('--row');
  if (rowArgIdx !== -1 && args[rowArgIdx + 1]) {
    const targetRowNum = parseInt(args[rowArgIdx + 1], 10);
    const rowIdx = targetRowNum - 1; // 1-based to 0-based
    if (rawRows[rowIdx]) {
      const headers = rawRows[0] || [];
      const row = rawRows[rowIdx];
      console.log(`🔍 [INSPECT ROW #${targetRowNum}]`);
      headers.forEach((h, i) => {
        console.log(`  • ${h || `Col ${i+1}`}: ${row[i] !== undefined ? row[i] : ''}`);
      });
      return;
    } else {
      console.log(`❌ Walang row #${targetRowNum}`);
      return;
    }
  }

  // JSON Option
  if (args.includes('--json')) {
    const jsonObjects = XLSX.utils.sheet_to_json(ws);
    console.log(JSON.stringify(jsonObjects, null, 2));
    return;
  }

  // Display Table summary
  const headers = rawRows[0] || [];
  console.log(`🏷️ COLUMNS (${headers.length}):`);
  headers.forEach((h, idx) => {
    console.log(`  [${String.fromCharCode(65 + (idx % 26))}] ${h}`);
  });

  console.log(`\n📊 ROWS SAMPLE (First 5 rows):`);
  const sampleRows = rawRows.slice(1, 6);
  sampleRows.forEach((r, idx) => {
    console.log(`\n--- [Row #${idx + 2}] ---`);
    headers.forEach((h, cIdx) => {
      console.log(`  ${h}: ${r[cIdx] !== undefined ? r[cIdx] : ''}`);
    });
  });

  console.log(`\n✅ Ready! Gamitin ang 'npm run dev' para buksan ang buong visual web interface.`);
}

run();
