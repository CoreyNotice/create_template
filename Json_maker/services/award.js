import xlsx from 'xlsx';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

export function award(cleanedData) {
  const __filename = fileURLToPath(import.meta.url);
  const __dirname = dirname(__filename);
  const cpFilePath = path.resolve(__dirname, 'cp_number.xlsx');

  const workbook = xlsx.readFile(cpFilePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const cpData = xlsx.utils.sheet_to_json(sheet);

  function cleanOmbCp(value) {
    const num = Number(value);
    return Number.isFinite(num) && num !== 0 ? value.toString().trim() : 'N/A';
  }

  // Step 1: Build structured array (n x 3)
  const structuredArray = cpData.map(row => [
    row['LLW']?.toString().trim() || '',
    cleanOmbCp(row['OMB CP #']),
    row['Award Number']?.toString().trim() || ''
  ]);

  // Step 2: Convert structuredArray to a quick lookup map for LLW
  const llwMap = {};
  for (const [llw, cp, award] of structuredArray) {
    llwMap[llw] = { cp, award };
  }

  // Step 3: Add CP value to cleanedData by matching LLW
  const enrichedData = cleanedData.map(entry => {
    const llw = entry.LLW?.toString().trim();
    const match = llwMap[llw];

    return {
      ...entry,
      CP: match?.cp || 'N/A'
    };
  });

  console.log('✅ LLW Match & Enrichment complete. Sample:\n', enrichedData[0]);

  return enrichedData;
}

// Allow script to run standalone
if (process.argv[1] === fileURLToPath(import.meta.url)) {
  award([]);
}