import xlsx from 'xlsx';
import fs from 'fs/promises';
import fsSync from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';
import {award} from './services/award.js';

async function create_template() {
  // ✅ Define file paths
  const __filename = fileURLToPath(import.meta.url);
  const __dirname = dirname(__filename);
  const jsonDir = path.resolve(__dirname, 'Data');
  const files = await fs.readdir(jsonDir);

  // ✅ Track processing stats
  let processedCount = 0; // ← ADDED: count of successful files
  let skippedCount = 0;   // ← ADDED: count of skipped files
  let skippedFiles = [];  // ← ADDED: names of skipped files

  for (const fileName of files.filter(f => f.endsWith('.xlsx'))) {
    if (!fileName) {
      throw new Error('❌ No .xlsx file found in JsonData folder.');
    }

    const filePath = path.join(jsonDir, fileName);
    const outputFolder = path.resolve(__dirname, 'Output');
    const outputFilePath = path.join(
      outputFolder,
      `cleaned_data_${path.basename(fileName, '.xlsx')}.json`
    );

    try {
      const workbook = xlsx.readFile(filePath);
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];

      const headerGroups = {
        Project: ['Project Name', 'FMS ID', 'LLW', 'Award'],
        'Vendor Information': [
          'Vendor ID',
          'Vendor Number',
          'Contract No',
          'Contract Expiration Date',
          'Vendor Name',
          'Vendor Address',
          'City1',
          'State1',
          'Zip Code1'
        ],
        'Delivery To Information (School where work performed)': [
          'School ID / Name',
          'Address1',
          'City2',
          'State2',
          'Zip Code2',
          'Attention To: (Custodain)',
          'Attention To Phone No. (Custodian)',
          'Title'
        ],
        'Invoice  To Information': [
          'Agency',
          'Address2',
          'City3',
          'State3',
          'Zip Code3',
          'Attention To Phone No. (Borough Director)',
          'Attention To (Borough Director)'
        ],
        'Items to be purchased': [
          'Description',
          'Quantity',
          'Unit',
          '$ Unit Price',
          'Amount Owed'
        ]
      };

      // ✅ Set child headers to exclude from validation
      const excludedChildren = ['FMS ID', 'Award', 'Vendor Number', 'Vendor ID']; // ← MOVED OUTSIDE to be shared
function findHeaderCells() {
  const range = xlsx.utils.decode_range(worksheet['!ref']);
  let headerPositions = {};
  let fullPurchaseAmount = null; // Store extracted value

  for (let R = range.s.r; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cell_address = { c: C, r: R };
      const cell_ref = xlsx.utils.encode_cell(cell_address);

      if (worksheet[cell_ref] && worksheet[cell_ref].v) {
        let cellValue = worksheet[cell_ref].v.toString().trim();
        headerPositions[cellValue] = { ref: cell_ref, row: R, col: C };

        // ✅ Check for "Full Purchase Amount"
        if (cellValue.toLowerCase() === "Full Purchase Amount") {
          let targetCol = C + 1; // Start searching to the right

          while (targetCol <= range.e.c) {
            let nextCellRef = xlsx.utils.encode_cell({ c: targetCol, r: R });
            let nextCellValue = worksheet[nextCellRef]?.v;

            if (nextCellValue && !isNaN(parseFloat(nextCellValue))) {
              fullPurchaseAmount = parseFloat(nextCellValue);
              break; // Stop searching once we find the first number
            }

            targetCol++;
          }
        }
      }
    }
  }

  if (fullPurchaseAmount !== null) {
    baseData["Full Purchase Amount"] = fullPurchaseAmount;
  }

  return headerPositions;
}

      const headerCells = findHeaderCells();

      function validateRequiredHeaders(headerGroups, headerCells) {
        let missingHeaders = [];

        for (const [parent, children] of Object.entries(headerGroups)) {
          if (!excludedChildren.includes(parent)) {
            if (!headerCells[parent]) {
              missingHeaders.push(`❌ Missing parent header: ${parent}`);
            }
          }

          children.forEach(child => {
            if (!excludedChildren.includes(child)) {
              if (!headerCells[child]) {
                missingHeaders.push(`❌ Missing child header: ${child} under ${parent}`);
              }
            }
          });
        }

        if (missingHeaders.length > 0) {
          console.error(`🚫 Header validation failed for file ${fileName}:\n${missingHeaders.join('\n')}`);
        }

        console.log(`✅ Header check passed for: ${fileName}`);
      }

      validateRequiredHeaders(headerGroups, headerCells);

      let baseData = {};
      let errors = [];

      for (const [parentHeader, childHeaders] of Object.entries(headerGroups)) {
        if (parentHeader !== 'Items to be purchased' && headerCells[parentHeader]) {
          const parentRow = headerCells[parentHeader].row;

          childHeaders.forEach(childHeader => {
            if (headerCells[childHeader] && headerCells[childHeader].row === parentRow + 1) {
              const cellRef = headerCells[childHeader].ref;
              const headerCellDecoded = xlsx.utils.decode_cell(cellRef);
              const cellBelowRef = xlsx.utils.encode_cell({
                c: headerCellDecoded.c,
                r: headerCellDecoded.r + 1
              });
              const cell = worksheet[cellBelowRef];

              if (!excludedChildren.includes(childHeader)) {
                if (!cell || cell.v == null || cell.v === '') {
                  errors.push(`❌ Missing value for required field "${childHeader}" under "${parentHeader}"`);
                }
              }

              function getCellValue(cell) {
                return cell?.w || cell?.v || '';
              }
              baseData[childHeader] = getCellValue(cell);
            }
          });
        }
      }

      if (errors.length > 0) {
        console.error(`🚫 Validation Errors in file ${fileName}:\n${errors.join('\n')}`);
        skippedCount++; // ← ADDED: count skipped
        skippedFiles.push(fileName); // ← ADDED: record file name
        continue; // ← ADDED: skip to next file
      }

      // ✅ Extract items logic (unchanged)
     let items = [];
if (headerCells['Items to be purchased']) {
  const itemRow = headerCells['Items to be purchased'].row + 1;
  let rowIndex = itemRow;

  while (true) {
    let descriptionCellRef = xlsx.utils.encode_cell({
      c: headerCells['Description'].col,
      r: rowIndex
    });
    if (!worksheet[descriptionCellRef]) break;

    let item = {
      Description: worksheet[descriptionCellRef]?.v || '',
      Quantity: worksheet[xlsx.utils.encode_cell({ c: headerCells['Quantity'].col, r: rowIndex })]?.v || '',
      Unit: worksheet[xlsx.utils.encode_cell({ c: headerCells['Unit'].col, r: rowIndex })]?.v || '',
      'Unit Price':
        worksheet[xlsx.utils.encode_cell({ c: headerCells['$ Unit Price'].col, r: rowIndex })]?.v || '',
      'Amount Owed':
        worksheet[xlsx.utils.encode_cell({ c: headerCells['Amount Owed'].col, r: rowIndex })]?.v || ''
    };

    if (!item.Description || item.Description.toLowerCase() === 'description') {
      rowIndex++;
      continue;
    }

    items.push(item);
    rowIndex++;
  }
}

const MAX_ITEMS_PER_PAGE = 8;
let totalPages = Math.ceil(items.length / MAX_ITEMS_PER_PAGE);
let paginatedData = [];

for (let i = 0; i < totalPages; i++) {
  let pageItems = items.slice(i * MAX_ITEMS_PER_PAGE, (i + 1) * MAX_ITEMS_PER_PAGE);
  while (pageItems.length < MAX_ITEMS_PER_PAGE) {
    pageItems.push({
      Description: '',
      Quantity: '',
      Unit: '',
      'Unit Price': '',
      'Amount Owed': ''
    });
  }

  // Handling Award Data
  const MAX_AWARDS = 8;
  let awardString = baseData["Award"] || ""; // Extracts the existing award string
  let awardData = awardString.split(" ").filter(a => a.trim() !== ""); // Splits awards into an array
  let formattedAwards = {};

  for (let j = 0; j < awardData.length; j++) {
    formattedAwards[`Award_${j + 1}`] = awardData[j];
  }

  for (let j = awardData.length; j < MAX_AWARDS; j++) {
    formattedAwards[`Award_${j + 1}`] = "";
  }

  paginatedData.push({
    ...baseData,
    'Items to be purchased': pageItems,
    Page: `${i + 1}`,
    Of: `${totalPages}`,
    Date: new Date().toISOString().slice(0, 10),
    ...formattedAwards // Injecting Award data into the JSON structure
  });
}


   

      function cleanJson(data) {
        return data.map(entry => {
          let cleanedEntry = {};
          for (let key in entry) {
            let newKey = key
              .replace(/Project Name/, 'Project_Name')
              .replace(/Vendor ID/, 'V_I')
              .replace(/Vendor Number/, 'V')
              .replace(/Contract No/, 'Contract_Number')
              .replace(/Contract Expiration Date/, 'Term_Date')
              .replace(/Vendor Name/, 'Vendor_Name')
              .replace(/Vendor Address/, 'Vendor_Address')
              .replace(/City1/, 'V_City')
              .replace(/State1/, 'V_State')
              .replace(/Zip Code1/, 'V_Zip')
              .replace(/City2/, 'Del_C')
              .replace(/State2/, 'Del_St')
              .replace(/Zip Code2/, 'D_Z')
              .replace(/City3/, 'In_Ci')
              .replace(/State3/, 'In_St')
              .replace(/Zip Code3/, 'I_Z')
              .replace(/\s+/g, '_')
              .replace(/\//g, '_');
            cleanedEntry[newKey] = entry[key] ?? '';
          }

          if (cleanedEntry.Items_to_be_purchased) {
            const itemsFlat = cleanedEntry.Items_to_be_purchased.reduce((acc, item, idx) => {
              acc[`Description_${idx + 1}`] = item.Description;
              acc[`Quantity_${idx + 1}`] = item.Quantity;
              acc[`Unit_${idx + 1}`] = item.Unit;
              acc[`Unit_Price_${idx + 1}`] = item['Unit Price'];
              acc[`Amount_Owed_${idx + 1}`] = item['Amount Owed'];
              return acc;
            }, {});
            cleanedEntry = { ...cleanedEntry, ...itemsFlat };
            delete cleanedEntry.Items_to_be_purchased;
          }

          return cleanedEntry;
        });
      }

      if (!fsSync.existsSync(outputFolder)) {
        fsSync.mkdirSync(outputFolder);
      }

      const cleanedData = cleanJson(paginatedData);
      const enrichedData = award(cleanedData)

      await fs.writeFile(outputFilePath, JSON.stringify(enrichedData, null, 2), 'utf8');
      console.log(`✅ Cleaned data saved: ${outputFilePath}`);

      for (const entry of enrichedData) {
        const llwValue = entry.LLW || 'UNKNOWN_LLW';
        const fileNameOut = `${llwValue}_${entry.Page}of${entry.Of}.json`;
        const filePath = path.join(outputFolder, fileNameOut);
        await fs.writeFile(filePath, JSON.stringify(entry, null, 2), 'utf8');
        console.log(`✅ JSON file saved: ${filePath}`);
      }

      processedCount++; // ← ADDED: count successful
    } catch (err) {
      skippedCount++;
      skippedFiles.push(fileName);
      console.error(`❌ Failed to process ${fileName}: ${err.message}`);
    }
  }


  // ✅ Summary Log
  console.log(`\n📊 Done!`);
  console.log(`✅ Processed: ${processedCount}`);
  console.log(`🚫 Skipped: ${skippedCount}`);
  if (skippedFiles.length > 0) {
    console.log(`📝 Skipped Files: ${skippedFiles.join(', ')}`);
  }
}

create_template();