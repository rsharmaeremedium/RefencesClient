/**
 * transformDeals.js
 * ------------------
 * Browser-side port of transform_deals.py, built for EREMEDIUM
 * (https://rsharmaeremedium.github.io/RefencesClient/).
 *
 * Converts a HubSpot CRM deals export into the standard "reference client"
 * format:
 *   Deal Name | Number of Screens | Product Category | Speciality |
 *   Deal owner | City | Area | Address | Pin Code
 *
 * Requires SheetJS (the `XLSX` global) to already be loaded on the page,
 * e.g. via:
 *   <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
 *
 * All processing happens locally in the browser — no data leaves the device,
 * consistent with the rest of the EREMEDIUM app.
 */

const OUTPUT_COLUMNS = [
  "Deal Name",
  "Number of Screens",
  "Product Category",
  "Speciality",
  "Deal owner",
  "City",
  "Area",
  "Address",
  "Pin Code",
];

/**
 * Replace every semicolon with a comma in a string value.
 * Collapses " ;" -> "," and ";" -> "," the same way the Python version does.
 */
function cleanSemicolons(value) {
  if (typeof value !== "string") return value;
  return value.split(" ;").join(",").split(";").join(",");
}

/**
 * Read a File/Blob (from an <input type="file"> or drop event) and return
 * a Promise that resolves to a SheetJS workbook object.
 */
function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array", cellDates: false });
        resolve(workbook);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Core transform: takes a SheetJS workbook (and optional sheet name) and
 * returns { rows, missingColumns } where rows is an array of row objects
 * in OUTPUT_COLUMNS order, ready to re-export.
 */
function buildOutputRows(workbook, sheetName) {
  const targetSheet = sheetName || workbook.SheetNames[0];
  const sheet = workbook.Sheets[targetSheet];
  if (!sheet) {
    throw new Error(
      `Sheet "${targetSheet}" not found. Available sheets: ${workbook.SheetNames.join(", ")}`
    );
  }

  // raw:false -> formatted strings (keeps Pin Code from becoming a lossy number)
  const json = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });

  if (json.length === 0) {
    return { rows: [], missingColumns: [] };
  }

  const sourceColumns = Object.keys(json[0]);
  const missingColumns = OUTPUT_COLUMNS.filter((c) => !sourceColumns.includes(c));
  if (missingColumns.length > 0) {
    throw new Error(
      `Input file is missing required column(s): ${missingColumns.join(", ")}`
    );
  }

  const rows = json.map((sourceRow) => {
    const outRow = {};
    for (const col of OUTPUT_COLUMNS) {
      let value = sourceRow[col];
      value = cleanSemicolons(value);
      outRow[col] = value;
    }
    // Keep Number of Screens as a clean integer when possible.
    const screens = outRow["Number of Screens"];
    const asNumber = Number(screens);
    if (screens !== "" && !Number.isNaN(asNumber)) {
      outRow["Number of Screens"] = Math.trunc(asNumber);
    }
    return outRow;
  });

  return { rows, missingColumns: [] };
}

/**
 * Build a downloadable workbook (SheetJS) from the transformed rows,
 * with column widths similar to the Python version.
 */
function buildOutputWorkbook(rows) {
  const worksheet = XLSX.utils.json_to_sheet(rows, { header: OUTPUT_COLUMNS });

  worksheet["!cols"] = [
    { wch: 38 }, // Deal Name
    { wch: 10 }, // Number of Screens
    { wch: 16 }, // Product Category
    { wch: 28 }, // Speciality
    { wch: 18 }, // Deal owner
    { wch: 16 }, // City
    { wch: 22 }, // Area
    { wch: 50 }, // Address
    { wch: 10 }, // Pin Code
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Data");
  return workbook;
}

/**
 * End-to-end helper: pass in a File object (e.g. from a file input or
 * drag-and-drop), get back a downloaded .xlsx named `outputFileName`.
 *
 * Usage:
 *   fileInput.addEventListener("change", (e) => {
 *     transformDealsFile(e.target.files[0]);
 *   });
 */
async function transformDealsFile(file, outputFileName, sheetName) {
  const workbook = await readWorkbook(file);
  const { rows } = buildOutputRows(workbook, sheetName);
  const outWorkbook = buildOutputWorkbook(rows);

  const name =
    outputFileName ||
    file.name.replace(/\.(xlsx|xls|csv|ods)$/i, "") + "_formatted.xlsx";

  XLSX.writeFile(outWorkbook, name);
  return rows.length;
}

// Expose for use in a plain <script> tag (non-module) context.
if (typeof window !== "undefined") {
  window.EremediumTransform = {
    transformDealsFile,
    buildOutputRows,
    buildOutputWorkbook,
    readWorkbook,
    OUTPUT_COLUMNS,
  };
}

// Also support ES module / CommonJS import if the page uses bundling.
if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    transformDealsFile,
    buildOutputRows,
    buildOutputWorkbook,
    readWorkbook,
    OUTPUT_COLUMNS,
  };
}
