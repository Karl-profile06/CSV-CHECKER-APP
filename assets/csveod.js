// =========================
// CSVEOD.JS
// =========================

const excelUploadExisting = document.getElementById("excelUploadExisting");
const multiUploadFiles = document.getElementById("multiUploadFiles");
const multiUploadTitle = document.getElementById("multiUploadTitle");

const uploadExcelButton = document.getElementById("uploadExcelButton");
const uploadMultiple = document.getElementById("uploadMultiple");
const downloadMultiple = document.getElementById("downloadMultiple");
const clearMultiple = document.getElementById("clearMultiple");

let mainWorkbook = null; // ExcelJS Workbook
let csvFilesData = [];   // Array of {trnDate, data}

// =========================
// HELPER FUNCTIONS
// =========================

// Remove invalid characters from sheet names
function sanitizeSheetName(name) {
    return name.replace(/[:\\/?*\[\]]/g, "").substring(0, 31);
}

// Parse CSV/TXT to array of arrays
function parseTextToRows(text) {
    return text
        .split(/\r?\n/)
        .filter(line => line.trim())
        .map(line =>
            line.split(/[\t,]/).map(cell => {
                const trimmed = cell.trim();
                return trimmed !== "" && !isNaN(trimmed)
                    ? Number(trimmed)
                    : trimmed;
            })
        );
}

// Extract TRN_DATE from CSV/TXT (first column TRN_DATE)
function extractTRNDate(data) {
    for (const row of data) {
        if (row[0] === "TRN_DATE" && row[1]) {
            return row[1].toString(); // format YYYY-MM-DD
        }
    }
    return null;
}

// Find last used column in a sheet
function getLastUsedColumn(sheet) {
    let lastCol = 1;
    for (let col = 1; col <= sheet.columnCount; col++) {
        const colValues = sheet.getColumn(col).values;
        if (colValues.some((v, i) => i > 0 && v !== null && v !== undefined && v !== "")) {
            lastCol = col;
        }
    }
    return lastCol;
}

// =========================
// UPLOAD EXISTING EXCEL
// =========================
uploadExcelButton.addEventListener("click", async () => {
    const file = excelUploadExisting.files[0];
    if (!file) {
        alert("Please upload an Excel file first.");
        return;
    }

    const buffer = await file.arrayBuffer();
    mainWorkbook = new ExcelJS.Workbook();
    await mainWorkbook.xlsx.load(buffer);
    alert("Excel uploaded successfully!");
});

// =========================
// UPLOAD MULTIPLE CSV/TXT FILES
// =========================
uploadMultiple.addEventListener("click", () => {
    if (!mainWorkbook) {
        alert("Please upload the main Excel file first.");
        return;
    }

    const files = [...multiUploadFiles.files];
    if (!files.length) {
        alert("Please select at least one CSV/TXT file.");
        return;
    }

    csvFilesData = []; // clear previous

    files.forEach(file => {
        const reader = new FileReader();
        reader.onload = e => {
            const text = e.target.result;
            const data = parseTextToRows(text);
            const trnDate = extractTRNDate(data) || file.name.replace(/\.[^/.]+$/, "");
            csvFilesData.push({ trnDate, data });
        };
        reader.readAsText(file, "UTF-8");
    });

    alert(`${files.length} CSV/TXT file(s) uploaded and ready to process.`);
});

// =========================
// PROCESS & DOWNLOAD COMBINED EXCEL
// =========================
downloadMultiple.addEventListener("click", async () => {
    if (!mainWorkbook) {
        alert("Please upload the main Excel file first.");
        return;
    }

    if (!csvFilesData.length) {
        alert("Please upload CSV/TXT files to append.");
        return;
    }

    // Process each CSV
    csvFilesData.forEach(file => {
        const sheetName = sanitizeSheetName(file.trnDate);
        let sheet = mainWorkbook.getWorksheet(sheetName);
        if (!sheet) {
            // If sheet does not exist, create new
            sheet = mainWorkbook.addWorksheet(sheetName);
        }

        const lastCol = getLastUsedColumn(sheet);
        const startCol = lastCol + 2; // skip 1 column

        // Append CSV data starting at row 3
        file.data.forEach((row, rowIndex) => {
            const excelRow = sheet.getRow(rowIndex + 3); // row 3 start
            row.forEach((cell, colIndex) => {
                excelRow.getCell(startCol + colIndex).value = cell;
            });
            excelRow.commit();
        });
    });

    // Download workbook
    const title = multiUploadTitle.value.trim() || "Combined_Excel";
    const buffer = await mainWorkbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `${title}.xlsx`;
    link.click();
});


// =========================
// CLEAR ALL
// =========================
clearMultiple.addEventListener("click", () => {
    mainWorkbook = null;
    csvFilesData = [];
    excelUploadExisting.value = "";
    multiUploadFiles.value = "";
    multiUploadTitle.value = "";
});
