// ======================================================
// EXCEL MANAGER – MAIN STATE & ELEMENTS
// ======================================================

const excelUploadInput = document.getElementById("excelUpload");
const multiUploadFiles = document.getElementById("multiUploadFiles");
const updateExcelTitleInput = document.getElementById("finalExcelTitle");

const uploadAllBtn = document.getElementById("uploadAll"); // <-- new combined button
const processDownloadBtn = document.getElementById("processDownload");
const clearExcelBtn = document.getElementById("clearAll");

let uploadedWorkbook = null;     // Loaded Excel workbook
let csvFilesData = [];           // Parsed CSV/TXT data

// ======================================================
// HELPER FUNCTIONS
// ======================================================

function sanitizeSheetName(name) {
    return name.replace(/[:\\/?*\[\]]/g, "").substring(0, 31);
}

function parseTextToRows(text) {
    return text
        .split(/\r?\n/).filter(line => line.trim())
        .map(line =>
            line.split(/[\t,]/).map(cell => {
                const t = cell.trim();
                return t !== "" && !isNaN(t) ? Number(t) : t;
            })
        );
}

function extractTRNDate(data) {
    for (const row of data) {
        if (row[0] === "TRN_DATE" && row[1]) return row[1].toString();
    }
    return null;
}

function getLastUsedDataColumn(sheet) {
    let lastCol = 1;
    for (let col = 1; col <= sheet.columnCount; col++) {
        const values = sheet.getColumn(col).values;
        if (values.some((v, i) => i > 1 && v !== null && v !== "")) lastCol = col;
    }
    return lastCol;
}

// ======================================================
// COMBINED UPLOAD BUTTON
// ======================================================

uploadAllBtn.addEventListener("click", async () => {

    // 1️⃣ Upload Excel
    const excelFile = excelUploadInput.files[0];
    if (!excelFile) return alert("Please upload an Excel file first.");

    const buffer = await excelFile.arrayBuffer();
    uploadedWorkbook = new ExcelJS.Workbook();
    await uploadedWorkbook.xlsx.load(buffer);

    // 2️⃣ Upload CSV/TXT files
    const files = [...multiUploadFiles.files];
    csvFilesData = [];

    if (files.length) {
        await Promise.all(
            files.map(file => new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = e => {
                    const data = parseTextToRows(e.target.result);
                    const trnDate = extractTRNDate(data) || file.name.replace(/\.[^/.]+$/, "");
                    csvFilesData.push({ trnDate, data });
                    resolve();
                };
                reader.onerror = reject;
                reader.readAsText(file);
            }))
        );
        alert(`Excel uploaded successfully. ${files.length} CSV/TXT file(s) loaded.`);
    } else {
        alert("Excel uploaded successfully. No CSV/TXT files selected.");
    }
});

// ======================================================
// PROCESS EACH EXCEL SHEET
// ======================================================

function processSheet(sheet) {
    const lastDataCol = getLastUsedDataColumn(sheet);
    const descCol = lastDataCol + 2;
    const totalCol = lastDataCol + 3;

    addHeaderFormulas(sheet, lastDataCol);
    addDescriptionAndTotals(sheet, lastDataCol, descCol, totalCol);
}

// ======================================================
// ADD HEADER FORMULAS
// ======================================================

function addHeaderFormulas(sheet, lastDataCol) {
    const headerRow = sheet.getRow(1);
    for (let col = 2; col <= lastDataCol; col++) {
        const letter = sheet.getColumn(col).letter;
        headerRow.getCell(col).value = {
            formula: `=${letter}46+${letter}45+${letter}44+${letter}43+${letter}42+${letter}41+${letter}31+${letter}30+${letter}29+${letter}28+${letter}27+${letter}26+${letter}25+${letter}22+${letter}21+${letter}20+${letter}19+${letter}18+${letter}17+${letter}16+${letter}14-${letter}9-${letter}58`
        };
        headerRow.getCell(col).font = { name: "Calibri", bold: true };
        headerRow.getCell(col).alignment = { horizontal: "center" };
        headerRow.getCell(col).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF00FF00" } };
    }
    headerRow.commit();
}

// ======================================================
// ADD DESCRIPTION & TOTAL COLUMNS
// ======================================================

function addDescriptionAndTotals(sheet, lastDataCol, descCol, totalCol) {
    const labels = [
        "GROSS_SLS","VAT_AMNT","VATABLE_SLS","NONVAT_SLS","VATEXEMPT_SLS","VATEXEMPT_AMNT",
        "LOCAL_TAX","PWD_DISC","SNRCIT_DISC","EMPLO_DISC","AYALA_DISC","STORE_DISC",
        "OTHER_DISC","REFUND_AMT","SCHRGE_AMT","OTHER_SCHR","CASH_SLS","CARD_SLS",
        "EPAY_SLS","DCARD_SLS","OTHERSL_SLS","CHECK_SLS","GC_SLS","MASTERCARD_SLS",
        "VISA_SLS","AMEX_SLS","DINERS_SLS","JCB_SLS","GCASH_SLS","PAYMAYA_SLS",
        "ALIPAY_SLS","WECHAT_SLS","GRAB_SLS","FOODPANDA_SLS","MASTERDEBIT_SLS",
        "VISADEBIT_SLS","PAYPAL_SLS","ONLINE_SLS","OPEN_SALES","OPEN_SALES_2",
        "OPEN_SALES_3","OPEN_SALES_4","OPEN_SALES_5","OPEN_SALES_6","OPEN_SALES_7",
        "OPEN_SALES_8","OPEN_SALES_9","OPEN_SALES_10","OPEN_SALES_11","GC_EXCESS",
        "MOBILE_NO","NO_CUST","TRN_TYPE","SLS_FLAG","VAT_PCT","QTY_SLD","QTY",
        "ITEMCODE","PRICE","LDISC"
    ];

    const headerRow = sheet.getRow(1);
    headerRow.getCell(descCol).value = "DESCRIPTION";
    headerRow.getCell(totalCol).value = "TOTAL";
    [descCol, totalCol].forEach(col => {
        const cell = headerRow.getCell(col);
        cell.font = { name: "Calibri", bold: true };
        cell.alignment = { horizontal: "center" };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
    });
    headerRow.commit();

    labels.forEach((label, i) => {
        const row = sheet.getRow(9 + i);
        row.getCell(descCol).value = label;
        row.getCell(descCol).font = { name: "Calibri", bold: true };
        row.commit();
    });

    for (let rowNum = 9; rowNum <= 68; rowNum++) {
        const row = sheet.getRow(rowNum);
        row.getCell(totalCol).value = { formula: `=SUM(A${rowNum}:${sheet.getColumn(lastDataCol).letter}${rowNum})` };
        row.getCell(totalCol).font = { name: "Calibri", bold: true };
        row.commit();
    }
}

// ======================================================
// PROCESS ALL & DOWNLOAD FINAL EXCEL
// ======================================================

processDownloadBtn.addEventListener("click", async () => {
    if (!uploadedWorkbook) return alert("Upload Excel first.");

    uploadedWorkbook.eachSheet(sheet => processSheet(sheet));

    csvFilesData.forEach(file => {
        const sheetName = sanitizeSheetName(file.trnDate);
        const sheet = uploadedWorkbook.getWorksheet(sheetName) || uploadedWorkbook.addWorksheet(sheetName);
        const lastDataCol = getLastUsedDataColumn(sheet);
        const eodStartCol = lastDataCol + 2;

        file.data.forEach((row, rowIndex) => {
            const excelRow = sheet.getRow(rowIndex + 3);
            row.forEach((cell, colIndex) => {
                excelRow.getCell(eodStartCol + colIndex).value = cell;
            });
            excelRow.commit();
        });
    });

    const title = updateExcelTitleInput.value.trim() || "Updated_Excel";
    const buffer = await uploadedWorkbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `${title}.xlsx`;
    link.click();
});

// ======================================================
// CLEAR ALL STATE
// ======================================================

clearExcelBtn.addEventListener("click", () => {
    uploadedWorkbook = null;
    csvFilesData = [];
    excelUploadInput.value = "";
    multiUploadFiles.value = "";
    updateExcelTitleInput.value = "";
});
