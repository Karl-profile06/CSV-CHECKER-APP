// ======================================================
// EXCEL MANAGER – MAIN STATE & ELEMENTS
// ======================================================

const mainCsvInput = document.getElementById("mainCsvInput");
const eodCsvInput = document.getElementById("eodCsvInput");
const processBtn = document.getElementById("processData");
const clearBtn = document.getElementById("clearAll");
const excelTitleInput = document.getElementById("finalExcelTitle");

let mainFilesData = [];
let eodFilesData = [];
let workbook = null;

// ======================================================
// HELPER FUNCTIONS
// ======================================================

function sanitizeSheetName(name) {
    return name.replace(/[:\\/?*\[\]]/g, "").substring(0, 31);
}

function parseTextToRows(text) {
    return text
        .split(/\r?\n/)
        .filter(line => line.trim())
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

// Promise-based file reader
function readCsvFiles(files) {
    return Promise.all([...files].map(file => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = e => {
                const data = parseTextToRows(e.target.result);
                const trnDate = extractTRNDate(data) || sanitizeSheetName(file.name.replace(/\.[^/.]+$/, ""));
                resolve({ sheetName: trnDate, data });
            };
            reader.onerror = reject;
            reader.readAsText(file);
        });
    }));
}

// ======================================================
// DESCRIPTION LABELS
// ======================================================
const descriptionLabels = [
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

// ======================================================
// PROCESS & DOWNLOAD
// ======================================================

processBtn.addEventListener("click", async () => {
    if (!mainCsvInput.files.length) return alert("Upload main CSV/TXT files first.");

    mainFilesData = await readCsvFiles(mainCsvInput.files);
    eodFilesData = eodCsvInput.files.length ? await readCsvFiles(eodCsvInput.files) : [];

    workbook = new ExcelJS.Workbook();

    // Sort main CSV sheets by date
    mainFilesData.sort((a, b) => new Date(a.sheetName) - new Date(b.sheetName));

    // ======================= TOP FORMULA AND STYLING =======================

mainFilesData.forEach(file => {
    const sheet = workbook.addWorksheet(file.sheetName);

    // Add main CSV data
    file.data.forEach((row, rIdx) => sheet.addRow(row));

    const lastDataCol = Math.max(...file.data.map(r => r.length));
    const emptyAfterDataCol = lastDataCol + 1;
    const descCol = emptyAfterDataCol + 1;
    const totalCol = descCol + 1;
    const emptyAfterTotalCol = totalCol + 1;

    // Freeze first column
    sheet.views = [{ state: "frozen", xSplit: 1 }];

    // ---------------- TOP FORMULA (row 1) ----------------
const headerRow = sheet.getRow(1);
for (let c = 2; c <= lastDataCol; c++) {
    const letter = sheet.getColumn(c).letter;
    headerRow.getCell(c).value = {
        formula: `=${letter}46+${letter}45+${letter}44+${letter}43+${letter}42+${letter}41+${letter}31+${letter}30+${letter}29+${letter}28+${letter}27+${letter}26+${letter}25+${letter}22+${letter}21+${letter}20+${letter}19+${letter}18+${letter}17+${letter}16+${letter}14-${letter}9-${letter}58`
    };
    headerRow.getCell(c).font = { name: "Calibri", bold: true };
    headerRow.getCell(c).alignment = { horizontal: "center" };
    
    // ✅ Highlight formula cells with lime color
    headerRow.getCell(c).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF00FF00" } // Lime green
    };
}

    // ---------------- FIRST COLUMN STYLING ----------------
    sheet.getColumn(1).eachCell(cell => {
        cell.font = { name: "Calibri", bold: true };
        cell.alignment = { vertical: "middle", horizontal: "left" };
    });

    headerRow.commit();

    // ---------------- DESCRIPTION & TOTAL ----------------
    headerRow.getCell(descCol).value = "DESCRIPTION";
    headerRow.getCell(totalCol).value = "TOTAL";
    [descCol, totalCol].forEach(c => {
        const cell = headerRow.getCell(c);
        cell.font = { name: "Calibri", bold: true };
        cell.alignment = { horizontal: "center" };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
    });
    headerRow.commit();

    // DESCRIPTION labels starting at row 9
    descriptionLabels.forEach((label, i) => {
        const row = sheet.getRow(9 + i);
        row.getCell(descCol).value = label;
        row.getCell(descCol).font = { name: "Calibri", bold: true };
        row.commit();
    });

    // TOTAL formulas starting from row 9 downward
    for (let r = 9; r <= sheet.rowCount; r++) {
        const row = sheet.getRow(r);
        row.getCell(totalCol).value = { formula: `SUM(A${r}:${sheet.getColumn(lastDataCol).letter}${r})` };
        row.getCell(totalCol).font = { name: "Calibri", bold: true };
        row.commit();
    }

    // Set uniform width + room for EOD
    const totalCols = emptyAfterTotalCol + 30;
    for (let i = 1; i <= totalCols; i++) sheet.getColumn(i).width = 15;

    sheet._eodStartCol = emptyAfterTotalCol + 1;
});


    // Append EOD CSVs starting at row 3
    eodFilesData.forEach(file => {
        let sheet = workbook.getWorksheet(file.sheetName);
        if (!sheet) sheet = workbook.addWorksheet(file.sheetName);
        const startCol = sheet._eodStartCol || 1;

        file.data.forEach((row, rIdx) => {
            const excelRow = sheet.getRow(rIdx + 3); // EOD starts at row 3
            row.forEach((cell, cIdx) => {
                excelRow.getCell(startCol + cIdx).value = cell;
            });
            excelRow.commit();
        });
    });

    

    // Download final workbook
    const title = excelTitleInput.value.trim() || "Final_Excel";
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `${title}.xlsx`;
    link.click();
});

// ======================================================
// CLEAR ALL
// ======================================================
clearBtn.addEventListener("click", () => {
    mainFilesData = [];
    eodFilesData = [];
    workbook = null;
    mainCsvInput.value = "";
    eodCsvInput.value = "";
    excelTitleInput.value = "";
});
