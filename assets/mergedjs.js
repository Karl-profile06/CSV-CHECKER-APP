// ======================================================
// ELEMENTS
// ======================================================
const mainCsvInput = document.getElementById("mainCsvInput");
const eodCsvInput = document.getElementById("eodCsvInput");
const processBtn = document.getElementById("processData");
const clearBtn = document.getElementById("clearAll");
const excelTitleInput = document.getElementById("finalExcelTitle");

let workbook = null;

// ======================================================
// HELPERS
// ======================================================
function sanitizeSheetName(name) {
    return name.replace(/[:\\/?*\$\$]/g, "").substring(0, 31);
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
"OTHER_DISC","REFUND_AMT","SCHRGE_AMNT","OTHER_SCHR","CASH_SLS","CARD_SLS",
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
// SET GLOBAL FONT (Calibri, size 11)
// ======================================================
function setGlobalFont(sheet) {
    sheet.eachRow({ includeEmpty: true }, (row) => {
        row.eachCell({ includeEmpty: true }, (cell) => {
            cell.font = { name: "Calibri", size: 11 };
        });
    });
}

// ======================================================
// SECOND SCRIPT LOGIC (BLACK COL + CONDITIONAL)
// ======================================================
function processSheet(sheet) {
    let descCol = null;
    let totalCol = null;

    sheet.getRow(1).eachCell((cell, colNumber) => {
        const val = cell.value ? cell.value.toString().trim().toUpperCase() : "";
        if (val === "DESCRIPTION") descCol = colNumber;
        if (val === "TOTAL") totalCol = colNumber;
    });

    if (!descCol || !totalCol) return;

    const lastDataCol = descCol - 2;

    // BLACK COLUMN BEFORE DESCRIPTION
    const blackCol1 = sheet.getColumn(descCol - 1);
    blackCol1.width = 1;
    blackCol1.eachCell({ includeEmpty: true }, cell => {
        cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF000000" }
        };
    });

    // BLACK COLUMN AFTER TOTAL
    const blackCol2 = sheet.getColumn(totalCol + 1);
    blackCol2.width = 1;
    blackCol2.eachCell({ includeEmpty: true }, cell => {
        cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF000000" }
        };
    });

    // CONDITIONAL FORMATTING (data columns 2+ only)
    const startCol = 2;
    const endCol = lastDataCol;

    if (startCol <= endCol) {
        const startLetter = sheet.getColumn(startCol).letter;
        const endLetter = sheet.getColumn(endCol).letter;

        sheet.addConditionalFormatting({
            ref: `${startLetter}1:${endLetter}1`,
            rules: [
                {
                    type: "expression",
                    formulae: [`OR(${startLetter}1>=1, ${startLetter}1<=-1)`],
                    style: {
                        fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFF0000" } },
                        font: { color: { argb: "FFFFFFFF" }, bold: true }
                    }
                },
                {
                    type: "expression",
                    formulae: [`AND(${startLetter}1<1, ${startLetter}1>-1)`],
                    style: {
                        fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FF00FF00" } }
                    }
                }
            ]
        });
    }
}

// ======================================================
// MAIN PROCESS
// ======================================================
processBtn.addEventListener("click", async () => {
    if (!mainCsvInput.files.length) return alert("Upload main CSV/TXT files first.");

    const mainFilesData = await readCsvFiles(mainCsvInput.files);
    const eodFilesData = eodCsvInput.files.length ? await readCsvFiles(eodCsvInput.files) : [];

    workbook = new ExcelJS.Workbook();

    mainFilesData.sort((a, b) => new Date(a.sheetName) - new Date(b.sheetName));

    mainFilesData.forEach(file => {
        let sheet = workbook.addWorksheet(file.sheetName);

        file.data.forEach(row => sheet.addRow(row));

        const lastDataCol = Math.max(...file.data.map(r => r.length));
        const blackCol1 = lastDataCol + 1;
        const descCol = blackCol1 + 1;
        const totalCol = descCol + 1;
        const blackCol2 = totalCol + 1;

        sheet.views = [{ state: "frozen", xSplit: 1 }]; // ✅ Freeze first column

        const headerRow = sheet.getRow(1);

        // ✅ FIRST COLUMN (A1) = LABEL, NO FORMULA, BOLD
        headerRow.getCell(1).font = { name: "Calibri", size: 11, bold: true };

        // TOP FORMULA (columns 2+ only, skip column 1)
        for (let c = 2; c <= lastDataCol; c++) {
            const letter = sheet.getColumn(c).letter;
            headerRow.getCell(c).value = {
                formula: `=${letter}46+${letter}45+${letter}44+${letter}43+${letter}42+${letter}41+${letter}31+${letter}30+${letter}29+${letter}28+${letter}27+${letter}26+${letter}25+${letter}22+${letter}21+${letter}20+${letter}19+${letter}18+${letter}17+${letter}16+${letter}14-${letter}9-${letter}58`
            };
            headerRow.getCell(c).fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FF00FF00" }
            };
        }

        // DESCRIPTION HEADER
        const descHeader = headerRow.getCell(descCol);
        descHeader.value = "DESCRIPTION";
        descHeader.font = { name: "Calibri", size: 11, bold: true };
        descHeader.alignment = { horizontal: "center" };
        descHeader.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFFF00" }
        };

        // TOTAL HEADER
        const totalHeader = headerRow.getCell(totalCol);
        totalHeader.value = "TOTAL";
        totalHeader.font = { name: "Calibri", size: 11, bold: true };
        totalHeader.alignment = { horizontal: "center" };
        totalHeader.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFFF00" }
        };

        headerRow.commit();

        // ✅ GLOBAL FONT + FIRST COLUMN & DESCRIPTION BOLD
        setGlobalFont(sheet);
        
        // Make first column bold (all rows)
        sheet.getColumn(1).eachCell({ includeEmpty: true }, cell => {
            cell.font = { name: "Calibri", size: 11, bold: true };
        });

        // Make DESCRIPTION column bold (rows 9+)
        descriptionLabels.forEach((label, i) => {
            const row = sheet.getRow(9 + i);
            row.getCell(descCol).value = label;
            row.getCell(descCol).font = { name: "Calibri", size: 11, bold: true };
        });

        // TOTAL FORMULA
        for (let r = 9; r <= sheet.rowCount; r++) {
            sheet.getRow(r).getCell(totalCol).value = {
                formula: `SUM(B${r}:${sheet.getColumn(lastDataCol).letter}${r})` // Start from B (skip A)
            };
        }

        processSheet(sheet);

        // WIDTHS
        sheet.getColumn(blackCol1).width = 1;
        sheet.getColumn(blackCol2).width = 1;
        for (let i = 1; i <= totalCol + 30; i++) {
            if (i !== blackCol1 && i !== blackCol2) {
                sheet.getColumn(i).width = 15;
            }
        }

        sheet._eodStartCol = blackCol2 + 1;
    });

    // EOD APPEND
    eodFilesData.forEach(file => {
        let sheet = workbook.getWorksheet(file.sheetName);
        if (!sheet) sheet = workbook.addWorksheet(file.sheetName);

        const startCol = sheet._eodStartCol || 1;

        file.data.forEach((row, rIdx) => {
            const excelRow = sheet.getRow(rIdx + 3);
            row.forEach((cell, cIdx) => {
                excelRow.getCell(startCol + cIdx).value = cell;
            });
        });
    });

    // DOWNLOAD
    const title = excelTitleInput.value.trim() || "Final_Excel";
    const buffer = await workbook.xlsx.writeBuffer();

    const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `${title}.xlsx`;
    link.click();
});

// ======================================================
// CLEAR
// ======================================================
clearBtn.addEventListener("click", () => {
    workbook = null;
    mainCsvInput.value = "";
    eodCsvInput.value = "";
    excelTitleInput.value = "";
});
