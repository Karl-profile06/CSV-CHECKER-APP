const excelUploadInput = document.getElementById("excelUpload");
const uploadExcelBtn = document.getElementById("uploadExcel");
const downloadUpdatedBtn = document.getElementById("downloadUpdated");
const updateExcelTitleInput = document.getElementById("updateExcelTitle");
const clearExcelBtn = document.getElementById("clearExcel");

let uploadedWorkbook = null;

/* =========================
   UPLOAD EXCEL
========================= */
uploadExcelBtn.addEventListener("click", async () => {
    const file = excelUploadInput.files[0];
    if (!file) {
        alert("Please upload an Excel file first.");
        return;
    }

    const buffer = await file.arrayBuffer();
    uploadedWorkbook = new ExcelJS.Workbook();
    await uploadedWorkbook.xlsx.load(buffer);
    alert("Upload successful.");
});

/* =========================
   MAIN PROCESSOR
========================= */
function processSheet(sheet) {
    addHeaderFormulas(sheet);
    addDescriptionAndTotals(sheet);
}

/* =========================
   ADD FORMULAS TO FIRST ROW (GREEN)
========================= */
function addHeaderFormulas(sheet) {
    const headerRow = sheet.getRow(1);
    const totalCols = sheet.columnCount;

    for (let col = 2; col <= totalCols; col++) {
        const colLetter = sheet.getColumn(col).letter;

        headerRow.getCell(col).value = {
            formula: `=${colLetter}46+${colLetter}45+${colLetter}44+${colLetter}43+${colLetter}42+${colLetter}41+${colLetter}31+${colLetter}30+${colLetter}29+${colLetter}28+${colLetter}27+${colLetter}26+${colLetter}25+${colLetter}22+${colLetter}21+${colLetter}20+${colLetter}19+${colLetter}18+${colLetter}17+${colLetter}16+${colLetter}14-${colLetter}9-${colLetter}58`
        };

        headerRow.getCell(col).font = { name: "Calibri", bold: true };
        headerRow.getCell(col).alignment = { horizontal: "center", vertical: "middle" };
        headerRow.getCell(col).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF00FF00" } // Bright green
        };
    }

    headerRow.commit();
}

/* =========================
   ADD DESCRIPTION + TOTAL COLUMNS
========================= */
function addDescriptionAndTotals(sheet) {
    let lastDataCol = 1;
    const totalCols = sheet.columnCount;

    // Find last column with actual data
    for (let col = 1; col <= totalCols; col++) {
        const colValues = sheet.getColumn(col).values;
        if (colValues.some((v, i) => i > 0 && v !== null && v !== undefined && v !== "")) {
            lastDataCol = col;
        }
    }

    const descriptionCol = lastDataCol + 2; // skip 1 column
    const totalCol = lastDataCol + 3;

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

    // Header titles
    const headerRow = sheet.getRow(1);
    headerRow.getCell(descriptionCol).value = "DESCRIPTION";
    headerRow.getCell(totalCol).value = "TOTAL";

    [descriptionCol, totalCol].forEach(col => {
        const cell = headerRow.getCell(col);
        cell.font = { name: "Calibri", bold: true };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFFF00" } // Yellow
        };
    });

    // Insert labels (rows 9–68)
    for (let i = 0; i < labels.length; i++) {
        const rowNum = 9 + i;
        const cell = sheet.getRow(rowNum).getCell(descriptionCol);
        cell.value = labels[i];
        cell.font = { name: "Calibri", bold: true };
        cell.alignment = { horizontal: "left", vertical: "middle" };
    }

    // Add row totals (rows 9–68)
    for (let rowNum = 9; rowNum <= 68; rowNum++) {
        const cell = sheet.getRow(rowNum).getCell(totalCol);
        cell.value = {
            formula: `=SUM(A${rowNum}:${sheet.getColumn(lastDataCol).letter}${rowNum})`
        };
        cell.font = { name: "Calibri", bold: true };
        cell.alignment = { horizontal: "center", vertical: "middle" };
    }
}

/* =========================
   DOWNLOAD UPDATED EXCEL
========================= */
downloadUpdatedBtn.addEventListener("click", async () => {
    if (!uploadedWorkbook) {
        alert("Please upload an Excel file first.");
        return;
    }

    uploadedWorkbook.eachSheet(sheet => {
        processSheet(sheet);
    });

    const title = updateExcelTitleInput.value.trim() || "Updated_Excel_File";
    const buffer = await uploadedWorkbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `${title}.xlsx`;
    link.click();
});

/* =========================
   CLEAR ALL
========================= */
clearExcelBtn.addEventListener("click", () => {
    uploadedWorkbook = null;
    excelUploadInput.value = "";
    updateExcelTitleInput.value = "";
});
