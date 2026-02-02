const fileInput = document.getElementById("fileInput");
const generateBtn = document.getElementById("generateExcel");
const clearBtn = document.getElementById("clearAll");
const excelTitleInput = document.getElementById("excelTitle");

let filesData = [];

/* =========================
   FILE UPLOAD
========================= */
fileInput.addEventListener("change", () => {
    [...fileInput.files].forEach(file => readFile(file));
});

/* =========================
   READ TEXT FILE
========================= */
function readFile(file) {
    const reader = new FileReader();

    reader.onload = e => {
        const text = e.target.result;
        const data = parseTextToRows(text);
        const sheetName = extractTRNDate(data) || fallbackSheetName(file.name);

        filesData.push({ sheetName, data });
    };

    reader.readAsText(file, "UTF-8");
}

/* =========================
   PARSE TEXT INTO ROWS
========================= */
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

/* =========================
   EXTRACT TRN_DATE
========================= */
function extractTRNDate(data) {
    for (const row of data) {
        if (row[0] === "TRN_DATE" && row[1]) {
            return sanitizeSheetName(row[1].toString());
        }
    }
    return null;
}

function fallbackSheetName(filename) {
    return sanitizeSheetName(filename.replace(/\.[^/.]+$/, ""));
}

function sanitizeSheetName(name) {
    return name.replace(/[:\\/?*\[\]]/g, "").substring(0, 31);
}

/* =========================
   GENERATE EXCEL
========================= */
generateBtn.addEventListener("click", async () => {
    if (!filesData.length) {
        alert("Please upload at least one text file.");
        return;
    }

    const title = excelTitleInput.value.trim() || "Converted_File";
    const workbook = new ExcelJS.Workbook();

    /* ✅ SORT SHEETS BY DATE */
    filesData.sort((a, b) => new Date(a.sheetName) - new Date(b.sheetName));

    filesData.forEach(file => {
        const sheet = workbook.addWorksheet(file.sheetName);

        /* ✅ FREEZE FIRST COLUMN */
        sheet.views = [{ state: "frozen", xSplit: 1 }];
  
        /* ADD DATA */
        file.data.forEach(row => sheet.addRow(row));

        /* BOLD FIRST COLUMN (KEEP CALIBRI) */
        sheet.getColumn(1).eachCell(cell => {
            cell.font = {
                name: "Calibri",
                bold: true
            };
            cell.alignment = {
                vertical: "middle",
                horizontal: "left"
            };
        });

 /* SET ALL COLUMNS TO SAME WIDTH + ADD 10 EXTRA EMPTY COLUMNS */
const uniformWidth = 15; // adjust as needed
const maxCols = Math.max(...file.data.map(row => row.length));
const totalCols = maxCols + 30;

for (let i = 1; i <= totalCols; i++) {
    sheet.getColumn(i).width = uniformWidth;
}



}); 



    /* DOWNLOAD FILE */
    const buffer = await workbook.xlsx.writeBuffer();
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
clearBtn.addEventListener("click", () => {
    filesData = [];
    fileInput.value = "";
    excelTitleInput.value = "";
});
