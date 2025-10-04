export async function checkColumnCInFolios() {
    await Excel.run(async (context) => {
        // Get the worksheet by name
        const sheet = context.workbook.worksheets.getItem("FoliosFilteredByDoCalc");
        // Detect last used row
        const h1 = sheet.getRange("H1");
        h1.load("values");
        await context.sync();
        const lastRow = parseInt(h1.values[0][0]);
        await context.sync();

        // Load columns B and D from Folios ----
        const range = sheet.getRange(`B1:D${lastRow}`);
        range.load("values");
        await context.sync();
        console.log(range.values);
        let matchingValues = [];
        for (let i = 1; i < lastRow; i++) {
            let colB = range.values[i][0]; // column B
            let colC = range.values[i][1]; // column C
            let colD = range.values[i][2]; // column D

            console.log(`Row ${i + 1} → Column B value: ${colB}, Column C value: ${colC},,Column D value: ${colD}`);
            matchingValues.push(colB);

        }
        console.log("All matching values from column B (Folios):", matchingValues);
    });
}




