/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("container").innerHTML = `
            <h1>Welcome to My Office Add-in</h1>
            <button id="insertData">Insert Sample Data</button>
        `;

        document.getElementById("insertData").addEventListener("click", insertSampleData);
    }
});

async function insertSampleData() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.values = [
                ["Product", "Quantity", "Price"],
                ["Widget A", 10, 19.99],
                ["Widget B", 15, 29.99],
                ["Widget C", 20, 39.99]
            ];
            range.format.autofitColumns();
            await context.sync();
        });
    } catch (error) {
        console.error(error);
    }
} 