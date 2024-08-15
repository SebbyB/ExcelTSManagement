


function main(workbook: ExcelScript.Workbook, starting: number, ending: number) {
    // Define the configuration within the script
    let config = {
        startColumn: starting, // Index 2 corresponds to column C
        endColumn: ending    // Index 4 corresponds to column E (this is inclusive)
    };

    // Create an object to hold data from all sheets
    let result: { [key: string]: string[][] } = {};

    // Get all worksheets in the workbook
    let sheets: ExcelScript.Worksheet[] = workbook.getWorksheets();

    // Loop through each sheet
    sheets.forEach((sheet) => {
        // Get the used range of the sheet
        let range: ExcelScript.Range = sheet.getUsedRange();

        if (range) {
            // Get the text values in the used range
            let values: string[][] = range.getTexts();

            // Prepare the data as a 2D array (array of arrays)
            let sheetData: string[][] = [];
            values.forEach((row) => {
                // Extract the desired range of columns from each row
                let selectedColumns = row.slice(config.startColumn, config.endColumn + 1);
                sheetData.push(selectedColumns);
            });

            // Add the formatted data to the result object
            result[sheet.getName()] = sheetData;
        }
    });

    // Convert the result object to a JSON string
    let jsonResult: string = JSON.stringify(result);

    // Log the JSON string to the console (you could also return it, save it, etc.)
    console.log(jsonResult);

    // Return the JSON string if you need to use it elsewhere
    return jsonResult;
}
