
function main(workbook: ExcelScript.Workbook, employeeID: number, issueType: string, testMneumonic: string, accession: string, eitNeeded: boolean, resolvedBy: number, relatedAccessions ?: string[]) {

    //This just gets the active worksheet
    let sheet = workbook.getActiveWorksheet();

    //Formats the employee information in a way that the sheet uses. 
    let currentDate: Date = new Date();
    let empAdded: string = employeeID.toString() + ", " + currentDate.toDateString();


    //This is for the related Accessions column
    //Declares a string to be added to the column later. 
    let relatedAccession: string = "";
    //Since the parameter can be a list of accessions, a single accession, or nothing, we need to check for each case.
    //If it is an array
    if (Array.isArray(relatedAccessions)) {
        // If the array is empty it will use the empty string.
        if (relatedAccessions.length < 1) {
        }
        //If the length is one it will simply convert to string.
        else if(relatedAccession.length === 1){
            relatedAccession = relatedAccessions[0];
        }
        //Otherwise it will iterate throught the list and create a string to be added to the column.
        else {
            relatedAccessions.forEach(accession => {
                relatedAccession += (accession + ", ")
            });
        }
        //If the parameter is not an array it will simply add the empty string to the column.
    }


    // Data to append, it can either be string, number or bool.
    let newData: (string | number | boolean)[] = [empAdded, issueType, testMneumonic, accession, relatedAccession, "" ,eitNeeded, resolvedBy];

    // Columns to check for duplicates
    const columnCIndex = 2; // Column C is index 2
    const columnDIndex = 3; // Column D is index 3

    // Get the used range of the sheet
    let range = sheet.getUsedRange();

    if (range) {
        let values = range.getValues();

        // Check for duplicates in columns C and D
        let duplicateFound = values.some(row =>
            row[columnCIndex] === newData[columnCIndex] && row[columnDIndex] === newData[columnDIndex]
        );

        if (duplicateFound) {
            console.log("The Accession is already on the pending log.");
            return;
        }

        // Append the new row at the end of the sheet
        let lastRow = range.getRowCount();
        let newRange = sheet.getRangeByIndexes(lastRow, 0, 1, newData.length);
        newRange.setValues([newData]);

        // Maintain formatting rules by copying from the previous row
        let previousRowRange = sheet.getRangeByIndexes(lastRow - 1, 0, 1, newData.length);
        newRange.copyFrom(previousRowRange, ExcelScript.RangeCopyType.formats);

        //If an EIT is needed it will make it the appropriate color.
    
        let fill = newRange.getColumn(5).getFormat().getFill();

        if(eitNeeded){
            fill.setColor('orange');
        }
        else{
            fill.setColor('green');
        }
        console.log("New row added successfully.");
    } else {
        console.log("The sheet is empty. Appending as the first row.");
        let newRange = sheet.getRangeByIndexes(0, 0, 1, newData.length);
        newRange.setValues([newData]);
    }
}
