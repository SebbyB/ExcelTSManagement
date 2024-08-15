
enum issueType{
    CPRPULL,
    TRANSMITTED_ORDER_NOT_PROCESSED,
    TRACKING_TEST_ORDERED_INCORRECTLY,
    MISLABEL_BY_TEST,
    MISLABEL_BY_TYPE,
}
function main(workbook: ExcelScript.Workbook, employeeID : number, issueType: issueType, testMneumonic:string, accession : string, relatedAccessions : string[], eitNeeded : boolean, resolvedBy : number) {
    
 //This just gets the active worksheet
    let sheet = workbook.getActiveWorksheet();

//Formats the employee information in a way that the sheet uses. 
    let currentDate : Date = new Date();
    let empAdded : string = employeeID.toString() + ", "+ currentDate.toDateString();


//This is for the related Accessions column
    //Declares a string to be added to the column later. 
    let relatedAccession : string= "";
    //Since the parameter can be a list of accessions, a single accession, or nothing, we need to check for each case.
    //If it is an array
    if(Array.isArray(relatedAccessions)){
        // If the array is empty it will use the empty string.
        if(relatedAccessions.length < 1){
        }
        //Otherwise it will iterate throught the list and create a string to be added to the column.
        else{
            relatedAccessions.forEach(accession => {
                relatedAccession += (accession+", " ) 
            });
        }
    //If the parameter is not an array it will simply add the empty string to the column.
    }
    else{
        relatedAccession = relatedAccessions;
    }


    // Data to append, it can either be string, number or bool.
    let newData: (string | number | boolean)[] = [empAdded, issueType,testMneumonic ,accession, relatedAccession, eitNeeded, "",resolvedBy];

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

        console.log("New row added successfully.");
    } else {
        console.log("The sheet is empty. Appending as the first row.");
        let newRange = sheet.getRangeByIndexes(0, 0, 1, newData.length);
        newRange.setValues([newData]);
    }
}
