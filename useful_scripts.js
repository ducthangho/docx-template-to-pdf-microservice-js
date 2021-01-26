/*
    Use this function to perform 'Update', 'Create', or 'Delete'
    async actions on batches of records that could potentially 
    more than 50 records.

    ::PARAMETERS::
    action = string; one of 3 values:
           - 'Update' to call table.updateRecordsAsync()
           - 'Create' to call table.createRecordsAsync()
           - 'Delete' to call table.deleteRecordsAsync()

    table = Table; the table the action will be performed in

    records = Array; the records to perform the action on
            - Ensure the record objects inside the array are
            formatted properly for the action you wish to
            perform

    ::RETURNS::
    recordsActedOn = integer, array of recordId's, or null; 
                   - Update Success: integer; the number of records processed by the function
                   - Delete Success: integer; the number of records processed by the function
                   - Create Success: array; the id strings of records created by the function
                   - Failure: null;
*/
async function batchAnd(action, table, records) {
    let recordsActedOn;

    switch (action) {
        case 'Update':
            recordsActedOn = records.length;
            while (records.length > 0) {
                await table.updateRecordsAsync(records.slice(0, 50));
                records = records.slice(50);
            };
            break;

        case 'Create':
            recordsActedOn = [];
            while (records.length > 0) {
                let recordIds = await table.createRecordsAsync(records.slice(0, 50));
                recordsActedOn.push(...recordIds)
                records = records.slice(50);
            };
            break;

        case 'Delete':
            recordsActedOn = records.length;
            while (records.length > 0) {
                await table.deleteRecordsAsync(records.slice(0, 50));
                records = records.slice(50);
            }
            break;

        default:
            output.markdown(`**Please use either 'Update', 'Create', or 'Delete' as the "action" parameter for the "batchAnd()" function.**`);
            recordsActedOn = null;
    }
    return recordsActedOn;
}

// use if both fields use the same basic data type (string based, number based, dateTime, etc.)
// note that fields can be different field types as long as the data type is the same
async function copySameFieldTypes(table, record, sourceFieldName, targetFieldName) {
    // the read and write formats are the same
    const writeValue = record.getCellValue(sourceFieldName);
    await table.updateRecordAsync(record, {
        [targetFieldName]: writeValue
    });
    return writeValue;
}


// use if the target field is a string (singleLineText, multilineText, richText, email, phone, url, etc)
async function copyToStringField(table, record, sourceFieldName, targetFieldName) {
    // Airtable already provides the cell value as a string
    const writeValue = record.getCellValueAsString(sourceFieldName);
    await table.updateRecordAsync(record, {
        [targetFieldName]: writeValue
    });
    return writeValue;
}


// use if the source is a string, but the target is a number
async function copyStringToNumberField(table, record, sourceFieldName, targetFieldName) {
    let writeValue = record.getCellValue(sourceFieldName);
    if ((writeValue === null) || (Number(writeValue) === NaN)) {
        // if the string is not a number, make the target field blank
        // this check is necessary to avoid having a null source field end up as zero in the target
        writeValue = null;
    } else {
        // convert the string to a number
        // Note: the string must be an UNFORMATTED number. Values like "1,000" will not work
        writeValue = Number(writeValue);
    }
    await table.updateRecordAsync(record, {
        [targetFieldName]: writeValue
    });
    return writeValue;
}


// use if the source is a string, but the target is a date
async function copyStringToDateTimeField(table, record, sourceFieldName, targetFieldName) {
    let writeValue = record.getCellValue(sourceFieldName);
    if (writeValue) {
        let dateTime = new Date(writeValue);
        writeValue = dateTime.toISOString();
    }
    await table.updateRecordAsync(record, {
        [targetFieldName]: writeValue
    });
    return writeValue;
}


// use if the target is a singleSelect
async function copyToSingleSelect(table, record, sourceFieldName, targetFieldName) {
    let writeValue = record.getCellValueAsString(sourceFieldName);
    if (writeValue) {
        writeValue = {
            "name": writeValue
        };
    } else {
        writeValue = null;
    }
    await table.updateRecordAsync(record, {
        [targetFieldName]: writeValue
    });
    return writeValue;
}


// use if the target is a multipeSelects
async function copyToMultipleSelects(table, record, sourceFieldName, targetFieldName) {
    let writeValue = record.getCellValueAsString(sourceFieldName);
    if (writeValue) {
        writeValue = writeValue.split(", ");
        writeValue = writeValue.map(value => {
            return {
                "name": value
            }
        });
    } else {
        writeValue = null;
    }
    await table.updateRecordAsync(record, {
        [targetFieldName]: writeValue
    });
    return writeValue;
}