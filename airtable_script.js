// const serverURL = "http://localhost:8000/docx/docx-pdf";
const serverURL = "https://converter-3n6nh2nscq-as.a.run.app/docx/docx-pdf";

const CONTRACT_TBL = "HĐLĐ"
const HR_FIELD = "Nhân sự"
const TEMPLATE_FIELD = "Mẫu HĐ"
let table = base.getTable("Hồ sơ nhân sự");
let contractTbl = base.getTable(CONTRACT_TBL);
let templateTbl = base.getTable("Biểu mẫu");
let conversionFieldTbl = base.getTable("Mã chuyển đổi");
const DEBUG = true;
const exceptionFields = {
    "Ngày ký HĐLĐ chính thức bằng chữ": -3,
    "Ngày ký phụ lục hợp đồng bằng chữ": -3,
    "Ngày thử việc bằng chữ": -3,
    "Ngày nghỉ việc bằng chữ": -3

}

const mapField2Tbl = {}


// Prompt the user to pick a record 
// If this script is run from a button field, this will use the button's record instead.
let tblRecords = await table.selectRecordsAsync();
let templateRecords = await templateTbl.selectRecordsAsync();
let fieldRecords = await conversionFieldTbl.selectRecordsAsync();
let record = await input.recordAsync('Select a record to use', contractTbl);


await main()


async function main() {
    // pseudo code for guard clause
    if (record) {
        // Customize this section to handle the selected record
        // You can use record.getCellValue("Field name") to access
        // cell values from the record
        if (DEBUG) output.text(`You selected this record: ${record.name}`);

        /****************Now get an employee if not selected ******************/
        let employeeID = record.getCellValue(HR_FIELD);
        if (!employeeID) {
            console.log("Chưa chọn nhân sự");
            return;
        }
        let employee = tblRecords.getRecord(employeeID[0].id);
        if (DEBUG) output.text("Template is " + JSON.stringify(employee));

        /****************Now get a template if not already selected ******************/
        let templateID = record.getCellValue(TEMPLATE_FIELD);
        if (!templateID) {
            console.log("Chưa chọn biểu mẫu");
            return;
        }
        let template = templateRecords.getRecord(templateID[0].id);
        if (DEBUG) output.text("Template is " + JSON.stringify(template));

        //Now get all the parameters    
        let parameters = template.getCellValue("Parameters");

        var data = {};

        // if (DEBUG) output.table(fieldRecords);
        // let contract = contractRecords.getRecord(record.id);
        let contract = record;
        table.fields.forEach(item => {
            mapField2Tbl[item.name] = {
                obj: employee,
                tbl: table,
                type: item.type
            }
        });
        contractTbl.fields.forEach(item => {
            mapField2Tbl[item.name] = {
                obj: contract,
                tbl: contractTbl,
                type: item.type
            };
        });

        // console.log(JSON.stringify(mapField2Tbl));

        for (let obj in parameters) {
            let name = parameters[obj].name;
            let val = getValue(name); //  employee.getCellValueAsString( name);        
            let record = await findRecord(fieldRecords.records, "Name", name);
            if (!record) continue;
            let key = record.getCellValueAsString("docx");
            data[key] = val;
        }

        let s = await findRecord(fieldRecords.records, "Name", "cmdDelimiter_s");
        let e = await findRecord(fieldRecords.records, "Name", "cmdDelimiter_e");

        let Folder = "Documents";

        let out = {
            data: data,
            cmdDelimiter: [s.getCellValueAsString("docx"), e.getCellValueAsString("docx")],
            folder: Folder,
        }

        let filenames = template.getCellValueAsString("Filenames").split("\n");
        let alternativeURLs = template.getCellValueAsString("Alternative URL").split("\n");
        let templateFolder = template.getCellValueAsString("Folder");

        let templateFiles = template.getCellValue("Template file");
        for (let obj in templateFiles) {
            let url = templateFiles[obj].url;
            out.file = url;
            out.filename = templateFolder + "/" + filenames[obj];
            out.alternativeURL = alternativeURLs[obj];
            out.outputFilename = CONTRACT_TBL + "_" + employee.getCellValueAsString("Họ và tên") + "_" + employee.getCellValueAsString("Mã NV");
            let payload = JSON.stringify(out);
            // if (DEBUG) console.log("Sending "+payload);

            let response = await fetch(serverURL, {
                method: 'POST',
                body: payload,
                headers: {
                    'Content-Type': 'application/json',
                },
            });


            let ret = await response.json();
            if (ret.status == "ok") {
                if (DEBUG) output.text(JSON.stringify(ret.attachments));
                let attachments = ret.attachments;
                let docx = [];
                let pdf = [];

                if (attachments) {
                    attachments.forEach(attachment => {
                        let ext = attachment.filename.toString().split('.').pop().trim();

                        if (ext == "docx") {
                            docx.push(attachment)
                        } else {
                            pdf.push(attachment)
                        }

                    });

                    await contractTbl.updateRecordAsync(
                        record, {
                            "docx": docx,
                            "pdf": pdf,
                        }
                    )
                }
            } else console.log(JSON.stringify(ret.errors))

        }
    } else {
        if (DEBUG) output.text('No record was selected');
    }
}



function numberWithCommas(x) {
    return x.toString().replace(/\B(?<!\.\d*)(?=(\d{3})+(?!\d))/g, ",");
}

function addWorkDays(startDate, days) {
    if(isNaN(days)) {
        console.log("Value provided for \"days\" was not a number");
        return
    }
    if (!(startDate instanceof Date)) {
        console.log("Value provided for \"startDate\" was not a Date object");
        return
    }
    // Get the day of the week as a number (0 = Sunday, 1 = Monday, .... 6 = Saturday)
    var dow = startDate.getDay();
    var daysToAdd = parseInt(days);
    // If the current day is Sunday add one day
    if (dow == 0)
        daysToAdd++;
    // If the start date plus the additional days falls on or after the closest Saturday calculate weekends
    if (dow + daysToAdd >= 6) {
        //Subtract days in current working week from work days
        var remainingWorkDays = daysToAdd - (5 - dow);
        //Add current working week's weekend
        daysToAdd += 2;
        if (remainingWorkDays > 5) {
            //Add two days for each working week by calculating how many weeks are included
            daysToAdd += 2 * Math.floor(remainingWorkDays / 5);
            //Exclude final weekend if remainingWorkDays resolves to an exact number of weeks
            if (remainingWorkDays % 5 == 0)
                daysToAdd -= 2;
        }
    }
    startDate.setDate(startDate.getDate() + daysToAdd);
    return startDate;
}


function dateToStr(key, isLookup = false) {
    let mapVal = mapField2Tbl[key];
    let record = mapVal.obj;
    let tbl = mapVal.tbl;
    let val = record.getCellValueAsString(key).trim();
    let field = tbl.getField(key);
    let date_format = (isLookup) ? field.options.result.options.dateFormat.format : field.options.dateFormat.format;
    let signed_date = parseString(val, date_format); //Date type
    let specialFieldVal = exceptionFields[key + " bằng chữ"];
    if (specialFieldVal !== undefined) {
        signed_date = addWorkDays(signed_date, specialFieldVal);
    }


    const options = {
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    };
    if (signed_date)
        val = "ngày " + signed_date.toLocaleDateString("vi-VN", options).replace(",", " năm"); //Convert to Vietnamese style
    else val = "ngày    tháng    năm "; //Convert to Vietnamese style                   
    console.log(val);
    return val;
}

function getFieldType(fieldName) {
    let mapVal = mapField2Tbl[fieldName];
    let record = mapVal.obj;
    let tbl = mapVal.tbl;
    let tp = mapVal.type;
    let field = tbl.getField(fieldName);

    if (tp == "multipleLookupValues") {
        let obj = record.getCellValue(fieldName);
        tp = field.options.result.type;
        return {
            type: tp,
            isLookup: true
        };
    }
    return {
        type: tp,
        isLookup: false
    };
}

function removeEnclosingQuotes(str) {
    str = str.replace(/^"(.*)"$/, "$1");
    return str;
}

function getValue(key) {
    if (!mapField2Tbl[key]) {
        if (key.endsWith("bằng chữ")) {
            let key2 = key.substring(0, key.length - 9);
            let fieldType = getFieldType(key2);
            let tp = fieldType.type;
            let isLookup = fieldType.isLookup;
            if (tp == 'date') {
                // let tmp = dateToStr(key2, isLookup);
                // console.log("Testing "+tmp);
                return dateToStr(key2, isLookup);
            }

        }
        console.log(key);
        return ""
    }

    let mapVal = mapField2Tbl[key];
    let record = mapVal.obj;
    let tbl = mapVal.tbl;
    let val = record.getCellValueAsString(key).trim();
    let field = tbl.getField(key);
    if (field.type == "multipleRecordLinks") {
        val = removeEnclosingQuotes(val);
    }

    let fieldType = getFieldType(key);
    let tp = fieldType.type;
    let isLookup = fieldType.isLookup;

    if (val == "null" || val == "0") val = ""

    if (tp == 'number') {
        if (key != "STK BANK")
            val = numberWithCommas(val);
        if (val == "") val = "                ";
    } else if (tp == 'date') {
        let date_format = (isLookup) ? field.options.result.options.dateFormat.format : field.options.dateFormat.format;
        // if (isLookup) console.log("Look up " + val + "     " + date_format);
        let signed_date = parseString(val, date_format); //Date type            
        if (signed_date)
            val = signed_date.toLocaleDateString("vi-VN"); //Convert to Vietnamese style
        else val = "   /  /    "
        // if (isLookup) console.log("Result look up " + val);
    }

    return val;
}

async function findRecord(records, fieldName, searchValue) {
    for (let idx in records) {
        let record = records[idx];
        if (record && record.getCellValueAsString(fieldName) == searchValue)
            return record;
    }
    return null;
}


// Utility function to append a 0 to single-digit numbers
function LZ(x) {
    return (x < 0 || x > 9 ? "" : "0") + x
};
// Full month names. Change this for local month names
const monthNames = new Array('January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December');
// Month abbreviations. Change this for local month names
const monthAbbreviations = new Array('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec');
// Full day names. Change this for local month names
const dayNames = new Array('Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday');
// Day abbreviations. Change this for local month names
const dayAbbreviations = new Array('Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat');
// Used for parsing ambiguous dates like 1/2/2000 - default to preferring 'American' format meaning Jan 2.
// Set to false to prefer 'European' format meaning Feb 1
const preferAmericanFormat = true;

// Parse a string and convert it to a Date object.
// If no format is passed, try a list of common formats.
// If string cannot be parsed, return null.
// Avoids regular expressions to be more portable.
function parseString(val, format) {
    // If no format is specified, try a few common formats
    // If no format is specified, try a few common formats
    if (typeof(format) == "undefined" || format == null || format == "") {
        var generalFormats = new Array('y-M-d', 'MMM d, y', 'MMM d,y', 'y-MMM-d', 'd-MMM-y', 'MMM d', 'MMM-d', 'd-MMM');
        var monthFirst = new Array('M/d/y', 'M-d-y', 'M.d.y', 'M/d', 'M-d');
        var dateFirst = new Array('d/M/y', 'd-M-y', 'd.M.y', 'd/M', 'd-M');
        var checkList = new Array(generalFormats, preferAmericanFormat ? monthFirst : dateFirst, preferAmericanFormat ? dateFirst : monthFirst);
        for (var i = 0; i < checkList.length; i++) {
            var l = checkList[i];
            for (var j = 0; j < l.length; j++) {
                var d = parseString(val, l[j]);
                if (d != null) {
                    return d;
                }
            }
        }
        return null;
    };


    function isInteger(val) {
        for (var i = 0; i < val.length; i++) {
            if ("1234567890".indexOf(val.charAt(i)) == -1) {
                return false;
            }
        }
        return true;
    };

    function getInt(str, i, minlength, maxlength) {
        for (var x = maxlength; x >= minlength; x--) {
            var token = str.substring(i, i + x);
            if (token.length < minlength) {
                return null;
            }
            if (isInteger(token)) {
                return token;
            }
        }
        return null;
    };
    val = val + "";
    format = format + "";
    var i_val = 0;
    var i_format = 0;
    var c = "";
    var token = "";
    var token2 = "";
    var x, y;
    var year = new Date().getFullYear();
    var month = 1;
    var date = 1;
    var hh = 0;
    var mm = 0;
    var ss = 0;
    var ampm = "";
    // console.log("Parsing "+val+"  "+format );
    while (i_format < format.length) {
        // Get next token from format string
        c = format.charAt(i_format);
        token = "";
        while ((format.charAt(i_format) == c) && (i_format < format.length)) {
            token += format.charAt(i_format++);
        }
        // Extract contents of value based on format token
        if (token == "yyyy" || token == "yy" || token == "y" || token == "YYYY" || token == "YY" || token == "Y") {
            if (token == "yyyy" || token == "YYYY") {
                x = 4;
                y = 4;
            }
            if (token == "yy" || token == "YY") {
                x = 2;
                y = 2;
            }
            if (token == "y" || token == "Y") {
                x = 2;
                y = 4;
            }
            year = getInt(val, i_val, x, y);
            // console.log("Year "+year);
            if (year == null) {
                return null;
            }
            i_val += year.length;
            if (year.length == 2) {
                if (year > 70) {
                    year = 1900 + (year - 0);
                } else {
                    year = 2000 + (year - 0);
                }
            }
        } else if (token == "MMM" || token == "NNN") {
            month = 0;
            var names = (token == "MMM" ? (monthNames.concat(monthAbbreviations)) : monthAbbreviations);
            for (var i = 0; i < names.length; i++) {
                var month_name = names[i];
                if (val.substring(i_val, i_val + month_name.length).toLowerCase() == month_name.toLowerCase()) {
                    month = (i % 12) + 1;
                    i_val += month_name.length;
                    break;
                }
            }
            if ((month < 1) || (month > 12)) {
                return null;
            }
        } else if (token == "EE" || token == "E") {
            var names = (token == "EE" ? dayNames : dayAbbreviations);
            for (var i = 0; i < names.length; i++) {
                var day_name = names[i];
                if (val.substring(i_val, i_val + day_name.length).toLowerCase() == day_name.toLowerCase()) {
                    i_val += day_name.length;
                    break;
                }
            }
        } else if (token == "MM" || token == "M") {
            month = getInt(val, i_val, token.length, 2);
            // console.log("Month "+month);
            if (month == null || (month < 1) || (month > 12)) {
                return null;
            }
            i_val += month.length;
        } else if (token == "dd" || token == "d" || token == "DD" || token == "D") {
            date = getInt(val, i_val, token.length, 2);
            // console.log("Day "+date);
            if (date == null || (date < 1) || (date > 31)) {
                return null;
            }
            i_val += date.length;
        } else if (token == "hh" || token == "h") {
            hh = getInt(val, i_val, token.length, 2);
            if (hh == null || (hh < 1) || (hh > 12)) {
                return null;
            }
            i_val += hh.length;
        } else if (token == "HH" || token == "H") {
            hh = getInt(val, i_val, token.length, 2);
            if (hh == null || (hh < 0) || (hh > 23)) {
                return null;
            }
            i_val += hh.length;
        } else if (token == "KK" || token == "K") {
            hh = getInt(val, i_val, token.length, 2);
            if (hh == null || (hh < 0) || (hh > 11)) {
                return null;
            }
            i_val += hh.length;
            hh++;
        } else if (token == "kk" || token == "k") {
            hh = getInt(val, i_val, token.length, 2);
            if (hh == null || (hh < 1) || (hh > 24)) {
                return null;
            }
            i_val += hh.length;
            hh--;
        } else if (token == "mm" || token == "m") {
            mm = getInt(val, i_val, token.length, 2);
            if (mm == null || (mm < 0) || (mm > 59)) {
                return null;
            }
            i_val += mm.length;
        } else if (token == "ss" || token == "s") {
            ss = getInt(val, i_val, token.length, 2);
            if (ss == null || (ss < 0) || (ss > 59)) {
                return null;
            }
            i_val += ss.length;
        } else if (token == "a") {
            if (val.substring(i_val, i_val + 2).toLowerCase() == "am") {
                ampm = "AM";
            } else if (val.substring(i_val, i_val + 2).toLowerCase() == "pm") {
                ampm = "PM";
            } else {
                return null;
            }
            i_val += 2;
        } else {
            if (val.substring(i_val, i_val + token.length) != token) {
                return null;
            } else {
                i_val += token.length;
            }
        }
    }
    // If there are any trailing characters left in the value, it doesn't match
    if (i_val != val.length) {
        // console.log("Not valid");
        return null;
    }
    // Is date valid for month?
    if (month == 2) {
        // Check for leap year
        if (((year % 4 == 0) && (year % 100 != 0)) || (year % 400 == 0)) { // leap year
            if (date > 29) {
                return null;
            }
        } else {
            if (date > 28) {
                return null;
            }
        }
    }
    if ((month == 4) || (month == 6) || (month == 9) || (month == 11)) {
        if (date > 30) {
            return null;
        }
    }
    // Correct hours value
    if (hh < 12 && ampm == "PM") {
        hh = hh - 0 + 12;
    } else if (hh > 11 && ampm == "AM") {
        hh -= 12;
    }
    // console.log("Now");
    return new Date(year, month - 1, date, hh, mm, ss);
};

// Format a date into a string using a given format string
/*function formatDate(date, format) {
    format = format + "";
    var result = "";
    var i_format = 0;
    var c = "";
    var token = "";
    var y = date.getYear() + "";
    var M = date.getMonth() + 1;
    var d = date.getDate();
    var E = date.getDay();
    var H = date.getHours();
    var m = date.getMinutes();
    var s = date.getSeconds();
    var yyyy, yy, MMM, MM, dd, hh, h, mm, ss, ampm, HH, H, KK, K, kk, k;
    // Convert real date parts into formatted versions
    var value = new Object();
    if (y.length < 4) {
        y = "" + (+y + 1900);
    }
    value["y"] = "" + y;
    value["yyyy"] = y;
    value["yy"] = y.substring(2, 4);
    value["M"] = M;
    value["MM"] = LZ(M);
    value["MMM"] = monthNames[M - 1];
    value["NNN"] = monthAbbreviations[M - 1];
    value["d"] = d;
    value["dd"] = LZ(d);
    value["E"] = dayAbbreviations[E];
    value["EE"] = dayNames[E];
    value["H"] = H;
    value["HH"] = LZ(H);
    if (H == 0) {
        value["h"] = 12;
    } else if (H > 12) {
        value["h"] = H - 12;
    } else {
        value["h"] = H;
    }
    value["hh"] = LZ(value["h"]);
    value["K"] = value["h"] - 1;
    value["k"] = value["H"] + 1;
    value["KK"] = LZ(value["K"]);
    value["kk"] = LZ(value["k"]);
    if (H > 11) {
        value["a"] = "PM";
    } else {
        value["a"] = "AM";
    }
    value["m"] = m;
    value["mm"] = LZ(m);
    value["s"] = s;
    value["ss"] = LZ(s);
    while (i_format < format.length) {
        c = format.charAt(i_format);
        token = "";
        while ((format.charAt(i_format) == c) && (i_format < format.length)) {
            token += format.charAt(i_format++);
        }
        if (typeof(value[token]) != "undefined") {
            result = result + value[token];
        } else {
            result = result + token;
        }
    }
    return result;
};//*/