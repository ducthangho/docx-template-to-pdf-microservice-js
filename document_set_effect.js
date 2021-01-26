const LOOKUP_TABLE = "HĐTV"
const DESTINATION_TABLE = "Hồ sơ nhân sự"
const HR_FIELD = "Nhân sự"
const ID_FIELD = "Mã NV"
const EXPIRY_DATE_FIELD = "Ngày kết thúc thử việc"
const check_expiry_date = true
const muc_luong = "Mức lương"
const luong_bhxh = "Lương BHXH"
const he_so_tham_nien = "Hệ số thâm niên"
const phu_cap_tham_nien = "Phụ cấp thâm niên"
const phu_cap_chuyen_can = "Phụ cấp chuyên cần"
const phu_cap_trach_nhiem = "Phụ cấp trách nhiệm"
const tien_nang_suat = "Tiền năng suất"
const tien_dien_thoai = "Tiền điện thoại"
const tien_an = "Tiền ăn"
const tien_nha_o = "Tiền nhà ở"
const tien_xang_xe = "Tiền xăng xe"
const bo_phan = "Bộ phận"
const bac = "Bậc"
const chuc_vu = "Chức vụ/chức danh"
const contract_state = "Tình trạng Hợp đồng lao động"

let destinationTbl = base.getTable(DESTINATION_TABLE);
let lookupTbl = base.getTable(LOOKUP_TABLE);

const DEBUG = true;


// Prompt the user to pick a record 
// If this script is run from a button field, this will use the button's record instead.
// let tblRecords = await table.selectRecordsAsync();
let destRecords = await destinationTbl.selectRecordsAsync();
let record = await input.recordAsync('Select a record to use', lookupTbl);


await main()

async function main() {
    if (check_expiry_date) {
        let expiry_date = record.getCellValueAsString(EXPIRY_DATE_FIELD);
        let fieldType = getFieldType(lookupTbl, record, EXPIRY_DATE_FIELD);
        let field = lookupTbl.getField(EXPIRY_DATE_FIELD);
        let tp = fieldType.type;
        let isLookup = fieldType.isLookup;
        if (tp == "date" && expiry_date) {
            console.log("Found expiry date");
            let date_format = (isLookup) ? field.options.result.options.dateFormat.format : field.options.dateFormat.format;
            // if (isLookup) console.log("Look up " + val + "     " + date_format);
            let end_date = parseString(expiry_date, date_format); //Date type     
            let currentDate = new Date();
            if (currentDate > end_date) {
                console.log("HĐTV has expired since " + end_date.toLocaleDateString("vi-VN"));
                return;
            }
        }
    }


    let linkedRecID = record.getCellValueAsString(HR_FIELD);
    let bo_phan_val = record.getCellValueAsString(bo_phan);
    let chuc_vu_val = record.getCellValueAsString(chuc_vu);
    let bac_val = record.getCellValueAsString(bac);
    let muc_luong_val = record.getCellValue(muc_luong);
    let luong_bhxh_val = record.getCellValue(luong_bhxh);
    let he_so_tham_nien_val = record.getCellValue(he_so_tham_nien);
    let phu_cap_tham_nien_val = record.getCellValue(phu_cap_tham_nien);
    let phu_cap_chuyen_can_val = record.getCellValue(phu_cap_chuyen_can);
    let phu_cap_trach_nhiem_val = record.getCellValue(phu_cap_trach_nhiem);
    let tien_nang_suat_val = record.getCellValue(tien_nang_suat);
    let tien_dien_thoai_val = record.getCellValue(tien_dien_thoai);
    let tien_an_val = record.getCellValue(tien_an);
    let tien_nha_o_val = record.getCellValue(tien_nha_o);
    let tien_xang_xe_val = record.getCellValue(tien_xang_xe);
    let contract_state_val = "Thử việc"


    console.log("Selected:   " + linkedRecID);
    let matchingItem = destRecords.records.filter(record => record.getCellValueAsString(ID_FIELD) == linkedRecID)

    if (matchingItem.length > 0) {
        console.log("Found matching item " + JSON.stringify(matchingItem));
        let item = destRecords.getRecord(matchingItem[0].id);

        // console.log(muc_luong_val+"    "+luong_bhxh_val+"    "+phu_cap_tham_nien_val+"      "+phu_cap_trach_nhiem_val);
        // console.log(tien_nang_suat_val+"    "+tien_dien_thoai_val+"    "+tien_an_val+"      "+tien_nha_o_val+"      "+tien_xang_xe_val );

        let updated_record = {}
        updated_record[bo_phan] = checkAndSetVal(destinationTbl, bo_phan, bo_phan_val); //Bo_phan is single select type
        updated_record[chuc_vu] = checkAndSetVal(destinationTbl, chuc_vu, chuc_vu_val);
        updated_record[bac] = checkAndSetVal(destinationTbl, bac, bac_val); //bac is single select type
        updated_record[muc_luong] = muc_luong_val;
        updated_record[luong_bhxh] = luong_bhxh_val;
        updated_record[he_so_tham_nien] = he_so_tham_nien_val;
        updated_record[phu_cap_tham_nien] = phu_cap_tham_nien_val;
        updated_record[phu_cap_chuyen_can] = phu_cap_chuyen_can_val;
        updated_record[phu_cap_trach_nhiem] = phu_cap_trach_nhiem_val;
        updated_record[tien_nang_suat] = tien_nang_suat_val;
        updated_record[tien_dien_thoai] = tien_dien_thoai_val;
        updated_record[tien_an] = tien_an_val;
        updated_record[tien_nha_o] = tien_nha_o_val;
        updated_record[tien_xang_xe] = tien_xang_xe_val;
        updated_record[contract_state] = contract_state_val;

        await destinationTbl.updateRecordAsync(item, updated_record);

    } else console.log("No matching items");
}

function getFieldType(tbl, record, fieldName) {
    let field = tbl.getField(fieldName);
    let tp = field.type;

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


function checkAndSetVal(tbl, fieldName, val) {
    let field = tbl.getField(fieldName);
    let tp = field.type;
    console.log(field);
    console.log(tp);
    if (tp == "singleSelect") {
        return {
            name: val
        }
    }
    return val;
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