const express = require("express");
const bodyParser = require("body-parser");
const {createReport} = require("docx-templates");
const toPdf = require("office-to-pdf");
const got = require("got");
const fs = require("fs");
const stream = require("stream");
const app = express();
const promiseRetry = require("promise-retry");
const mobx = require("mobx");
const mobx_utils = require("mobx-utils");
const VietnameseTextNormalizer = require("VietnameseTextNormalizer/lib/binding.js");
const {default:JsExcelTemplate} = require("js-excel-template/nodejs/nodejs");
const {
    default: PQueue
} = require("p-queue");

const {
    getStore,
    document_category_type,
    ccns_type,
    relationship_type,
    country_list_type,
    ethnic_group_type,
    religons_type,
    marital_status_type,
    sampleRecord_type,
    sampleFamilyMember_type,
    Location,
    searchByWName,
    getWardFullname,
    getDistrictFullname,
    getProvinceFullname,
    strToLocation,
    getAddressLine1
} = require("./Store");

const rxjs = require("rxjs");
const {
    from,
    defer,
    Observable,
    debounceTime,
    debounce
} = rxjs;

const {
    fromPromise
} = mobx_utils;

const {
    observable,
    computed,
    autorun,
    runInAction
} = mobx;

const LOCATION_TBL = "Địa điểm";

class JsExcelTemplateEx extends JsExcelTemplate {
    static fromBufer(data) {
        const workbook = xlsx.read(data, {
            type: "buffer",
            cellNF: true,
            cellStyles: true,
            cellDates: true,
        });
        return new JsExcelTemplateEx(workbook);
    }

    toBuffer(bookType) {
      return xlsx.write(this.workbook, { bookType : bookType, bookSST: false, type: 'buffer' })
    }

    // saveAs(filepath: string) {
    //   XLSX.writeFile(this.workbook, filepath)
    // }
} //*/

const vn_normalizer = new VietnameseTextNormalizer();
const excelToJson = require("convert-excel-to-json");
const xlsx = require("xlsx");

const DEBUG = true;
const retries = 5;
const intervalBetweenRetries = 1000;
// Imports the Google Cloud client library
const {
    Storage
} = require("@google-cloud/storage");
const multer = require("multer");
// Creates a client
// const storage = new Storage();
// Creates a client from a Google service account key.
const storage = new Storage({
    keyFilename: "Converter-23608d7fef1a.json",
});
const bucket_name = "ducthangho_storage";
const bucket = storage.bucket(bucket_name);

const queue = new PQueue({
    concurrency: 1,
});

var multer_storage = multer.memoryStorage();

var store = observable.map({});

// var rx = new rxjs.Subject();

// POST parsing
app.use(
    bodyParser.json({
        limit: "50mb",
    })
);

// CORS support
app.use(function(req, res, next) {
    res.header("Access-Control-Allow-Origin", "*");
    res.header(
        "Access-Control-Allow-Headers",
        "Origin, X-Requested-With, Content-Type, Accept"
    );
    next();
});

function saveBufferToGSFile(res, path, contentType) {
    const file = bucket.file(path);

    if (DEBUG) console.log("Saving.............. " + path);
    const stream = file.createWriteStream({
        metadata: {
            contentType: contentType,
        },
    });
    stream.end(res);

    return new Promise((resolve, reject) => {
        stream.on("error", (err) => {
            if (DEBUG) console.log("Error saving file " + path + "  : " + err);
            reject(err);
        });
        stream.on("finish", (data) => {
            // if (DEBUG) console.log(path+" saved successfully.");
            resolve(data);
        });
    });
}

function fetchURL(url, data, contentType) {
    return got(url, {
        responseType: "buffer",
        resolveBodyOnly: true,
    }).then((res) => {
        return saveBufferToGSFile(res, data.filename, contentType).then(() =>
            createReport({
                template: res,
                output: "buffer",
                data: data.data || {},
                cmdDelimiter: data.cmdDelimiter,
            })
        );
    });
}

async function getTemplateContent(data){
    // console.log(JSON.stringify(data));
    var path = data.filename;
    let alternativeURL = data.alternativeURL;

    if (path != "") {
        const entry = store.get(path);
        if (!entry) {
            if (DEBUG) console.log("Fetching from Google Storage...");
            const file = bucket.file(path);
            var pr = file.download({
                // don't set destination here
            });
            runInAction(() => {
                console.log("Store " + path + " to store.");
                store.set(path, fromPromise(pr));
            });
        }
        return store
            .get(path)
            .then((contents) => {
                if (DEBUG)
                    console.log("Fetching " + path + " from Store ... Finished. ");
                const content = contents[0]; // contents is the file as Buffer
                return content;
            })
            .catch((error) => {
                if (DEBUG) console.log(error);
                let url = data.file;
                return got(url, {
                    responseType: "buffer",
                    resolveBodyOnly: true,
                });
            });
    }

    if (alternativeURL != "") {
        if (DEBUG) console.log("Receive alternativeURL " + alternativeURL);
        if (DEBUG) console.log("Fetching ...");
        return got(alternativeURL, {
            responseType: "buffer",
            resolveBodyOnly: true,
        });       
    }

    let url = data.file;
    if (DEBUG) console.log("Receive URL " + url);
    if (DEBUG) console.log("Fetching ...");

    return got(url, {
        responseType: "buffer",
        resolveBodyOnly: true,
    });
    // return fetchURL(url, data);
}

async function createReportPromise(data) {
    var path = data.filename;
    let alternativeURL = data.alternativeURL;

    if (path != "") {
        const entry = store.get(path);
        if (!entry) {
            if (DEBUG) console.log("Fetching from Google Storage...");
            const file = bucket.file(path);
            var pr = file.download({
                // don't set destination here
            });
            runInAction(() => {
                console.log("Store " + path + " to store.");
                store.set(path, fromPromise(pr));
            });
        }
        return store
            .get(path)
            .then((contents) => {
                if (DEBUG)
                    console.log("Fetching " + path + " from Store ... Finished. ");
                const content = contents[0]; // contents is the file as Buffer                
                return createReport({
                    template: content,
                    output: "buffer",
                    data: data.data || {},
                    cmdDelimiter: data.cmdDelimiter,
                });
            })
            .catch((error) => {
                if (DEBUG) console.log(error);
                let url = data.file;
                return fetchURL(url, data);
            });
    }

    if (alternativeURL != "") {
        if (DEBUG) console.log("Receive alternativeURL " + alternativeURL);
        if (DEBUG) console.log("Fetching ...");
        return fetchURL(alternativeURL, data);
    }

    let url = data.file;
    if (DEBUG) console.log("Receive URL " + url);
    if (DEBUG) console.log("Fetching ...");

    return fetchURL(url, data);
}

async function getAttachment(path) {
    try {
        const file = bucket.file(path);
        const signedUrls = await file.getSignedUrl({
            action: "read",
            expires: "03-09-2491",
        });

        // if (DEBUG) console.log(signedUrls[0]);
        return {
            url: signedUrls[0],
            filename: path.substring(path.lastIndexOf("/") + 1),
        };
    } catch (error) {
        console.log(error);
    }
}

app.post("/docx/docx", (req, res) => {
    // if (DEBUG) console.log(req.body.cmdDelimiter);
    let body = req.body;
    let promises = [];
    if (!Array.isArray(body)) promises.push(processDocx(body));
    else {
        body.forEach((data) => {
            promises.push(processDocx(data));
        });
    }

    return Promise.all(promises)
        .then((values) => {
            if (DEBUG) console.log("docx done!");
            res.json({
                status: "ok",
                attachments: values,
            });
        })
        .catch((err) => {
            if (DEBUG) console.log("Catch error " + err);
            res.json({
                status: "error",
                errors: ["" + err],
            });
        });
});

app.post("/docx/pdf", (req, res) => {
    let body = req.body;
    let promises = [];
    if (!Array.isArray(body)) {
        promises.push(processPdf(body));
    } else {
        body.forEach((data) => {
            promises.push(processPdf(data));
        });
    }

    return Promise.all(promises)
        .then((values) => {
            if (DEBUG) console.log("pdf done! ");
            res.json({
                status: "ok",
                attachments: values,
            });
        })
        .catch((err) => {
            if (DEBUG) console.log("Catch error " + err);
            res.json({
                status: "error",
                errors: ["" + err],
            });
        });
});

async function processDocx(data) {
    return createReportPromise(data).then((buffer) => {
        let folder = data.folder;
        let fn = data.outputFilename + ".docx";
        var path = folder + "/docx/" + fn;
        let contentType =
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        //Save file docx first
        return saveBufferToGSFile(buffer, path, contentType).then(() =>
            getAttachment(path)
        );
    });
}

async function processPdf(data) {
    return (
        createReportPromise(data)
        //get the docx buffer
        .then((buffer) => {
            let folder = data.folder;
            let fn = data.outputFilename + ".docx";
            var path = folder + "/docx/" + fn;
            let contentType =
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            //Save file docx first

            return new Promise((resolve, reject) => {
                queue.add(() => {
                    return (async(path) => {
                        try {
                            if (DEBUG) console.log("Saving " + path + " to pdf now.");
                            const pdfBuffer = await promiseRetry((retry, number) => {
                                if (DEBUG)
                                    console.log(
                                        "Converting " + path + " to pdf. Attemp no: " + number
                                    );
                                return toPdf(buffer).catch(retry);
                            });

                            let folder = data.folder;
                            let fn = data.outputFilename + ".pdf";
                            var path = folder + "/pdf/" + fn;
                            var docx = folder + "/docx/" + data.outputFilename + ".docx";
                            let contentType = "application/pdf";
                            // console.log("Save PDF to "+path+" ..... "+JSON.stringify(buffer));

                            await saveBufferToGSFile(pdfBuffer, path, contentType);
                            const pdfPromise = getAttachment(path);
                            resolve(pdfPromise);
                        } catch (e) {
                            reject(e);
                        }
                    })(path);
                });
            });
        })
    );
}

async function processPdfDocx(data) {
    return createReportPromise(data).then((buffer) => {
        let folder = data.folder;
        let fn = data.outputFilename + ".docx";
        var path = folder + "/docx/" + fn;
        let contentType =
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        //Save file docx first
        return saveBufferToGSFile(buffer, path, contentType).then(() => {
            return new Promise((resolve, reject) => {
                queue.add(() => {
                    return (async(path) => {
                        try {
                            const pdfBuffer = await promiseRetry((retry, number) => {
                                if (DEBUG)
                                    console.log(
                                        "Converting " + path + " to pdf. Attemp no: " + number
                                    );
                                return toPdf(buffer).catch(retry);
                            });
                            let folder = data.folder;
                            let fn = data.outputFilename + ".pdf";
                            var path = folder + "/pdf/" + fn;
                            var docx = folder + "/docx/" + data.outputFilename + ".docx";
                            let contentType = "application/pdf";
                            // console.log("Save PDF to "+path+" ..... "+JSON.stringify(buffer));

                            await saveBufferToGSFile(pdfBuffer, path, contentType);
                            console.log(path + " has been saved");
                            pdfPromise = getAttachment(path);
                            docxPromise = getAttachment(docx);
                            resolve(Promise.all([pdfPromise, docxPromise]));
                        } catch (e) {
                            reject(e);
                        }
                    })(path);
                });
            });
        });
    });
}


async function processJSON2excel(data) {
    // console.log("Received  " + JSON.stringify(data));

    let content = await getTemplateContent(data);
    let ws_bookType = data.bookType;
    // console.log("bookType  " + ws_bookType);

    let config = {
      source: content,
      type: "buffer"
    }

    const excelTemplate = JsExcelTemplateEx.fromBufer(
        config.source
    );
    excelTemplate.set("ns", data.records.ns);
    excelTemplate.set("mem", data.records.mem);
    // excelTemplate.saveAs("out.xls");
    return excelTemplate.toBuffer(ws_bookType);        
}

app.post("/json2excel", (req, res) => {
    let body = req.body;
    let promises = [];
    // console.log(req);

    if (!Array.isArray(body)) {
        promises.push(processJSON2excel(body));
    } else {
        body.forEach((data) => {
            promises.push(processJSON2excel(data));
        });
    }
    // console.log(JSON.stringify(req));

    return Promise.all(promises)
        .then((contents) => {
            if (DEBUG) console.log("json2xls done!");            
            

            contents.map(content => {
              const fileName = "out.xls";
              const fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
              res.set({              
                'Content-Type': fileType,
                'Content-Disposition': `attachment;inline;filename="${fileName}"`,
                'Content-Length': content.length
              })
            
              var readStream = new stream.PassThrough();
              readStream.end(content);            
              readStream.pipe(res);
            });
          
            // res.json({
            //     status: "ok",
            //     attachments: values,
            // });
        })
        .catch((err) => {
            console.log("Catch error " + err);
            res.json({
                status: "error",
                errors: ["" + err],
            });
            console.log("Return error");
        });
});

app.post("/docx/docx-pdf", (req, res) => {
    let body = req.body;
    let promises = [];

    if (!Array.isArray(body)) {
        promises.push(processPdfDocx(body));
    } else {
        body.forEach((data) => {
            promises.push(processPdfDocx(data));
        });
    }

    return Promise.all(promises)
        .then((values) => {
            if (DEBUG) console.log("pdf-docx done!");
            res.json({
                status: "ok",
                attachments: values,
            });
        })
        .catch((err) => {
            console.log("Catch error " + err);
            res.json({
                status: "error",
                errors: ["" + err],
            });
            console.log("Return error");
        });
});

function cpFromLocationObj(location, obj) {
  obj["Mã tỉnh thành"] = location.provinceCode;
  obj["Mã quận huyện"] = location.districtCode;
  obj["Mã phường xã"] = location.wardCode;
  obj["Tỉnh, thành phố"] = location.provinceName;
  obj["Quận, huyện"] = location.districtName;
  obj["Phường, xã"] = location.wardName;
  obj["Name"] = location.getFullName();
}

function fromLocationObj(location) {
  let obj = {};
  obj["Mã tỉnh thành"] = location.provinceCode;
  obj["Mã quận huyện"] = location.districtCode;
  obj["Mã phường xã"] = location.wardCode;
  obj["Tỉnh, thành phố"] = location.provinceName;
  obj["Quận, huyện"] = location.districtName;
  obj["Phường, xã"] = location.wardName;
  obj["Name"] = location.getFullName();
  for (let key in location) {
    if (obj[key] === undefined) obj[key] = location[key];
  }
  return obj;
}

app.post("/vnaddress", (req, res) => {
  try{

    if (!Location.provinceName2Code) 
        Location.init();
    let content = req.body;    
    // let str = vn_normalizer.Normalize(content.text);
    let str = vn_normalizer.Normalize(content.address);
    console.log(str);
    let location = strToLocation(str);
    if (!location){
      res.json({
        status: "error",
        errors: ["Định dạng hoặc địa chỉ không đúng."],
      });
    } else {
      let out = {}; 
      cpFromLocationObj(location,out);
      res.json({
        status: "ok",
        data: out
      });
    }
    
  } catch(err){
    res.json({
        status: "error",
        errors: ["" + err],
    });
  }
  
})

app.post("/templates", (req, res) => {
    if (DEBUG) console.log("Uploading new template");
    let content = req.body;

    let promises = [];
    content.forEach(
        (payload) => {
            let folder = payload.folder;
            let url = payload.url;
            let fn = payload.filename;
            var path = folder + "/" + fn;
            var contentType = payload.type;
            if (DEBUG) console.log(JSON.stringify(payload));
            // if (DEBUG) console.log("Receive URL " + url);

            let res = got(url, {
                    responseType: "buffer",
                    resolveBodyOnly: true,
                })
                .then((res) => saveBufferToGSFile(res, path, contentType))
                .then(() => {
                    const file = bucket.file(path);
                    if (DEBUG) console.log("Template saved... ");
                    return file.download({
                        // don't set destination here
                    });
                });
            runInAction(() => {
                if (DEBUG) console.log("Store " + path + " to store (/templates)");
                store.set(path, fromPromise(res));
            });
            promises.push(
                res.then((buffer) => {
                    const file = bucket.file(path);
                    return file
                        .getSignedUrl({
                            action: "read",
                            expires: "03-09-2491",
                        })
                        .then((signedUrls) => {
                            // signedUrls[0] contains the file's public URL

                            return {
                                attachment: {
                                    url: signedUrls[0],
                                    filename: fn,
                                },
                            };
                        });
                })
            );
        } //end of for
    );

    return Promise.all(promises)
        .then((values) => {
            if (DEBUG) console.log("All promises resolved!");
            if (DEBUG) console.log(JSON.stringify(values));
            return res.json({
                status: "ok",
                data: values,
            });
        })
        .catch((err) => {
            console.log("Error " + err);
            res.json({
                status: "error",
                errors: ["" + err],
            });
        });
});

var multer_upload = multer({
    storage: multer.memoryStorage(),
}).single('source');

app.post("/excel_to_json", (req, res) => {
    try {
        multer_upload(req, res, (err) => {
            if (err) {
                // An error occurred when uploading to server memory
                return res.json({
                    status: "error",
                    errors: ["" + err],
                });
            }
            let config = req.body;
            config.source = req.file.buffer;
            config.sheets = JSON.parse(config.sheets);
            // console.log(req.body);
            // console.log(req.files);
            // console.log(config);
            // var source = req.files[0].source;
            // console.log(JSON.stringify(source));      
            const result = excelToJson(config);
            console.log(JSON.stringify(config.type));
            console.log(JSON.stringify(config.sheets));
            return res.json({
                status: "ok",
                data: result,
            });
        });
    } catch (err) {
        if (DEBUG) console.log("Error " + err);
        res.json({
            status: "error",
            errors: ["" + err],
        });
    }
});

app.get("/", (req, res) => {
    res.sendFile(`${__dirname}/index.html`);
});

if (DEBUG) console.log("Listening...");
app.listen(8000);