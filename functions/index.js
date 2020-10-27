const functions = require('firebase-functions');
const logger = functions.logger;

const admin = require('firebase-admin');

var serviceAccount = require("./firebase/nodedefrais-firebase-adminsdk.json");

admin.initializeApp({
    credential: admin.credential.cert(serviceAccount),
});

const firestore = admin.firestore();

const cors = require("cors");
const express = require("express");
const app = express();

const excel = require("exceljs");

app.use(cors({ origin: true }));

const DATE_FORMAT = "dd/MM/yyyy";

const XLS_SHEET_NAME = "N2F";
const XLS_HEADER_DATE = "Date";
const XLS_HEADER_AMOUNT = "Amount";
const XLS_CELL_FORMAT_DECIMAL = "0.00";
const XLS_FILE_NAME = "export-n2f.xlsx";

const FIREBASE_COLLECTION = "datas";
const FIREBASE_FIELD_DATE = "date";
const FIREBASE_FIELD_AMOUNT = "amount";

// const moment = require('moment');
// moment().locale('fr');

// --------------------------------------------------------
// api rest 
// --------------------------------------------------------

app.get('/isUp', (req, res) => {
    res.sendStatus(200);
});

app.get('/exportAll', (req, res) => {
    loadDatas(res, {});
});

app.get('/export/:year/:month', (req, res) => {
    var year = req.params.year;
    var month = req.params.month;
    loadDatas(res, createFilter(year, month));
});

const api = functions.region("europe-west1").https.onRequest(app);

module.exports = {
    api
};

// --------------------------------------------------------
// function 
// --------------------------------------------------------

function createFilter(year, month) {
    let start = new Date(year, month - 1, 1);
    start.setHours(0, 0, 0, 0);
    let end = new Date(year, (month - 1) + 1, 1);
    end.setHours(0, 0, 0, 0);
    // logger.info("year=", year);
    // logger.info("month=", month);
    return { start: start, end: end };
}

function formartDate(d) {
    let date = ("0" + d.getDate()).slice(-2);
    let month = ("0" + (d.getMonth() + 1)).slice(-2);
    let year = d.getFullYear();
    return date + "/" + month + "/" + year;
}

// function br(text) {
//     return ((text) ? text : "") + "</br>"
// }

function loadDatas(res, filter) {
    let datas = [];
    let query = firestore.collection(FIREBASE_COLLECTION);
    // logger.info("filter.start=", formartDate(filter.start));
    // logger.info("filter.end=", formartDate(filter.end));
    if (filter.start && filter.end) {
        query = query
            .where("date", ">=", filter.start)
            .where("date", "<", filter.end);
    }
    query = query.orderBy(FIREBASE_FIELD_DATE, "asc");
    const snapshot = query.get();
    snapshot.then(querySnapshot => {
        querySnapshot.docs.forEach(doc => {
            let d = doc.get(FIREBASE_FIELD_DATE).toDate();
            let r = formartDate(d);
            datas.push({
                date: r,
                amount: doc.get(FIREBASE_FIELD_AMOUNT)
            });
        });
        createXls(res, datas);
        return;
    }).catch(error => {
        logger.log("Error: ", error);
        res.sendStatus(500);
        return;
    });
}

function createXls(res, datas) {
    const workbook = new excel.Workbook();
    const worksheet = workbook.addWorksheet(XLS_SHEET_NAME);

    // row header
    const row = worksheet.addRow([XLS_HEADER_DATE, XLS_HEADER_AMOUNT]);
    row.font = { bold: true };

    // row data
    for (let i = 0; i < datas.length; i += 1) {
        logger.info("date=", datas[i].date, ", amount=", datas[i].amount);
        const row_data = worksheet.addRow([datas[i].date, datas[i].amount]);
        // row_data.getCell(1).numFmt = "dd/mm/yyyy";
        // row_data.getCell(2).numFmt = "0.00";
        row_data.getCell(1).numFmt = "@";
        row_data.getCell(2).numFmt = "@";
        row_data.getCell(2).alignment = { horizontal: 'right' };
    }

    // create file
    res.setHeader('Content-Disposition', 'attachment; filename=' + XLS_FILE_NAME);
    res.setHeader('Content-Type', 'application/octet-stream');

    // set column with
    worksheet.getColumn(1).width = 12;
    worksheet.getColumn(2).width = 12;

    // xls send to response
    workbook.xlsx.write(res);
}