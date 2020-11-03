const functions = require('firebase-functions');
const logger = functions.logger;

const admin = require('firebase-admin');

var serviceAccount = require("./firebase/nodedefrais-firebase-adminsdk.json");

admin.initializeApp({
    credential: admin.credential.cert(serviceAccount),
});

const firestore = admin.firestore();

const excel = require("exceljs");
const cors = require("cors");
const express = require("express");
const app = express();
app.use(cors({ origin: true }));

// Date 
const moment = require('moment');
moment.locale('fr');

// --------------------------------------------------------
// constantes
// --------------------------------------------------------

// Xls
const XLS_SHEET_NAME = "N2F";
const XLS_FILE_NAME = "export-n2f.xlsx";
const XLS_HEADER_DATE = "Date";
const XLS_HEADER_AMOUNT = "Amount";
// const XLS_CELL_FORMAT_DATE = "dd/mm/yyyy";
const XLS_CELL_FORMAT_DATE = "@";
const XLS_CELL_FORMAT_AMOUNT = "# ###.00";

// Firebaase 
const FIREBASE_COLLECTION_DATAS = "datas";
const FIREBASE_FIELD_DATE = "date";
const FIREBASE_FIELD_AMOUNT = "amount";

const FIREBASE_COLLECTION_USERS = "users";

// --------------------------------------------------------
// api rest 
// --------------------------------------------------------

app.get('/isUp', (req, res) => {
    res.sendStatus(200);
})

app.get('/exportAll', (req, res) => {
    loadDatas(res, {});
})

app.get('/export/:uid/:year/:month', (req, res) => {
    var uid = req.params.uid;
    var year = req.params.year;
    var month = req.params.month;
    loadDatas(res, createFilter(uid, year, month));
})

// set region for api
const api = functions.region("europe-west1").https.onRequest(app);

module.exports = {
    api
}

// --------------------------------------------------------
// function 
// --------------------------------------------------------

function br(text) {
    return ((text) ? text : "") + "</br>"
}

function formatDate(d) {
    return new moment(d).format("L");
}

// function formatDate(d) {
//     let date = ("0" + d.getDate()).slice(-2);
//     let month = ("0" + (d.getMonth() + 1)).slice(-2);
//     let year = d.getFullYear();
//     return date + "/" + month + "/" + year;
// }

function createFilter(uid, year, month) {
    // logger.info("createFilter: uid=", uid, "year=", year, "month=", month);
    let start = new Date(year, month - 1, 1);
    start.setHours(0, 0, 0, 0);
    let end = new Date(year, (month - 1) + 1, 1);
    end.setHours(0, 0, 0, 0);
    return { uid: uid, start: start, end: end };
}

function loadDatas(res, filter) {
    // logger.info("filter: start=", formatDate(filter.start), "end=", formatDate(filter.end));
    let datas = [];
    let query = firestore.collection(`${FIREBASE_COLLECTION_USERS}/${filter.uid}/${FIREBASE_COLLECTION_DATAS}`);
    if (filter.start && filter.end) {
        query = query
            .where(FIREBASE_FIELD_DATE, ">=", filter.start)
            .where(FIREBASE_FIELD_DATE, "<", filter.end);
    }
    query = query.orderBy(FIREBASE_FIELD_DATE, "asc");
    const snapshot = query.get();
    snapshot.then(querySnapshot => {
        querySnapshot.docs.forEach(doc => {
            // date
            let d = doc.get(FIREBASE_FIELD_DATE).toDate();
            let rd = formatDate(d);
            // amount 
            let a = doc.get(FIREBASE_FIELD_AMOUNT);
            let ra = parseInt(a);
            // push in array datas 
            datas.push({ date: rd, amount: ra });
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
        // logger.info("date=", datas[i].date, ", amount=", datas[i].amount);
        const row_data = worksheet.addRow([datas[i].date, datas[i].amount]);
        row_data.getCell(1).numFmt = XLS_CELL_FORMAT_DATE;
        row_data.getCell(2).numFmt = XLS_CELL_FORMAT_AMOUNT;
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