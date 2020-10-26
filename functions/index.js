// /* eslint-disable promise/catch-or-return */
// /* eslint-disable promise/always-return */

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

const XLS_SHEET_NAME = "N2F";
const XLS_HEADER_DATE = "Date";
const XLS_HEADER_AMOUNT = "Amount";
const XLS_CELL_FORMAT_DECIMAL = "0.00";
const XLS_FILE_NAME = "export-n2f.xlsx";

const FIREBASE_COLLECTION = "datas";
const FIREBASE_FIELD_DATE = "date";
const FIREBASE_FIELD_AMOUNT = "amount";

app.get('/isUp', (req, res) => {
    res.sendStatus(200);
});

app.get('/exportAll', (req, res) => {

    logger.info("exportAll");

    let datas = getDatas();

    for (let i = 0; i < datas.length; i += 1) {
        logger.info("date=" + datas[i].date);
        logger.info("amount=" + datas[i].amount);
    }

    createXlsWithDatas(res, datas);

    // const colRef = firestore.collection(FIREBASE_COLLECTION);
    // // var query = colRef.orderBy(FIREBASE_FIELD_DATE, Query.Direction.ASCENDING);
    // // query.get()
    // const snapshot = colRef.get();
    // snapshot.then(querySnapshot => {
    //     querySnapshot.docs.forEach(doc => {
    //         logger.info("date: " + doc.get(FIREBASE_FIELD_DATE));
    //         logger.info("amount: " + doc.get(FIREBASE_FIELD_AMOUNT));
    //     });
    // });

    // snapshot.then(querySnapshot => {
    //     querySnapshot.docs.forEach(doc => {
    //         logger.info("doc ; " + doc.data());
    //     });
    // }).catch((e) => {
    //     logger.error(e);
    // });

    // logger.info("create array datas ...");
    // let datas = [];
    // datas.push({ date: "my_date", amount: "my_amount" });
    // for (let i = 0; i < datas.length; i += 1) {
    //     logger.info("date=" + datas[i].date);
    //     logger.info("amount=" + datas[i].amount);
    // }
    // logger.info("create array datas.");

    // createXls(res);
    res.sendStatus(200);
});

app.get('/export/:year/:month', (req, res) => {
    var year = req.params.year;
    var month = req.params.month;
    logger.info("year=", year);
    logger.info("month=", month);
    res.sendStatus(200);
});

app.get('/test', (req, res) => {
    createXls(res);
});

function getDatas() {
    let datas = [];
    const colRef = firestore.collection(FIREBASE_COLLECTION);
    // var query = colRef.orderBy(FIREBASE_FIELD_DATE, Query.Direction.ASCENDING);
    // query.get()
    const snapshot = colRef.get();
    snapshot.then(querySnapshot => {
        querySnapshot.docs.forEach(doc => {
            datas.push({
                date: doc.get(FIREBASE_FIELD_DATE),
                amount: doc.get(FIREBASE_FIELD_AMOUNT)
            });

        });
        return datas;
    }).catch(error => {
        return datas;
    });
    return datas;
}

// function log(datas) {
//     for (let i = 0; i < datas.length; i += 1) {
//         logger.info("datas");
//     }
// }

function createXls(res) {
    logger.info("createXls");
    const workbook = new excel.Workbook();
    workbook.addWorksheet(XLS_SHEET_NAME);
    res.setHeader('Content-Type', 'application/octet-stream');
    res.setHeader('Content-Disposition', 'attachment; filename=' + XLS_FILE_NAME);
    workbook.xlsx.write(res);
}

function createXlsWithDatas(res, datas) {
    logger.info("createXls");
    const workbook = new excel.Workbook();
    workbook.addWorksheet(XLS_SHEET_NAME);
    res.setHeader('Content-Type', 'application/octet-stream');
    res.setHeader('Content-Disposition', 'attachment; filename=' + XLS_FILE_NAME);
    workbook.xlsx.write(res);
}


const api = functions.https.onRequest(app);

module.exports = {
    api
};