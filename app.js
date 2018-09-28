const express = require('express');
const cors = require('cors');
const compress = require('compression');
const methodOverride = require('method-override');
const helmet = require('helmet');
const httpStatus = require('http-status');
const createError = require('http-errors');
const errorHandler = require('api-error-handler');
// const XlsxPopulate = require('xlsx-populate');

const K3Controller = require('./app/controllers/k3.controller');

const app = express();

app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(errorHandler());
app.use(compress());
app.use(methodOverride());
app.use(helmet());
app.use(cors());

app.use('/api/k3', K3Controller);

// const router = express.Router({ mergeParams: true });
//
//
//
// router.get('/download', function (req, res, next) {
//   console.log('run');
//   Promise.all([
//     XlsxPopulate.fromFileAsync('./excel/k3_template.xlsx'),
//     XlsxPopulate.fromFileAsync('./excel/test.xlsx'),
//   ])
//     .then(workbooks => {
//       console.log('HELLO!!!');
//       const [templateWorkBook, valuesWorkBook] = workbooks;
//       const values = valuesWorkBook.sheet(0).usedRange().value();
//       // const usedTemplateRange = templateWorkBook.sheet(0).usedRange();
//
//       const requiredValues = values.map(row => {
//         row.length = 22;
//         return row;
//       });
//
//       console.log('VALUES: ', values);
//       templateWorkBook.sheet(0).cell('A14').value(requiredValues);
//
//
//       // console.log(usedTemplateRange.address());
//
//
//       // console.log(valuesWorkBook.);
//       // console.log(valuesSheet.values());
//
//       // sheets2.forEach(sheet => {
//       //   const newSheet = workbook.addSheet(sheet.name());
//       //   const usedRange = sheet.usedRange();
//       //   const oldValues = usedRange.value();
//       //
//       //   newSheet.range(usedRange.address()).value(oldValues);
//       // });
//       //
//
//       templateWorkBook.toFileAsync('./excel/hello.xlsx');
//       return requiredValues;//templateWorkBook.toFileAsync('./excel/hello.xlsx');
//     }).then(data => {
//     // res.attachment('output.xlsx');
//     // res.send(data);
//     res.send(data);
//   }).catch(err => {
//     console.log('Err: ', err);
//     res.send(null);
//   });
// });
//
// app.use('/', router);

app.use((err, req, res, next) => {
  res.locals.message = err.message;
  res.locals.error = err;

  res.status(err.status || 500).json(
    createError(err.statusCode || httpStatus.INTERNAL_SERVER_ERROR, err.message, {
      expose: true,
    }),
  );
  next();
});

module.exports = app;
