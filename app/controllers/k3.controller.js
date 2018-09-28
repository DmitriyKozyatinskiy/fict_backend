const express = require('express');
const httpStatus = require('http-status');
const XlsxPopulate = require('xlsx-populate');

const router = express.Router({ mergeParams: true });

router.get('/', async (req, res, next) => {
  try {
    const exercises = await Exercise.find().lean();
    res.status(200).send(exercises);
  } catch(err) {
    return next(err);
  }
});

router.post('/', async (req, res, next) => {
  try {
    const templateWorkBook = await XlsxPopulate.fromFileAsync('./excel/k3_template.xlsx');
    templateWorkBook.sheet(0).cell('A14').value(req.body.values);
    const generatedFile = await templateWorkBook.outputAsync('base64');

    res.status(200).json({ data: generatedFile });
  } catch(err) {
    return next(err);
  }
});

module.exports = router;
// router.get('/download', function (req, res, next) {
//   console.log('run');
//   Promise.all([
//     XlsxPopulate.fromFileAsync('./excel/k3_template.xlsx'),
//     XlsxPopulate.fromFileAsync('./excel/test.xlsx'),
//   ])
//     .then(workbooks => {
//       const [templateWorkBook, valuesWorkBook] = workbooks;
//       const values = valuesWorkBook.sheet(0).usedRange().value();
//
//       const requiredValues = values.map(row => {
//         row.length = 22;
//         return row;
//       });
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
