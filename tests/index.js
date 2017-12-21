const fs = require('fs')
const path = require('path')

// const ExcelReader = require('../index')

const filePath = path.resolve(__dirname, 'files', 'file1.xls')
// const filePath = path.resolve(__dirname, 'files', 'file2.xlsx')

let dataStream = fs.createReadStream(filePath)
// let reader = new ExcelReader(dataStream)

// console.log('starting parse')
// reader
//   .eachRow((rowData, rowNum, sheetSchema) => {
//     console.log(rowData)
//   })
//   .then(() => {
//     console.log('done parsing')
//   })
//   .catch(console.error)

const XlsxStreamReader = require('xlsx-stream-reader')

var workBookReader = new XlsxStreamReader()
workBookReader.on('error', function (error) {
  throw (error)
})
workBookReader.on('sharedStrings', function () {
  // console.log(workBookReader.workBookSharedStrings)
})

workBookReader.on('styles', function () {
  // console.log(workBookReader.workBookStyles)
})

workBookReader.on('worksheet', function (workSheetReader) {
  if (workSheetReader.id > 1) {
    // we only want first sheet
    workSheetReader.skip()
    return
  }
  // print worksheet name
  // console.log(workSheetReader.name)

  workSheetReader.on('row', function (row) {
    if (row.attributes.r === 1) {
      // do something with row 1 like save as column names
    } else {
      // second param to forEach colNum is very important as
      // null columns are not defined in the array, ie sparse array
      row.values.forEach(function (rowVal, colNum) {
        // do something with row values
        console.log(rowVal)
      })
    }
  })
  workSheetReader.on('end', function () {
    console.log(workSheetReader.rowCount)
  })

  // call process after registering handlers
  workSheetReader.process()
})
workBookReader.on('end', function () {
  // end of workbook reached
})

dataStream.pipe(workBookReader)
