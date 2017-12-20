const fs = require('fs')
const path = require('path')

const ExcelReader = require('../index')

const filePath = path.resolve(__dirname, 'files', 'file1.xls')
// const filePath = path.resolve(__dirname, 'files', 'file2.xlsx')

let dataStream = fs.createReadStream(filePath)
let reader = new ExcelReader(dataStream)

console.log('starting parse')
reader
  .eachRow((rowData, rowNum, sheetSchema) => {
    console.log(rowData)
  })
  .then(() => {
    console.log('done parsing')
  })
  .catch(console.error)
