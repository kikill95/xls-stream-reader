const fs = require('fs')
const path = require('path')

const filePath = path.resolve(__dirname, 'files', 'file3.xlsx')
// const filePath = path.resolve(__dirname, 'files', 'file2.xlsx')

let dataStream = fs.createReadStream(filePath)

const XLSXReader = require('../index')

let reader = new XLSXReader(dataStream)

console.log('starting parse')
reader
  .eachRow(console.log)
  .then(() => {
    console.log('finished')
  })
  .catch(console.error)
