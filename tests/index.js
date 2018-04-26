const fs = require('fs')
const path = require('path')

const filePath = path.resolve(__dirname, 'files', 'file.xlsx')

let dataStream = fs.createReadStream(filePath)

const XLSXReader = require('../index')

let reader = new XLSXReader(dataStream)

console.log('starting parse')
reader
  .eachRow(data => {
    // do something with `data` here
  })
  .then(() => {
    console.log('finished')
  })
  .catch(console.error)
