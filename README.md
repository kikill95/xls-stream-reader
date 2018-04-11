# ek-xls-stream-reader

node package for reading **only** xlsx, xlsm files as streams

## Installation

  npm i ek-xls-stream-reader

## Usage

  ```js
  const ExcelReader = require('ek-xls-stream-reader')

  // and somewhere from readable stream like `req`

  let reader = new ExcelReader(req)

  reader
    .eachRow(data => {
      console.log(data)
    })
    .then(() => {
      console.log('Finished')
    })
  ```

## Development

  npm start

## Tests

  npm test

## License

MIT
