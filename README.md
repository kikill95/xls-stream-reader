# ek-xls-stream-reader

node package for reading xlsx, xlsm files as streams

## Installation

    npm i ek-xls-stream-reader

## Usage

  ```js
  const XLSXReader = require('ek-xls-stream-reader')

  // and somewhere from readable stream like `req`

  let reader = new XLSXReader(req)

  reader
    .eachRow(data => {
      // do something with `data` here
    })
    .then(() => {
      console.log('Finished')
    })
  ```

## Be aware

  This package works only with first sheet where first **50** rows aren't empty

## Development (with test files)

  npm start

## Tests code style

  npm test

## License

MIT
