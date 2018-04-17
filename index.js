const Excel = require('exceljs')
const _ = require('lodash')
const dateFormat = require('dateformat')

const checkHeaderCount = 50 // WARNING: HARDCODED
const checkIsEmptyCount = 50 // WARNING: HARDCODED

class ExcelReader {
  constructor (dataStream, config, options) {
    this.stream = dataStream
    this.config = config || {}
    this.options = options || {sheets: []}
    this.workbook = new Excel.Workbook()
    this.afterRead = this._read()
  }

  _read () {
    return this.workbook.xlsx.read(this.stream)
      .then((workbook) => {
        if (this.options.sheets.length === 0) {
          // generate sheet if non passed in options
          workbook._worksheets = workbook._worksheets.filter(el => {
            if (el && el._rows && el._rows.slice(0, checkIsEmptyCount).reduce((acc, row) => acc + (row && row._cells ? row._cells.filter(el => el && el.value).length : 0), 0)) {
              // if sheet has any data in `checkIsEmptyCount` rows
              return true
            } else {
              return false
            }
          })
          let sheet = workbook._worksheets[0] // work only with first non empty and valid sheet
          let headerIndex = 0
          let biggestLength = 0
          for (let index = 0; index < checkHeaderCount; index++) {
            let row = sheet._rows[index]
            if (row) {
              let filtered = row._cells
                .map(el => el && el.value)
                .filter((value, index, array) => value && array.indexOf(value) === index)
              if (filtered.length > biggestLength) {
                biggestLength = filtered.length
                headerIndex = index
              }
            }
          }
          let headerRow = sheet._rows[headerIndex]
          this.options.sheets.push({
            name: sheet.name,
            rows: {
              headerRow: headerIndex + 1,
              allowedHeaders: headerRow && headerRow._cells.map(cell => {
                return {
                  name: cell.value,
                  key: cell.value
                }
              })
            }
          })
        }
        return this.workbook
      })
  }

  /**
  * Returns a json version of the row data, based on the
  * allowedHeaders of the given sheet.
  */
  _getRowData (rowObject, rowNum, allowedHeaders, headerRowValues) {
    let result = {}
    // predefine with empty strings
    headerRowValues.forEach(function (headerValue) {
      result[headerValue] = ''
    })
    rowObject.eachCell((cell, cellNo) => {
      // Finding the header value at this index
      if (!cell) {
        return
      }
      let header = headerRowValues[cellNo]
      if (header) {
        let currentHeader = _.find(allowedHeaders, {name: header})
        let cellValue = cell.value
        try {
          if (_.isObject(cell.value)) {
            // If this is an object, then a formula has been applied
            // We just take the result in that case
            // If this is a date, then we work with another format
            cellValue = cell.value.result
            if (!cellValue) {
              if (cell.value.toDateString) {
                cellValue = dateFormat(cell.value, cell.style.numFmt.toLowerCase())
              } else {
                cellValue = cell.value
              }
            }
            if (typeof cellValue === 'object') {
              try {
                // case when we have `rich text` - text that consists texts with different styles
                cellValue = Object.values(cell.value).reduce((acc, cur) => acc.concat(cur), []).map(el => el.text).join(' ')
              } catch (e) {}
            }
          }
        } catch (e) {
          cellValue = cell.value
        }
        result[currentHeader.key] = cellValue
      }
    })
    return result
  }

  /**
  * Takes a callback and runs it on every row of the every sheet, one by one.
  * Order of the sheets is not guaranteed.
  * This method provides each row in a json format based on the headers picked
  * up from options
  * callback params-
  *  1. rowData, a json with key being the header field, picked up from options.row
  *  2. rowNum, counting the headerRow
  *  3. sheetKey, key of the sheet. If no key exists, the name is provided
  *  The callback must return a promise
  */
  eachRow (callback) {
    return this.afterRead.then(async () => {
      let worksheet = this.workbook._worksheets[0] // work only with first sheet
      let sheetName = worksheet.name
      let sheetOptions = _.find(this.options.sheets, {name: worksheet.name})
      let sheetKey = sheetOptions.key ? sheetOptions.key : sheetName
      let headerRow = sheetOptions.rows.headerRow ? sheetOptions.rows.headerRow : 1
      let allowedHeaders = sheetOptions.rows.allowedHeaders
      let headerRowValues = worksheet.getRow(headerRow).values
      for (let rowNum = headerRow + 1; rowNum <= worksheet.rowCount; rowNum++) {
        // processing the rest rows
        let normalizedRowNum = rowNum - headerRow
        try {
          let rowData = this._getRowData(worksheet.getRow(rowNum), normalizedRowNum, allowedHeaders, headerRowValues)
          await callback(rowData, normalizedRowNum, sheetKey)
        } catch (e) {}
      }
      return Promise.resolve()
    })
  }
}

module.exports = ExcelReader
