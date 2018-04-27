const XlsxStreamReader = require('xlsx-stream-reader')

class XLSXReader {
  constructor (dataStream, customOptions = {}, customValidations = {}) {
    const defaultOptions = {
      sheetIndex: null,
      headers: [],
      validations: {
        checkEmptyCount: 50
      }
    }
    this.options = Object.assign(defaultOptions, customOptions)
    this.options.validations = Object.assign(defaultOptions.validations, customValidations)
    this.stream = dataStream
    this.headerIndex = null
  }

  _formatData (data) {
    let result = {}
    this.options.headers.forEach((header, index) => {
      if (header) {
        result[header] = data[index]
      }
    })
    return result
  }

  eachRow (callback) {
    return new Promise((resolve, reject) => {
      const workbookStream = new XlsxStreamReader({
        verbose: false,
        formatting: true
      })
      workbookStream.on('error', error => {
        reject(error)
      })
      let sheetIndex = 0
      workbookStream.on('worksheet', workSheetReader => {
        if (this.options.sheetIndex !== null && this.options.sheetIndex !== sheetIndex) {
          return workSheetReader.skip()
        }
        sheetIndex++
        let savedRows = []
        let index = 0
        let headerCount = 0
        let workWithFirstBanch = async () => {
          if (savedRows.length > 0 && this.options.headers.length === 0) {
            this.options.sheetIndex = sheetIndex
            this.options.headers = savedRows[this.headerIndex] || []
            savedRows = savedRows.slice(this.headerIndex + 1)
            for (let data of savedRows) {
              await callback(this._formatData(data))
            }
            savedRows = []
          }
        }
        workSheetReader.on('row', async row => {
          if (index < this.options.validations.checkEmptyCount) {
            let filteredCount = row.values.filter(el => el).length
            if (filteredCount > headerCount) {
              headerCount = filteredCount
              this.headerIndex = index
            }
            savedRows.push(row.values)
            index++
          } else {
            if (this.headerIndex === null) {
              return workSheetReader.skip()
            }
            if (this.options.headers.length === 0) {
              await workWithFirstBanch()
            } else {
              await callback(this._formatData(row.values))
            }
          }
        })
        workSheetReader.on('end', async () => {
          if (savedRows.length > 0 && this.options.headers.length === 0) {
            await workWithFirstBanch()
          }
        })
        workSheetReader.process()
      })
      workbookStream.on('end', () => {
        resolve()
      })

      this.stream.pipe(workbookStream)
    })
  }
}

module.exports = XLSXReader
