import { store } from '@things-factory/shell'
import { UPDATE_EXTENSION } from '@things-factory/export-base'

import * as XLSX from '!xlsx'
import Excel from '!exceljs'
import { saveAs } from 'file-saver'

function jsonToExcel({ extension, name, data }) {
  if (data === 0) {
    return
  }

  const sheetName = name
  const records = typeof data == 'function' ? data.call() : data

  const header = Object.keys(records[0])

  const wb = XLSX.utils.book_new()
  const ws = XLSX.utils.json_to_sheet(records, { header })

  XLSX.utils.book_append_sheet(wb, ws, sheetName)
  XLSX.writeFile(wb, `${sheetName}.${extension}`, {
    bookType: extension
  })
}

/**
 * Convert Object data with fixed structure into Excel.
 * @param {string} extension - Name of the file extension. eg. xls, xlsx
 * @param {string} name - Name of the file.
 * @param {Object} data - { header: [{headerName, fieldName, type, arrData}], data: [{fieldName: value}, sheetStyle:{}] }
 * @return file - Serve the file to client using file-saver.
 */
async function objDataToExcel({ extension, name, data }) {
  // data structure
  // {
  //    Styles Not Implemented yet. For future development.
  //    header: [{headerName, fieldName, type, arrData}], data: [{fieldName: value}, sheetStyle:{}]
  // }

  const records = typeof data == 'function' ? data.call() : data
  // ////Perform excel file manipulation. Requirement: import Excel from 'exceljs'
  const EXCEL_FORMATS = {
    DATE: { numFmt: 'dd.mmm.yyyy' },
    DATE_TIME: { numFmt: 'dd.mmm.yyyy hh:mm' },
    TIME: { numFmt: 'hh:mm' }
  }

  const wb = new Excel.Workbook(name)
  const ws = wb.addWorksheet(name)
  let header = [
    { header: 'id', key: 'id', width: 5 },
    ...records.header.map(column => {
      return {
        header: column.header || '',
        key: column.key || '',
        width: column.width || undefined
      }
    })
  ]
  ws.columns = header
  ws._rows[0]._cells.map((cell, index) => {
    cell.name = header[index].key
  })
  ws.addRows(records.data)
  ws.addRow({ id: '' })
  ws.getColumn('id').hidden = true

  // Cell Type: [ list, whole, decimal, textLength, date ]
  records.header
    .filter(column => column.type === 'array' && column.arrData instanceof Array)
    .map(async column => {
      let dataWs = {}
      if (!wb.getWorksheet(column.key)) {
        dataWs = wb.addWorksheet(column.key)
        let header = Object.keys(column.arrData[0]).map(column => {
          return {
            header: column || '',
            key: column || ''
          }
        })
        dataWs.columns = header
        dataWs.addRows(column.arrData)
        dataWs._rows[0]._cells.map((cell, index) => {
          cell.name = header[index].key
        })

        dataWs.state = 'veryHidden'
        await dataWs.protect(
          Math.random()
            .toString(36)
            .substring(2),
          {
            selectLockedCells: false,
            selectUnlockedCells: false
          }
        )
      } else {
        dataWs = wb.getWorksheet(column.key)
      }

      let charColumnCode = String.fromCharCode(97 + dataWs.columns.findIndex(ind => ind._key === 'name')).toUpperCase()

      ws.getColumn(column.key).eachCell(function(cell, rowNumber) {
        if (rowNumber !== 1)
          cell.dataValidation = {
            type: 'list',
            allowBlank: true,
            formulae: [dataWs.name + '!$' + charColumnCode + '$2:$' + charColumnCode + '$' + dataWs.rowCount.toString()]
          }
      })
    })

  records.header
    .filter(column => column.type === 'int')
    .map(column => {
      ws.getColumn(column.key).eachCell(function(cell, rowNumber) {
        if (rowNumber !== 1)
          cell.dataValidation = {
            type: 'whole',
            allowBlank: true
          }
      })
    })

  records.header
    .filter(column => column.type === 'float')
    .map(column => {
      ws.getColumn(column.key).eachCell(function(cell, rowNumber) {
        if (rowNumber !== 1)
          cell.dataValidation = {
            type: 'decimal',
            allowBlank: true
          }
      })
    })

  records.header
    .filter(column => column.type === 'date')
    .map(column => {
      ws.getColumn(column.key).eachCell(function(cell, rowNumber) {
        if (rowNumber !== 1) {
          cell.dataValidation = {
            type: 'date',
            allowBlank: true
          }
          cell.value = cell.value ? new Date(parseInt(cell.value)) : new Date()
        }
      })
    })

  //Save as file using "file-saver". Requirement: import { saveAs } from 'file-saver'
  await wb.xlsx.writeBuffer(EXCEL_FORMATS).then(buffer => {
    saveAs(
      new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }),
      name + '.xlsx'
    )
  })
}

export default function bootstrap() {
  store.dispatch({
    type: UPDATE_EXTENSION,
    extensions: {
      xlsx: {
        export: objDataToExcel
      },
      xls: {
        export: jsonToExcel
      }
    }
  })
}
