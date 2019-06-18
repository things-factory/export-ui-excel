import { store } from '@things-factory/shell'
import { UPDATE_EXTENSION } from '@things-factory/export-base'

import * as XLSX from 'xlsx'

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

export default function bootstrap() {
  store.dispatch({
    type: UPDATE_EXTENSION,
    extensions: {
      xlsx: {
        export: jsonToExcel
      },
      xls: {
        export: jsonToExcel
      }
    }
  })
}
