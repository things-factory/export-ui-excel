import { UPDATE_EXTENSION } from '@things-factory/export-base'
import { store } from '@things-factory/shell'
import * as XLSX from 'xlsx'

function jsonToXslx(params) {
  jsonToExcel('xlsx', params)
}

function jsonToXls(params) {
  jsonToExcel('xls', params)
}

function jsonToExcel(exts, params) {
  const sheetName = params.sheetName
  const header = params.header
  const data = params.data

  const wb = XLSX.utils.book_new()
  const ws = XLSX.utils.json_to_sheet(data, { header })
  XLSX.utils.book_append_sheet(wb, ws, sheetName)
  XLSX.writeFile(wb, `${sheetName}.${exts}`, {
    bookType: exts
  })
}

export default function bootstrap() {
  store.dispatch({
    type: UPDATE_EXTENSION,
    extensions: {
      xlsx: {
        export: jsonToXslx
      },
      xls: {
        export: jsonToXls
      }
    }
  })
}
