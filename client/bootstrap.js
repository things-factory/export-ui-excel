import { UPDATE_EXTENSION } from '@things-factory/export-base'
import { store } from '@things-factory/shell'

function jsonToExcel() {}

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
