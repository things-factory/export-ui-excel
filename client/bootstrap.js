import { store } from '@things-factory/shell'
import { UPDATE_EXTENSION } from '@things-factory/export-base'
import * as XLSX from '!xlsx'
import Excel from '!exceljs'
import { saveAs } from 'file-saver'

const _ = require('lodash')

async function jsonToExcel({ extension, name, data }) {
  if (data === 0) {
    return
  }

  const sheetName = name
  const records = typeof data == 'function' ? await data.call() : data

  const header = Object.keys(records[0])

  const wb = XLSX.utils.book_new()
  const ws = XLSX.utils.json_to_sheet(records, { header })

  XLSX.utils.book_append_sheet(wb, ws, sheetName)
  XLSX.writeFile(wb, `${sheetName}.${extension}`, {
    bookType: extension,
  })
}

/**
 * Convert Object data with fixed structure into Excel.
 * @param {string} extension - Name of the file extension. eg. xls, xlsx
 * @param {string} name - Name of the file.
 * @param {Object} data - { header: [{headerName, fieldName, type, arrData}], data: [{fieldName: value}], groups:[{ column, title }], totals: [value], sheetStyle:{} }
 * @return file - Serve the file to client using file-saver.
 */
async function objDataToExcel({ extension, name, data }) {
  try {
    // data structure
    // {
    //    Styles Not Implemented yet. For future development.
    //    header: [{headerName, fieldName, type, arrData}], data: [{fieldName: value}], groups:[{ column, title }], totals: [value], sheetStyle:{}
    // }

    const records = typeof data == 'function' ? await data.call() : data
    // ////Perform excel file manipulation. Requirement: import Excel from 'exceljs'
    const EXCEL_FORMATS = {
      DATE: { numFmt: 'dd.mmm.yyyy' },
      DATE_TIME: { numFmt: 'dd.mmm.yyyy hh:mm' },
      TIME: { numFmt: 'hh:mm' },
    }

    const wb = new Excel.Workbook(name)
    const ws = wb.addWorksheet(name)
    let header = [
      { header: 'id', key: 'id', width: 5 },
      ...records.header.map((column) => {
        return {
          header: column.header || '',
          key: column.key || '',
          width: column.width || undefined,
        }
      }),
    ]

    var headerObjStructure = records.header.reduce(function (obj, item) {
      obj[item.key] = ''
      return obj
    }, {})

    ws.columns = header
    ws._rows[0]._cells.map((cell, index) => {
      cell.name = header[index].key

      // ////Set Header Cell Fill
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF0C5394' },
      }
      // ////Set Header Cell Font style
      cell.font = {
        name: 'Arial',
        color: { argb: 'FFFFFFFF' },
        family: 1,
        size: 12,
        bold: true,
      }
      // ////Set Header Cell Alignment
      cell.alignment = {
        vertical: 'middle',
        horizontal: 'center',
      }
    })

    let [alternateA, alternateB] = ['FFFFFF', 'F3F3F3']
    let printData = JSON.parse(JSON.stringify(records.data))

    printData = printData.map((data) => {
      return { ...headerObjStructure, ...data }
    })

    if (!!records.groups && records.groups.length > 0) {
      printData = multiGroupTree(
        printData,
        records.groups.map((itm) => itm.column),
        records.groups,
        records.totals
      )
    } else {
      printData = printData.map((row) => {
        ;[alternateA, alternateB] = [alternateB, alternateA]
        return {
          data: row,
          style: {
            row: {
              // ////Set alternate Cell Fill
              fill: {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: alternateA },
              },
              border: {
                bottom: { style: 'thin' },
              },
            },
          },
        }
      })
    }

    // ////Set each cell's design
    printData.forEach((row, index) => {
      ws.addRow(row.data)

      if (!!row.style) {
        ws._rows[index + 1]._cells.forEach((cell) => {
          let cellStyle = row.style[cell._column._key] || {}
          cell.style = { ...cell.style, ...row.style.row, ...cellStyle }
        })
      }
    })

    ws.addRow({ id: '' })
    ws.getColumn('id').hidden = true

    // Cell Type: [ list, whole, decimal, textLength, date ]
    records.header
      .filter((column) => column.type === 'array' && column.arrData instanceof Array)
      .map(async (column) => {
        let dataWs = {}
        if (!wb.getWorksheet(column.key)) {
          dataWs = wb.addWorksheet(column.key)
          let header = Object.keys(column.arrData[0]).map((column) => {
            return {
              header: column || '',
              key: column || '',
            }
          })
          dataWs.columns = header
          dataWs.addRows(column.arrData)
          dataWs._rows[0]._cells.map((cell, index) => {
            cell.name = header[index].key
          })

          dataWs.state = 'veryHidden'
          await dataWs.protect(Math.random().toString(36).substring(2), {
            selectLockedCells: false,
            selectUnlockedCells: false,
          })
        } else {
          dataWs = wb.getWorksheet(column.key)
        }

        let charColumnCode = String.fromCharCode(
          97 + dataWs.columns.findIndex((ind) => ind._key === 'name')
        ).toUpperCase()

        ws.getColumn(column.key).eachCell(function (cell, rowNumber) {
          if (rowNumber !== 1)
            cell.dataValidation = {
              type: 'list',
              allowBlank: true,
              formulae: [
                dataWs.name + '!$' + charColumnCode + '$2:$' + charColumnCode + '$' + dataWs.rowCount.toString(),
              ],
            }
        })
      })

    records.header
      .filter((column) => column.type === 'int')
      .map((column) => {
        ws.getColumn(column.key).eachCell(function (cell, rowNumber) {
          if (rowNumber !== 1)
            cell.dataValidation = {
              type: 'whole',
              allowBlank: true,
            }
        })
      })

    records.header
      .filter((column) => column.type === 'float')
      .map((column) => {
        ws.getColumn(column.key).eachCell(function (cell, rowNumber) {
          if (rowNumber !== 1)
            cell.dataValidation = {
              type: 'decimal',
              allowBlank: true,
            }
        })
      })

    records.header
      .filter((column) => column.type === 'date')
      .map((column) => {
        ws.getColumn(column.key).eachCell(function (cell, rowNumber) {
          if (rowNumber !== 1) {
            cell.dataValidation = {
              type: 'date',
              allowBlank: true,
            }
            cell.value = cell.value ? new Date(parseInt(cell.value)) : new Date()
          }
        })
      })

    //Save as file using "file-saver". Requirement: import { saveAs } from 'file-saver'
    await wb.xlsx.writeBuffer(EXCEL_FORMATS).then((buffer) => {
      saveAs(
        new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }),
        name + '.xlsx'
      )
    })
  } catch (e) {
    throw e
  }
}

export default function bootstrap() {
  store.dispatch({
    type: UPDATE_EXTENSION,
    extensions: {
      xlsx: {
        export: objDataToExcel,
      },
      xls: {
        export: jsonToExcel,
      },
    },
  })
}

function multiGroupTree(array, groups, groupsRaw, totals) {
  if (!groups) {
    return array
  }
  const currentGroup = groups[0]
  const restGroups = [...groups.slice(1, groups.length)]
  let grouping = _.groupBy(array, currentGroup)

  if (!restGroups.length) {
    let rows = []
    Object.keys(grouping).forEach((itm) => {
      let currentGroupSetting = groupsRaw.filter((x) => x.column === currentGroup)[0]

      let [alternateA, alternateB] = ['F3F3F3', 'FFFFFF']
      grouping[itm] = grouping[itm].map((itm, index) => {
        if (index != 0) itm[currentGroup] = ''
        ;[alternateA, alternateB] = [alternateB, alternateA]

        return {
          data: itm,
          style: {
            row: {
              fill: {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: alternateA },
              },
            },
          },
        }
      })

      let newRow = []
      if (currentGroupSetting.title) {
        let sumData = stripObject(grouping[itm][0])
        sumData.data[currentGroup] = currentGroupSetting.title

        totals.forEach((total) => {
          sumData.data[total] = grouping[itm].reduce((acc, obj) => {
            acc = acc + (parseFloat(obj.data[total]) || 0)
            return acc
          }, 0)
        })

        sumData.style = getGroupRowStyle([currentGroup])
        newRow = [sumData]
      }

      rows = [...grouping[itm], ...newRow, ...rows]
    })

    return rows
  }

  let result = _.transform(
    grouping,
    (result, value, key) => {
      let rows = multiGroupTree(value, restGroups, groupsRaw, totals)

      let currentGroupSetting = groupsRaw.filter((x) => x.column === currentGroup)[0]

      rows.map((itm, index) => {
        if (index != 0) itm.data[currentGroup] = ''
        return itm
      })

      let newRow = []
      if (currentGroupSetting.title) {
        let sumData = stripObject(rows[0])
        sumData.data[currentGroup] = currentGroupSetting.title

        totals.forEach((total) => {
          sumData.data[total] = value.reduce((acc, obj) => {
            acc = acc + (parseFloat(obj[total]) || 0)
            return acc
          }, 0)
        })

        sumData.style = getGroupRowStyle([currentGroup])
        newRow = [sumData]
      }

      rows = [...rows, ...newRow]

      if (groupsRaw.findIndex((x) => x.column == currentGroupSetting.column) === 0) {
        rows.forEach((row, index) => {
          row.style = {
            ...row.style,
            [currentGroupSetting.column]: {
              fill: {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFCFE2F3' },
              },
              font: {
                name: 'Arial',
                color: { argb: 'FF000000' },
                family: 1,
                bold: true,
              },
            },
          }

          if (index != rows.length - 1) {
            row.style = {
              ...row.style,
              [currentGroupSetting.column]: {
                ...row.style[currentGroupSetting.column],
                border: {
                  top: { style: '' },
                  bottom: { style: '' },
                },
              },
            }
          }
        })

        rows[rows.length - 1].style = {
          ...rows[rows.length - 1].style,
          row: {
            ...rows[rows.length - 1].style.row,
            border: {
              bottom: { style: 'thin' },
            },
          },
        }
      }

      rows.map((row) => result.push(row))
    },
    []
  )
  return result
}

function stripObject(source) {
  var o = Array.isArray(source) ? [] : {}
  for (var key in source) {
    if (source.hasOwnProperty(key)) {
      var t = typeof source[key]
      o[key] =
        source[key] === null
          ? null
          : t == 'object'
          ? stripObject(source[key])
          : { string: '', number: 0, boolean: false }[t]
    }
  }
  return o
}

function getGroupRowStyle(groupColumnName) {
  return {
    [groupColumnName]: {
      alignment: {
        horizontal: 'center',
      },
    },
    row: {
      border: {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
      },
      fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFEEF7FF' },
      },
      font: {
        name: 'Arial',
        color: { argb: 'FF000000' },
        family: 1,
        bold: true,
      },
    },
  }
}
