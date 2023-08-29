/* eslint-disable */
// import { saveAs } from 'file-saver'
import XLSX from 'xlsx-style'

var charMap = [
  'A',
  'B',
  'C',
  'D',
  'E',
  'F',
  'G',
  'H',
  'I',
  'J',
  'K',
  'L',
  'M',
  'N',
  'O',
  'P',
  'Q',
  'R',
  'S',
  'T',
  'U',
  'V',
  'W',
  'X',
  'Y',
  'Z',
]
var columnMap = []
var maxColumn = 100
function generateArray(table) {
  var out = []
  var rows = table.querySelectorAll('tr')
  var ranges = []
  for (var R = 0; R < rows.length; ++R) {
    var outRow = []
    var row = rows[R]
    var columns = row.querySelectorAll('td')
    for (var C = 0; C < columns.length; ++C) {
      var cell = columns[C]
      var colspan = cell.getAttribute('colspan')
      var rowspan = cell.getAttribute('rowspan')
      var cellValue = cell.innerText
      if (cellValue !== '' && cellValue == +cellValue) cellValue = +cellValue

      //Skip ranges
      ranges.forEach(function(range) {
        if (
          R >= range.s.r &&
          R <= range.e.r &&
          outRow.length >= range.s.c &&
          outRow.length <= range.e.c
        ) {
          for (var i = 0; i <= range.e.c - range.s.c; ++i) outRow.push(null)
        }
      })

      //Handle Row Span
      if (rowspan || colspan) {
        rowspan = rowspan || 1
        colspan = colspan || 1
        ranges.push({
          s: {
            r: R,
            c: outRow.length,
          },
          e: {
            r: R + rowspan - 1,
            c: outRow.length + colspan - 1,
          },
        })
      }

      //Handle Value
      outRow.push(cellValue !== '' ? cellValue : null)

      //Handle Colspan
      if (colspan) for (var k = 0; k < colspan - 1; ++k) outRow.push(null)
    }
    out.push(outRow)
  }
  return [out, ranges]
}

function datenum(v, date1904) {
  if (date1904) v += 1462
  var epoch = Date.parse(v)
  return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000)
}

function sheet_from_array_of_arrays(data, opts) {
  var ws = {}
  var range = {
    s: {
      c: 10000000,
      r: 10000000,
    },
    e: {
      c: 0,
      r: 0,
    },
  }
  for (var R = 0; R != data.length; ++R) {
    for (var C = 0; C != data[R].length; ++C) {
      if (range.s.r > R) range.s.r = R
      if (range.s.c > C) range.s.c = C
      if (range.e.r < R) range.e.r = R
      if (range.e.c < C) range.e.c = C
      var cell = {
        v: data[R][C],
      }
      if (cell.v == null) continue
      var cell_ref = XLSX.utils.encode_cell({
        c: C,
        r: R,
      })

      if (typeof cell.v === 'number') cell.t = 'n'
      else if (typeof cell.v === 'boolean') cell.t = 'b'
      else if (cell.v instanceof Date) {
        cell.t = 'n'
        cell.z = XLSX.SSF._table[14]
        cell.v = datenum(cell.v)
      } else cell.t = 's'

      cell.s = {
        font: {
          name: '宋体',
        },
      }

      ws[cell_ref] = cell
    }
  }
  if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range)
  return ws
}

function Workbook() {
  if (!(this instanceof Workbook)) return new Workbook()
  this.SheetNames = []
  this.Sheets = {}
}

function s2ab(s) {
  var buf = new ArrayBuffer(s.length)
  var view = new Uint8Array(buf)
  for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xff
  return buf
}

function initColumnMap() {
  var time = maxColumn / 26
  for (var t = -1; t < time; t++) {
    for (var i = 0; i < 26; i++) {
      if (columnMap.length < maxColumn) {
        if (t === -1) {
          columnMap.push(charMap[i])
        } else {
          columnMap.push(charMap[t] + charMap[i])
        }
      }
    }
  }
}

function calTitleLevel(obj, fields) {
  var res = 1
  ;(function myDeep(obj, num) {
    if (typeof obj === 'object') {
      for (var key in obj) {
        if (typeof obj[key] === 'object') {
          myDeep(obj[key], num + 1)
        } else {
          res = res < num + 1 ? num + 1 : res
          fields.push(obj[key])
        }
      }
    } else {
      res = res > num ? res : num
      fields.push(obj[key])
    }
  })(obj, 1)
  return res - 1
}

function formatJson(filterVal, jsonData) {
  return jsonData.map(v =>
    filterVal.map(j => {
      if (j === 'timestamp') {
        return parseTime(v[j])
      } else {
        return v[j] ? v[j] : ''
      }
    })
  )
}

function exportBlob(blob, fileName) {
  if (window.navigator && window.navigator.msSaveOrOpenBlob) {
    window.navigator.msSaveOrOpenBlob(blob, fileName)
  } else {
    var url = window.URL.createObjectURL(blob)
    var a = document.createElement('a')
    a.href = url
    a.download = fileName
    a.dispatchEvent(
      new MouseEvent('click', {
        bubbles: true,
        cancelable: true,
        view: window,
      })
    )
    window.URL.revokeObjectURL(blob)
  }
}

function calHeaderChildren(fieldMap, level) {
  var obj = []
  if (level == 0) {
    var isFirst = true
    ;(function myDeep(map) {
      var blankNum = 0
      if (isFirst) {
        for (var key in map) {
          obj.push(key)
          if (typeof map[key] === 'object') {
            isFirst = false
            myDeep(map[key])
            isFirst = true
          }
        }
      } else {
        Object.keys(map).forEach((key, index) => {
          if (typeof map[key] === 'object') {
            myDeep(map[key])
          }
          blankNum++
          if (index == Object.keys(map).length - 1) {
            for (var i = 1; i <= blankNum - 1; i++) {
              obj.push('')
            }
          }
        })
      }
      isFirst = false
    })(fieldMap)
  } else {
    var isFirst = true
    ;(function deep(map, num) {
      var blankNum = 0
      if (isFirst) {
        for (var key in map) {
          if (typeof map[key] === 'object') {
            isFirst = false
            deep(map[key], num + 1)
            isFirst = true
          } else {
            obj.push('')
          }
        }
      } else {
        Object.keys(map).forEach((key, index) => {
          if (typeof map[key] === 'string') {
            level != num && blankNum++
            if (index == Object.keys(map).length - 1) {
              for (
                var i = 1;
                i <= (num > level ? blankNum - 1 : blankNum);
                i++
              ) {
                obj.push('')
              }
            } else if (typeof map[Object.keys(map)[index + 1]] === 'object') {
              for (var i = 1; i <= blankNum; i++) {
                obj.push('')
              }
              blankNum = 0
            }
          }
          if (level == num) {
            obj.push(key)
          }
          if (typeof map[key] === 'object') {
            deep(map[key], num + 1)
          }
        })
      }
      isFirst = false
    })(fieldMap, 0)
  }
  return obj
}

function calMerges(fieldMap, multiHeader, header, level) {
  initColumnMap()
  var merges = []
  for (var i = 0; i < level; i++) {
    ;(function myDeep(map, curLevel, num) {
      Object.keys(map).forEach((key, index) => {
        if (typeof map[key] === 'string') {
          if (curLevel == num && num < multiHeader.length) {
            var inx = multiHeader[num].indexOf(key)
            merges.push(
              columnMap[inx] + (num + 1) + ':' + columnMap[inx] + level
            )
          }
        } else {
          if (curLevel == num) {
            var inx = multiHeader[num].indexOf(key)
            var blankNum = 0
            if (num != multiHeader.length - 1 || multiHeader.length == 1) {
              for (var i = inx + 1; i < multiHeader[num].length; i++) {
                if (multiHeader[num][i] === '') {
                  blankNum++
                } else {
                  break
                }
              }
            } else {
              // for (var i = inx + 1; i < multiHeader[num].length; i++) {
              //   if (multiHeader[num][i] !== '') {
              for (var j = inx + 1; j < header.length; j++) {
                if (header[j]) {
                  blankNum++
                } else {
                  break
                }
              }
              //   } else {
              //     break
              //   }
              // }
            }
            merges.push(
              columnMap[inx] +
                (num + 1) +
                ':' +
                columnMap[inx + blankNum] +
                (num + 1)
            )
          }
          myDeep(map[key], curLevel, num + 1)
        }
      })
    })(fieldMap, i, 0)
  }
  return merges
}

function mergeDataRows(mergeColumns, data, fields, titleLevel) {
  var merges = []
  mergeColumns.forEach((column, index) => {
    const i = fields.indexOf(column) // columnMap[i] 就是第i列
    if (i >= 0) {
      var startIndex = 0
      var endIndex = 0
      data.forEach((row, rowIndex) => {
        if (rowIndex != 0) {
          if (row[column] != data[rowIndex - 1][column]) {
            merges.push(
              columnMap[i] +
                (startIndex + 1 + titleLevel) +
                ':' +
                columnMap[i] +
                (endIndex + 1 + titleLevel)
            )
            startIndex = rowIndex
          } else {
            endIndex = rowIndex
            if (rowIndex == data.length - 1) {
              merges.push(
                columnMap[i] +
                  (startIndex + 1 + titleLevel) +
                  ':' +
                  columnMap[i] +
                  (endIndex + 1 + titleLevel)
              )
            }
          }
        }
      })
    }
  })
  return merges
}

function mergeRowsByRow(mainMergeColumns, data, fields, titleLevel) {
  var merges = []
  Object.keys(mainMergeColumns).forEach(mainColumn => {
    const mergeColumns = mainMergeColumns[mainColumn]
    mergeColumns.forEach((column, index) => {
      const i = fields.indexOf(column) // columnMap[i] 就是第i列
      if (i >= 0) {
        var startIndex = 0
        var endIndex = 0
        data.forEach((row, rowIndex) => {
          if (rowIndex != 0) {
            if (row[mainColumn] != data[rowIndex - 1][mainColumn]) {
              if (endIndex > startIndex) {
                merges.push(
                  columnMap[i] +
                    (startIndex + 1 + titleLevel) +
                    ':' +
                    columnMap[i] +
                    (endIndex + 1 + titleLevel)
                )
              }
              startIndex = rowIndex
            } else {
              endIndex = rowIndex
              if (rowIndex == data.length - 1) {
                merges.push(
                  columnMap[i] +
                    (startIndex + 1 + titleLevel) +
                    ':' +
                    columnMap[i] +
                    (endIndex + 1 + titleLevel)
                )
              }
            }
          }
        })
      }
    })
  })
  return merges
}

function export_table_to_excel(id) {
  var theTable = document.getElementById(id)
  var oo = generateArray(theTable)
  var ranges = oo[1]

  /* original data */
  var data = oo[0]
  var ws_name = 'SheetJS'

  var wb = new Workbook(),
    ws = sheet_from_array_of_arrays(data)

  /* add ranges to worksheet */
  // ws['!cols'] = ['apple', 'banan'];
  ws['!merges'] = ranges

  /* add worksheet to workbook */
  wb.SheetNames.push(ws_name)
  wb.Sheets[ws_name] = ws

  var wbout = XLSX.write(wb, {
    bookType: 'xlsx',
    bookSST: false,
    type: 'binary',
  })

  saveAs(
    new Blob([s2ab(wbout)], {
      type: 'application/octet-stream',
    }),
    'test.xlsx'
  )
}

/**
 *
 * @param {Object} param0 参数：{fieldMap：表头字段映射，sourceData：数据源，filename：文件名，mergeColumns：要自动合并的列名（自动合并连续相同的行数据）}
 * @param {*} callback 回调函数
 */
function export_json_to_excel(
  {
    fieldMap = {},
    sourceData = [],
    filename,
    autoWidth = true,
    bookType = 'xlsx',
    mergeColumns = [],
    mainMergeColumns = {},
  } = {},
  callback
) {
  var fields = []
  var titleLevel = calTitleLevel(fieldMap, fields)

  var multiHeader = []
  var header = []
  for (var i = 0; i < titleLevel; i++) {
    if (i == titleLevel - 1) {
      // headerTmp
      header = calHeaderChildren(fieldMap, i)
    } else {
      // multiHeaderTmp
      multiHeader.push(calHeaderChildren(fieldMap, i))
    }
  }

  var data = formatJson(fields, sourceData)
  var merges = calMerges(fieldMap, multiHeader, header, titleLevel)
  if (mergeColumns && mergeColumns.length > 0) {
    merges = merges.concat(
      mergeDataRows(mergeColumns, sourceData, fields, titleLevel)
    )
  }
  if (mainMergeColumns && Object.keys(mainMergeColumns).length > 0) {
    merges = merges.concat(
      mergeRowsByRow(mainMergeColumns, sourceData, fields, titleLevel)
    )
  }
  /* original data */
  filename = filename || 'excel-list'
  data = [...data]
  data.unshift(header)
  for (let i = multiHeader.length - 1; i > -1; i--) {
    data.unshift(multiHeader[i])
  }

  var ws_name = 'SheetJS'
  var wb = new Workbook(),
    ws = sheet_from_array_of_arrays(data)

  if (merges.length > 0) {
    if (!ws['!merges']) ws['!merges'] = []
    merges.forEach(item => {
      ws['!merges'].push(XLSX.utils.decode_range(item))
    })
  }

  for (var i = 0; i < header.length; i++) {
    for (var j = 1; j <= multiHeader.length + 1; j++) {
      // ws['A1'].s =
      if (ws[columnMap[i] + j]) {
        ws[columnMap[i] + j].s = {
          font: {
            name: '宋体',
            // sz: 24,
            bold: true,
            // color: { rgb: "FFFFAA00" }
          },
          alignment: {
            horizontal: 'center',
            vertical: 'center',
            wrapText: true,
          },
          // fill: { bgcolor: 'rgb(217, 225, 242)' }
          // fill: { bgcolor: { rgba: '217, 225, 242, 0.6' } }
        }
      }
    }
    for (var j = titleLevel + 1; j <= sourceData.length + titleLevel; j++) {
      if (ws[columnMap[i] + j]) {
        ws[columnMap[i] + j].s = {
          font: { name: '宋体' },
          alignment: {
            horizontal: 'center',
            vertical: 'center',
            wrapText: true,
          },
        }
      }
    }
  }

  if (autoWidth) {
    /*设置worksheet每列的最大宽度*/
    const colWidth = data.map(row =>
      row.map(val => {
        /*先判断是否为null/undefined*/
        if (val == null) {
          return {
            wch: 10,
          }
        } else if (val.toString().charCodeAt(0) > 255) {
          /*再判断是否为中文*/
          return {
            wch: val.toString().length * 2,
          }
        } else {
          return {
            wch: val.toString().length,
          }
        }
      })
    )
    /*以第一行为初始值*/
    let result = colWidth[0]
    for (let i = 1; i < colWidth.length; i++) {
      for (let j = 0; j < colWidth[i].length; j++) {
        if (result[j]['wch'] < colWidth[i][j]['wch']) {
          result[j]['wch'] = colWidth[i][j]['wch']
        }
      }
    }
    ws['!cols'] = result
  }

  /* add worksheet to workbook */
  wb.SheetNames.push(ws_name)
  wb.Sheets[ws_name] = ws

  var wbout = XLSX.write(wb, {
    bookType: bookType,
    bookSST: false,
    type: 'binary',
  })
  // saveAs(new Blob([s2ab(wbout)], {
  //   type: "application/octet-stream"
  // }), `${filename}.${bookType}`);
  exportBlob(
    new Blob([s2ab(wbout)], {
      type: 'application/vnd.ms-excel',
    }),
    `${filename}.${bookType}`
  )
  if (callback && typeof callback === 'function') {
    callback()
  }
}

export default { export_json_to_excel }
