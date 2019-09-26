/**
 * iview table export to excel
 * export column with exportable is true
 * ref https://github.com/han6054/export-excel
 */
const tableToexcel = {
    downloadLoading: false,
    convertHeaderArr(subHeaders = [], rowCount = 0, columnCount = 0) {
        let headerArr = []
        subHeaders.forEach(item => {
            if (!item['exportable']) {
                return
            }
            if (!item.children) {
                if (!headerArr[rowCount]) {
                    headerArr[rowCount] = []
                }
                headerArr[rowCount][columnCount] = item
                columnCount++
            } else {
                let row = rowCount
                let column = columnCount
                let subHeaderArr = this.convertHeaderArr(item.children, rowCount + 1, columnCount)
                let subRowCount = 0
                subHeaderArr.forEach((item, index) => {
                    item.forEach((subItem, subIndex) => {
                        if (!headerArr[index]) {
                            headerArr[index] = []
                        }
                        headerArr[index][subIndex] = subItem
                        if (subItem.merge) {
                            columnCount += subItem.merge.e.c - subItem.merge.s.c + 1
                        } else {
                            if (subRowCount < 1) {
                                columnCount++
                            }
                        }
                    })
                    subRowCount++
                })
                if (!headerArr[row]) {
                    headerArr[row] = []
                }
                headerArr[row][column] = item
                if (subHeaderArr && subHeaderArr.length > 0) {
                    if (column !== columnCount - 1 || row !== rowCount) {
                        headerArr[row][column].merge = {s: {c: column, r: row}, e: {c: columnCount - 1, r: rowCount}}
                    }
                }
            }
        })
        return headerArr
    },
    checkAndPushMerge(mergeArr = [], columnIndex = 0, startRow = 0, endRow = 0) {
        let filterMerge = mergeArr.filter(item => {
            let minColumn = Math.min(item.s.c, item.e.c)
            let minRow = Math.min(item.s.r, item.e.r)
            let maxColumn = Math.max(item.s.c, item.e.c)
            let maxRow = Math.max(item.s.r, item.e.r)
            if (columnIndex >= minColumn && columnIndex <= maxColumn) {
                if (startRow <= maxRow && startRow >= minRow) {
                    return true
                }
                if (endRow >= minRow && endRow <= maxRow) {
                    return true
                }
            }
            return false
        })
        if (filterMerge.length < 1) {
            mergeArr.push({s: {c: columnIndex, r: startRow}, e: {c: columnIndex, r: endRow}})
        }
        return mergeArr
    },
    calcMergeArr(titleArr = [], mergeArr = [], maxRow, maxColumn) {
        for (let columnIndex = 0; columnIndex < maxColumn; columnIndex++) {
            let startRow = 0
            let endRow = 0
            for (let rowIndex = 0; rowIndex < maxRow; rowIndex++) {
                if (titleArr[rowIndex] && (!titleArr[rowIndex][columnIndex] || titleArr[rowIndex][columnIndex].isEmpty)) {
                    if (startRow === endRow) {
                        startRow = rowIndex - 1
                    }
                    startRow = Math.min(rowIndex - 1, startRow)
                    endRow = rowIndex
                } else {
                    if (startRow !== endRow) {
                        mergeArr = this.checkAndPushMerge(mergeArr, columnIndex, startRow, endRow)
                    }
                    startRow = 0
                    endRow = 0
                }
            }
            if (startRow !== endRow) {
                mergeArr = this.checkAndPushMerge(mergeArr, columnIndex, startRow, endRow)
            }
        }
        return {header: titleArr, merges: mergeArr}
    },
    convertHeader(target) {
        let headerArr = this.convertHeaderArr(target)
        let titleArr = []
        let mergeArr = []
        let maxRow = 0
        let maxColumn = 0
        for (let index = 0; index < headerArr.length; index++) {
            let item = headerArr[index]
            if (!titleArr[index]) {
                titleArr[index] = []
            }
            if (!item) {
                continue
            }
            maxRow++
            maxColumn = Math.max(maxColumn, item.length)
            for (let subIndex = 0; subIndex < maxColumn; subIndex++) {
                if (!titleArr[index][subIndex]) {
                    titleArr[index][subIndex] = {v: '', isEmpty: true}
                }
                let child = item[subIndex]
                if (child) {
                    titleArr[index][subIndex] = {v: child['title']}
                    if (child.merge) {
                        mergeArr.push(child.merge)
                    }
                }
            }
        }
        return this.calcMergeArr(titleArr, mergeArr, maxRow, maxColumn)
    },
    convertKeys(target = [], arr = []) {
        target.forEach((item) => {
            if (!item['exportable']) {
                return
            }
            if (item.children) {
                arr = this.convertKeys(item.children, arr)
            } else {
                arr.push({v: item['key']})
            }
        })
        return arr
    },
    filterByKeys(filterKeys, jsonData) {
        return jsonData.map(item => filterKeys.map(j => { return { v: item[j.v] } }))
    }
}

function downloadExcel(title, columnsData, tableData) {
    downloadExcelSheet([{columnsData: columnsData, tableData: tableData, title: title}], title)
}

function downloadExcelSheet(sheets = [{columnsData: [], tableData: [], title: String}], defaultFileName = String) {
     tableToexcel.downloadLoading = true
     require.ensure([], () => {
         const exportJsonToExcel = require('../../lib/excel/Export2Excel').export_json_to_excel
        exportJsonToExcel(sheets.map(sheet => {
            let filterVal = tableToexcel.cutValue(sheet.columnsData)
            return {
                title: sheet.title,
                opts: tableToexcel.cutHeader(sheet.columnsData),
                filterVal: filterVal,
                data: tableToexcel.formatJson(filterVal, sheet.tableData)
            }
        }), defaultFileName)
         tableToexcel.downloadLoading = false
     })
}
 
export {
   downloadExcel, downloadExcelSheet
}
