/**
 * iview table export to excel, so far only support 1 level children header
 * export column with exportable is true
 * ref https://github.com/han6054/export-excel
 */
const tableToexcel = {
    downloadLoading: false,
    formatJson(filterVal, jsonData) {
        return jsonData.map(item => filterVal.map(j => { return { v: item[j.v] } }))
    },
    cutHeader(target) {
        let opts = {header: [[]], merges: []}
        let hasChildren = 0
        let childrenLengthArr = []
        let skipCount = 0
        let maxIndex = -1
        target.forEach((item, index) => {
            if (!item['exportable']) {
                skipCount++
                return
            }
            maxIndex++
            maxIndex -= skipCount
            opts.header[0].push({v: item['title']})
            if (item.children) {
                hasChildren++
                childrenLengthArr[hasChildren] = item.children.filter(child => child['exportable']).length
                if (childrenLengthArr[hasChildren - 1]) {
                    maxIndex += childrenLengthArr[hasChildren - 1] - 1
                }
                opts.merges.push({s: {c: maxIndex, r: 0}, e: {c: maxIndex + childrenLengthArr[hasChildren] - 1, r: 0}})
                if (!opts.header[1] || opts.header[1].length < 1) {
                    opts.header[1] = []
                }
                if (hasChildren < 2) {
                    for (let i = 0; i < index; i++) {
                        opts.merges.push({s: {c: i, r: 0}, e: {c: i, r: hasChildren}})
                        opts.header[1].push({v: ''})
                    }
                }
                item.children.forEach((child, i) => {
                    if (child['exportable'] && child['title']) {
                        opts.header[1].push({v: child['title']})
                        if (opts.header[1].length > opts.header[0].length) {
                            opts.header[0].push({v: ''})
                        }
                    }
                })
            } else if (hasChildren > 0) {
                opts.merges.push({s: {c: maxIndex + childrenLengthArr[hasChildren] - 1, r: 0}, e: {c: maxIndex + childrenLengthArr[hasChildren] - 1, r: 1}})
            }
        })
        return opts
    },
    cutValue(target) {
        let arr = []
        target.forEach((item) => {
            if (!item['exportable']) {
                return
            }
            if (item.children) {
                item.children.forEach(child => {
                    if (child['exportable']) {
                        arr.push({v: child['key']})
                    }
                })
            } else {
                arr.push({v: item['key']})
            }
        })
        return arr
    }
}

function downloadExcel(title, columnsData, tableData) {
    tableToexcel.downloadLoading = true
    require.ensure([], () => {
        const opts = tableToexcel.cutHeader(columnsData)
        const filterVal = tableToexcel.cutValue(columnsData)
        const data = tableToexcel.formatJson(filterVal, tableData)
        const exportJsonToExcel = require('./Export2Excel').export_json_to_excel
        exportJsonToExcel(opts.header, data, title, opts)
        tableToexcel.downloadLoading = false
    })
}

export {
    downloadExcel
}
