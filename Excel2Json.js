require('file-saver');
require('./Blob');
const XLSX = require('xlsx/dist/xlsx.core.min');

function readWorkbookFile(data, opts){
    return XLSX.read(data, opts)
}

function jsonFromWorksheet(worksheet){
    return XLSX.utils.sheet_to_json(worksheet)
}

export function jsonFromExecl(data, opts){
    let workbook = readWorkbookFile(data, opts)
    let result = {}
    for(let sheetIndex in workbook.Sheets){
        result[sheetIndex] = jsonFromWorksheet(workbook.Sheets[sheetIndex])
    }
    return result
}
