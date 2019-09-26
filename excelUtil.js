const excelUtil = {
    loading: false,
    excelReader: null,
    fileHandler: [],
    getExcelReader() {
        if (this.excelReader === null) {
            this.excelReader = new FileReader()
            this.excelReader.onload = (e) => {
                this.fileHandler = e.target.result
            }
        }
        return this.excelReader
    },
    readAsBinaryString(file, callback = () => {}) {
        this.getExcelReader().readAsBinaryString(file)
        setTimeout(() => {
            if (this.fileHandler) {
                callback()
            }
        }, 500)
    }
}

function excel2Json(file, callback = () => {}) {
    excelUtil.loading = true
    excelUtil.readAsBinaryString(file, () => {
        const jsonFromExecl = require('./Excel2Json').jsonFromExecl
        let jsonData = jsonFromExecl(excelUtil.fileHandler, {type: 'binary'})
        callback(jsonData)
        excelUtil.loading = false
    })
}

export {
    excel2Json
}
