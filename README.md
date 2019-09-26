在[原项目](https://github.com/han6054/export-excel)上增加
1. 子表头导出支持, 并支持自动合并单元格
2. 文件使用相对路径, 方便直接复制及引用
3. 多sheet表格导出

> 引言：

最近使用vue在做一个后台系统，技术栈 `vue + iView`，在页面中生成表格后，
iView可以实现表格的导出，不过只能导出csv格式的，并不适合项目需求。

### 如果想要导出Excel
- 复制代码到本地
- `npm install -S file-saver` 用来生成文件的web应用程序
- `npm install -S xlsx` 电子表格格式的解析器
#### 表结构
- 渲染页面中的表结构是由`columns`列 和`tableData` 行，来渲染的 `columns`为表头数据`tableData`为表实体内容
```
columns1: [
    {
        title: '序号',
        exportable: true,
        key: 'serNum'
    },
    {
        title: '总表头1',
        key: 'column1',
        children:[{
            align: 'center',
            exportable: true,
            title: '分表头1',
            key: 'subColumn1'
        }]
    }
    ...
],
```
tableData 就不写了，具体数据结构查看[iViewAPI](https://www.iviewui.com/components/table)

在build 目录下`webpack.base.conf.js`配置 我们要加载时的路径
```
 alias: {
      'src': path.resolve(__dirname, '../src'),
    }
```
在当前页面中引入依赖
```
import { downloadExcel } from './table2excel.js'
```

当我们要导出表格执行`@click`事件调用`handleDownload`函数
```
handleDownload() {
      downloadExcel('表名', this.columnsData, this.tableData)
}
