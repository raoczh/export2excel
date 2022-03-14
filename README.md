# export2excel
js导出excel封装工具，支持多层表头，已验证单层和双层三层，理论上支持更多。

# 使用方法：
需要引入xlsx和xlsx-style

表头层级根据fieldMap的对象层级决定
```javascript
var fieldMap = {
    '字段1': 'field1', '字段2': 'field2',
    '字段3': {
        '字段4': { '字段5': 'field5', '字段6': 'field6' }, '字段7': 'field7'
    },
    '字段8': {
        '字段9': 'field9', '字段10': 'field10'
    }
}
var tableData=[
    {field1:'aaa', field2:'bbb',field5:'ccc',field6:'ddd',field7:'eee',field9:'fff', field10:'ggg'},
    {field1:'aaa', field2:'bbb',field5:'ccc',field6:'ddd',field7:'eee',field9:'fff', field10:'ggg'},
    {field1:'aaa', field2:'bbb',field5:'ccc',field6:'ddd',field7:'eee',field9:'fff', field10:'ggg'},
    {field1:'aaa', field2:'bbb',field5:'ccc',field6:'ddd',field7:'eee',field9:'fff', field10:'ggg'},
]
var callback = () =>{}

export2Excel && export2Excel.export_json_to_excel({ fieldMap, sourceData: tableData, filename: '导出数据' }, callback)
```


