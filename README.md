# xlsx-utils-go
xlsx常用模式function封装 -- go版本

### 安装

go get github.com/xuexin520/xlsx-utils-go

### 使用

```go
    c.Header("Content-Type", "application/octet-stream")
    c.Header("Content-Disposition", "attachment; filename="+"aaa.xlsx")
    c.Header("Content-Transfer-Encoding", "binary")

    回写到web形成下载
    file.Write(c.Writer)

    保存文件方式
    file.SaveAs("./aaa.xlsx")
```

```go
note
    sheetName 	文件sheet名称
    titles 	第一行 标题
    data 		要插入文档的数据 每格值为string

func CreateFileBySliceString(sheetName string, titles []string, data [][]string) (*xlsx.File, error)
```

```go
note
    sheetName 	文件sheet名称 默认值 sheet1
    titles 		第一行标题    示例 ["姓名", "年龄", "电话"]
    dataMapKeys  keys值为data的map-key (作用从map中取key的值插入对应title值下  title值和mapKeys值对应关系)  示例 ["name", "age", "phone"]
    data 		要插入文档的数据 每格值为map (会将map的value值写入表格) 示例 [{"name"="miss", "age"=18, "phone"="12345678901"}]

func CreateFileBySliceMap(sheetName string, titles, dataMapKeys []string, data []map[string]interface{}) (*xlsx.File, error)
```

```go
note
    属于 CreateFileBySliceMap 的byte调用  data接收byte 便于外部调用

func CreateFileBySliceMapByte(sheetName string, titles, dataMapKeys []string, data []byte) (*xlsx.File, error)
```
