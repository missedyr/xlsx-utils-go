package xlsxGo

import (
	"encoding/json"
	"github.com/sirupsen/logrus"
	"github.com/spf13/cast"
	"github.com/tealeg/xlsx"
)

//  c.Header("Content-Type", "application/octet-stream")
//	c.Header("Content-Disposition", "attachment; filename="+"aaa.xlsx")
//	c.Header("Content-Transfer-Encoding", "binary")
//
//	回写到web形成下载
//	file.Write(c.Writer)
//
//  保存文件方式
//  file.SaveAs("./aaa.xlsx")


// CreateFileBySliceString
// sheetName 	文件sheet名称
// titles 		第一行 标题
// data 		要插入文档的数据 每格值为string
func CreateFileBySliceString(sheetName string, titles []string, data [][]string) (*xlsx.File, error) {
	file := xlsx.NewFile()
	if len(sheetName) == 0 {
		sheetName = "sheet1"
	}
	sheet, err := file.AddSheet(sheetName)
	if err != nil {
		return file, err
	}
	insertRows(sheet, [][]string{titles})
	insertRows(sheet, data)
	return file, nil
}

// CreateFileBySliceMapByte
// 属于 CreateFileBySliceMap 的byte调用  data接收byte 便于外部调用
func CreateFileBySliceMapByte(sheetName string, titles, dataMapKeys []string, data []byte) (*xlsx.File, error) {
	var dataNew []map[string]interface{}
	json.Unmarshal(data, &dataNew)
	return CreateFileBySliceMap(sheetName, titles, dataMapKeys, dataNew)
}

// CreateFileBySliceMap
// sheetName 	文件sheet名称 默认值 sheet1
// titles 		第一行标题    示例 ["姓名", "年龄", "电话"]
// dataMapKeys  keys值为data的map-key (作用从map中取key的值插入对应title值下  title值和mapKeys值对应关系)  示例 ["name", "age", "phone"]
// data 		要插入文档的数据 每格值为map (会将map的value值写入表格) 示例 [{"name"="miss", "age"=18, "phone"="12345678901"}]
func CreateFileBySliceMap(sheetName string, titles, dataMapKeys []string, data []map[string]interface{}) (*xlsx.File, error) {
	file := xlsx.NewFile()
	if len(sheetName) == 0 {
		sheetName = "sheet1"
	}
	sheet, err := file.AddSheet(sheetName)
	if err != nil {
		return file, err
	}
	insertRows(sheet, [][]string{titles})

	if len(dataMapKeys) == 0 {
		logrus.Warnf("xlsx-->CreateFileBySliceMap dataMapKeys is empty")
		return file, err
	}

	rowDatas := [][]string{}
	for _, mapD := range data {
		if len(mapD) == 0 {
			continue
		}
		rowDsNew := []string{}
		for _, mapKey := range dataMapKeys {
			rowDsNew = append(rowDsNew, cast.ToString(mapD[mapKey]))
		}
		if len(rowDsNew) == 0 {
			continue
		}
		rowDatas = append(rowDatas, rowDsNew)
	}
	insertRows(sheet, rowDatas)
	return file, nil
}

func insertRows(sheet *xlsx.Sheet, data [][]string) {
	if len(data) == 0 {
		logrus.Warnf("xlsx-->insertRows insert fike rows data empty")
	}
	for _, rowStrArr := range data {
		row := sheet.AddRow()
		for _, str := range rowStrArr {
			cell := row.AddCell()
			cell.SetString(str)
		}
	}
}
