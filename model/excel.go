package model

import (
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
	"strings"
)

/**
* @Author: super
* @Date: 2020-08-28 16:09
* @Description:
**/

//将要读取的内容预定义在结构体内
type Excel struct {
	File  *excelize.File
	Sheet string //哪个sheet
	Col   string //第几列
	Row   string //第几行
}

//获取指定单元格的值
func (excel *Excel) GetCellValue() string {
	var builder strings.Builder
	builder.WriteString(excel.Col)
	builder.WriteString(excel.Row)
	value := excel.File.GetCellValue(excel.Sheet, builder.String())
	return value
}

//获取指定单元格的值
func (excel *Excel) GetCellValueWithPosition(col, row string) string {
	var builder strings.Builder
	builder.WriteString(col)
	builder.WriteString(row)
	value := excel.File.GetCellValue(excel.Sheet, builder.String())
	return value
}

//设置指定单元格的值
func (excel *Excel) SetCellValue(value string) {
	var builder strings.Builder
	builder.WriteString(excel.Col)
	builder.WriteString(excel.Row)
	excel.File.SetCellValue(excel.Sheet, builder.String(), value)
}



//替换单元格的值
func (excel *Excel) CellReplace(source *Excel) {
	value := source.GetCellValue()
	excel.SetCellValue(value)
}

//保存修改结果，操作较为耗时
//不保存修改会修改失败
func (excel *Excel) SaveChange() error {
	return excel.File.Save()
}

//获取excel有多少行
func (excel *Excel) Length() int {
	rows := excel.File.GetRows(excel.Sheet)
	return len(rows)
}

//获取source和target之间的唯一标识
//start代表从第几行开始
func (excel Excel) GetIds(start int) map[string]int {
	l := excel.Length()
	result := make(map[string]int)
	for ; start < l; start++ {
		excel.Col = strconv.Itoa(start)
		value := excel.GetCellValue()
		result[value] = start
	}
	return result
}
