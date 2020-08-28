package model

import (
	"github.com/360EntSecGroup-Skylar/excelize"
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

func (excel *Excel) GetCellValue() string {
	var builder strings.Builder
	builder.WriteString(excel.Col)
	builder.WriteString(excel.Row)
	value := excel.File.GetCellValue(excel.Sheet, builder.String())
	return value
}

func (excel *Excel) SetCellValue(value string) {
	var builder strings.Builder
	builder.WriteString(excel.Col)
	builder.WriteString(excel.Row)
	excel.File.SetCellValue(excel.Sheet, builder.String(), value)
}

func (excel *Excel) CellReplace(source *Excel){
	value := source.GetCellValue()
	excel.SetCellValue(value)
}

//保存较为耗时
func (excel *Excel) SaveChange() error {
	return excel.File.Save()
}
