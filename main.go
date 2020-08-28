package main

import (
	"Go-excel/model"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
)

/**
* @Author: super
* @Date: 2020-08-28 12:06
* @Description:
**/

func main() {
	f1, err := excelize.OpenFile("./data/excel1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	f2, err := excelize.OpenFile("./data/excel2.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	source := &model.Excel{
		File:  f1,
		Sheet: "Sheet1",
		Col:   "B",
		Row:   "1",
	}

	target := &model.Excel{
		File:  f2,
		Sheet: "Sheet1",
		Col:   "B",
		Row:   "2",
	}

	ids := target.GetIds(2)
	target.Col = "D"
	length := source.Length()
	for i := 2; i < length; i++ {
		row := strconv.Itoa(i)
		source.Row = row
		value := source.GetCellValue()
		if v, ok := ids[value]; ok {
			phone := source.GetCellValueWithPosition("C", row)
			target.Row = strconv.Itoa(v)
			target.SetCellValue(phone)
		}
	}
	err = target.SaveChange()
	if err != nil{
		fmt.Println(err)
	}
}
