package main

import (
	"Go-excel/model"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
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
		Col:   "A",
		Row:   "1",
	}

	target := &model.Excel{
		File:  f2,
		Sheet: "Sheet1",
		Col:   "A",
		Row:   "1",
	}

	target.CellReplace(source)
	err = target.SaveChange()
	if err != nil{
		fmt.Println(err)
	}
}
