package model

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"testing"
)

/**
* @Author: super
* @Date: 2020-08-28 16:40
* @Description:
**/


func BenchmarkCellReplace(b *testing.B) {
	f1, err := excelize.OpenFile("../data/excel1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	f2, err := excelize.OpenFile("../data/excel2.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	source := &Excel{
		File:  f1,
		Sheet: "Sheet1",
		Col:   "A",
		Row:   "1",
	}

	target := &Excel{
		File:  f2,
		Sheet: "Sheet1",
		Col:   "A",
		Row:   "1",
	}

	for i := 0; i<b.N; i++{
		target.CellReplace(source)
	}
	err = target.SaveChange()
	if err != nil{
		fmt.Println(err)
	}
}