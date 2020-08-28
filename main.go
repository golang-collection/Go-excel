package main

import (
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
	studentInfo, err := excelize.OpenFile("11.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	stu2018, err := excelize.OpenFile("22.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	studentRow := studentInfo.GetRows("Sheet0")
	stu2018Row := stu2018.GetRows("Sheet1")
	studentLen := len(studentRow)
	stu2018Len := len(stu2018Row)

	m := make(map[string]int)

	for i := 2; i <= studentLen; i++ {
		cell := studentInfo.GetCellValue("Sheet0", "B"+strconv.Itoa(i))
		m[cell] = i
	}

	for i := 2; i <= stu2018Len; i++ {
		// Get value from cell by given worksheet name and axis.
		cell := stu2018.GetCellValue("Sheet1", "H"+strconv.Itoa(i))
		if v, ok := m[cell]; ok {
			part := stu2018.GetCellValue("Sheet1", "F"+strconv.Itoa(i))
			studentInfo.SetCellValue("Sheet0", "O" + strconv.Itoa(v), part)

			signDate := stu2018.GetCellValue("Sheet1", "K"+strconv.Itoa(i))
			studentInfo.SetCellValue("Sheet0", "P" + strconv.Itoa(v), signDate)


			pushDate := stu2018.GetCellValue("Sheet1", "L"+strconv.Itoa(i))
			studentInfo.SetCellValue("Sheet0", "Q" + strconv.Itoa(v), pushDate)

			var pushDate2 string
			if pushDate == ""{
				pushDate2 = stu2018.GetCellValue("Sheet1", "M"+strconv.Itoa(i))
				studentInfo.SetCellValue("Sheet0", "Q" + strconv.Itoa(v), pushDate2)
			}

			if pushDate == "" && pushDate2 == ""{
				pushDate2 := stu2018.GetCellValue("Sheet1", "N"+strconv.Itoa(i))
				studentInfo.SetCellValue("Sheet0", "Q" + strconv.Itoa(v), pushDate2)
			}

			activateDate := stu2018.GetCellValue("Sheet1", "P"+strconv.Itoa(i))
			studentInfo.SetCellValue("Sheet0", "R" + strconv.Itoa(v), activateDate)

			activateTrainDate := stu2018.GetCellValue("Sheet1", "Q"+strconv.Itoa(i))
			studentInfo.SetCellValue("Sheet0", "T" + strconv.Itoa(v), activateTrainDate)

			activateFinish := stu2018.GetCellValue("Sheet1", "R"+strconv.Itoa(i))
			studentInfo.SetCellValue("Sheet0", "U" + strconv.Itoa(v), activateFinish)
		}
	}
	err = studentInfo.Save()
	if err != nil{
		fmt.Println(err)
	}
	fmt.Println("finish")
}