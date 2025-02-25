package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func main() {
	filename := "./CSF111_202425_01_GradeBook_stripped.xlsx"
	sheetname := "CSF111_202425_01_GradeBook"
	
	f, err := excelize.OpenFile(filename)

	
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// Get value from cell by given worksheet name and cell reference.
	cell, err := f.GetCellValue(sheetname, "B2")
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(cell)
	// Get all the rows in the Sheet1.
	rows, err := f.GetRows(sheetname)
	if err != nil {
		fmt.Println(err)
		return
	}
	for _, row := range rows {
		for _, colCell := range row {
			fmt.Print(colCell, "\t")
		}
		fmt.Println()
	}
}
