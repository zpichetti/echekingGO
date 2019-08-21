package main

import (
	"fmt"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	f1, err := excelize.OpenFile("./doc1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	// Get all the rows in the Sheet1.
	rows, err := f1.GetRows("Sheet1")
	for _, row := range rows {
		for _, colCell := range row {
			fmt.Print(colCell, "\t")
		}
		fmt.Println()
	}

	// write new file with result
	f := excelize.NewFile()
	// Set value of a cell.
	f.SetCellValue("Sheet1", "B2", 300)
	// Save xlsx file by the given path.
	err2 := f.SaveAs("./Result.xlsx")
	if err2 != nil {
		fmt.Println(err2)
	}
}
