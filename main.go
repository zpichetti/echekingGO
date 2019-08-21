package main

import (
	"fmt"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func contains(a []string, x string) bool {
	for _, n := range a {
		if x == n {
			return true
		}
	}
	return false
}

func getListID(file string, sheet string) (list []string) {
	f, err := excelize.OpenFile(file)
	if err != nil {
		fmt.Println(err)
		return
	}
	rows, err := f.GetRows(sheet)
	for _, row := range rows {
		list = append(list, row[2])
	}
	return
}

func main() {
	sheet := "Feuil1"
	f1 := getListID("./ancien.xlsx", sheet)

	f2, err := excelize.OpenFile("./nouveau.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	rowsf2, err := f2.GetRows(sheet)
	list := [][]string{}
	for _, row := range rowsf2 {
		if contains(f1, row[2]) == false {
			list = append(list, row)
		}
	}
	// fmt.Printf("%v\n", list)

	// write new file with result
	f := excelize.NewFile()
	// Set value of a cell.
	f.SetCellValue("Sheet1", "A1", "Code chantier")
	f.SetCellValue("Sheet1", "B1", "Libellé du chantier")
	f.SetCellValue("Sheet1", "C1", "N° Marché")
	f.SetCellValue("Sheet1", "D1", "Désignation du lot")
	f.SetCellValue("Sheet1", "E1", "Code ST")
	f.SetCellValue("Sheet1", "F1", "Nom du ST")
	f.SetCellValue("Sheet1", "G1", "Date de signature")
	f.SetCellValue("Sheet1", "H1", "Date envoi agrément")
	f.SetCellValue("Sheet1", "I1", "Statut M.")

	for i, row := range list {
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i+2), row[0])
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", i+2), row[1])
		f.SetCellValue("Sheet1", fmt.Sprintf("C%d", i+2), row[2])
		f.SetCellValue("Sheet1", fmt.Sprintf("D%d", i+2), row[3])
		f.SetCellValue("Sheet1", fmt.Sprintf("E%d", i+2), row[4])
		f.SetCellValue("Sheet1", fmt.Sprintf("F%d", i+2), row[5])
		f.SetCellValue("Sheet1", fmt.Sprintf("G%d", i+2), row[6])
		f.SetCellValue("Sheet1", fmt.Sprintf("H%d", i+2), row[7])
		f.SetCellValue("Sheet1", fmt.Sprintf("I%d", i+2), row[8])
	}
	// Save xlsx file by the given path.
	err2 := f.SaveAs("./nouveau_contrat.xlsx")
	if err2 != nil {
		fmt.Println(err2)
	}
	var sizeList int
	fmt.Println("Done !!!")
	fmt.Printf("Il y a %d nouveau(x) contrat(s)", len(list))
	fmt.Scanln(&sizeList)
}
