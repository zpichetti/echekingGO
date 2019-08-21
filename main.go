package main

import (
	"fmt"
	"io/ioutil"
	"strconv"
	"strings"

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
	// list xlsx files
	files, err := ioutil.ReadDir("./")
	if err != nil {
		fmt.Println(err)
	}
	listFiles := []string{}
	for _, f := range files {
		if strings.HasSuffix(f.Name(), ".xlsx") {
			listFiles = append(listFiles, f.Name())
		}
	}
	fmt.Printf("Comparaison des fichiers %s avec le fichier %s \n", listFiles[0], listFiles[1])
	fmt.Println("Patienter...")

	// List Old contract list
	sheet := "Feuil1"
	f1 := getListID(fmt.Sprintf("./%s", listFiles[0]), sheet)

	// extract new contract
	f2, err := excelize.OpenFile(fmt.Sprintf("./%s", listFiles[1]))
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
	// Set first row value.
	f.SetCellValue("Sheet1", "A1", "Code chantier")
	f.SetCellValue("Sheet1", "B1", "Libellé du chantier")
	f.SetCellValue("Sheet1", "C1", "N° Marché")
	f.SetCellValue("Sheet1", "D1", "Désignation du lot")
	f.SetCellValue("Sheet1", "E1", "Code ST")
	f.SetCellValue("Sheet1", "F1", "Nom du ST")
	f.SetCellValue("Sheet1", "G1", "Date de signature")
	f.SetCellValue("Sheet1", "H1", "Date envoi agrément")
	f.SetCellValue("Sheet1", "I1", "Statut M.")

	// Set value row of new contract

	for i, row := range list {
		i1, _ := strconv.ParseInt(row[6], 10, 64)
		i2, _ := strconv.ParseInt(row[7], 10, 64)

		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i+2), row[0])
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", i+2), row[1])
		f.SetCellValue("Sheet1", fmt.Sprintf("C%d", i+2), row[2])
		f.SetCellValue("Sheet1", fmt.Sprintf("D%d", i+2), row[3])
		f.SetCellValue("Sheet1", fmt.Sprintf("E%d", i+2), row[4])
		f.SetCellValue("Sheet1", fmt.Sprintf("F%d", i+2), row[5])
		if i1 == 0 {
			f.SetCellValue("Sheet1", fmt.Sprintf("G%d", i+2), "")
		} else {
			f.SetCellValue("Sheet1", fmt.Sprintf("G%d", i+2), i1)
		}
		if i2 == 0 {
			f.SetCellValue("Sheet1", fmt.Sprintf("H%d", i+2), "")
		} else {
			f.SetCellValue("Sheet1", fmt.Sprintf("H%d", i+2), i2)
		}
		f.SetCellValue("Sheet1", fmt.Sprintf("I%d", i+2), row[8])
	}

	// set time format
	f.SetColWidth("Sheet1", "A", "H", 13)
	f.SetColWidth("Sheet1", "D", "D", 26)
	f.SetColWidth("Sheet1", "B", "B", 26)
	f.SetColWidth("Sheet1", "F", "F", 26)
	style, err := f.NewStyle(`{"number_format": 14}`)
	if err != nil {
		fmt.Println(err)
	}
	err = f.SetCellStyle("Sheet1", "G2", fmt.Sprintf("H%d", len(list)+1), style)
	// Save xlsx file by the given path.
	err2 := f.SaveAs("./comparaison.xlsx")
	if err2 != nil {
		fmt.Println(err2)
	}
	var sizeList int
	fmt.Println("Done !!!")
	fmt.Printf("Il y a %d nouveaux marchés", len(list))
	fmt.Scanln(&sizeList)
}
