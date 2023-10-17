package main

import (
	"encoding/csv"
	"fmt"
	"flag"
	"log"
	"os"
	"strconv"

	"github.com/xuri/excelize/v2"
)

func main() {
	fileName := flag.String("filename","","filename to open")
	sheetNumber := flag.Int("sheetnumber",1,"sheet number to open")
	flag.Parse()
	//open the excel file
	file, err := excelize.OpenFile(*fileName)
	if err != nil {

		log.Fatal(err)
	}
	//read the sheet number that we are going to use
	firstSheet := file.WorkBook.Sheets.Sheet[*sheetNumber-1].Name
	//read the names of the columns needed
	var slColNames []string
	fmt.Println("give me the column names you want to use one by one:")
	for i := 0; i < 2; i++ {
		var str string
		fmt.Scan(&str)
		slColNames = append(slColNames, str)
	}

	// This is done for  specific excel sheet, of course it needs to be abstracter
	CellsNamesE := GetCellNames(file, firstSheet, slColNames[0])
	CellValuesE := GetCellValues(CellsNamesE, firstSheet, file)
	CellsNamesA := GetCellNames(file, firstSheet, slColNames[1])
	CellValuesA := GetCellValues(CellsNamesA, firstSheet, file)
	/***********************************/
	defer file.Close()
        /*********************************/
	// This is done on a specific excel sheet, of course it needs to be abstracter

	var Data [][]string
	var Data1 []string
	var Data2 []string
	Data1 = append(Data1, "Serial Number")
	Data2 = append(Data2, "Asset Tag")
	for  i:=0;i<len(CellValuesA);i++{
		Data2 = append(Data2, CellValuesA[i+1])
		Data1 = append(Data1, CellValuesE[i+1])
	}
	Data = append(Data, Data1)
	Data = append(Data, Data2)
	fmt.Println(Data)
	Data = RotateSlice90(Data)

	// create csv file
	csvFile, err := os.Create("inv3.csv")
	if err != nil {
		log.Fatal(err)
	}
	// create new writer to be able to write on the csv file created
	csvWr := csv.NewWriter(csvFile)
   
	for i := 0; i<len(Data); i++ {
			if(Data[i][0] != "" && Data[i][1] != ""){
			_ = csvWr.Write(Data[i])
		}
	}
	
	csvWr.Flush()
	// close the file
	csvFile.Close()
}

// this function creats a map of the cellnames to be used
func GetCellNames(file *excelize.File, SheetName string, CellLetter string) map[int]string {
	rows, err := file.GetRows(SheetName)
	if err != nil {
		fmt.Println(err)
	}
	MapCellsNames := make(map[int]string)
	for k, _ := range rows {
		k++
		kstr := strconv.Itoa(k)
		cellname := CellLetter + kstr
		MapCellsNames[k] = cellname
	}
	return MapCellsNames
}


// this function creats a map of the cellvalues to be used

func GetCellValues(MapCellsNames map[int]string, SheetName string, file *excelize.File) map[int]string {
	MapCellValues := make(map[int]string)
	var err error
	for i, j := range MapCellsNames {
		if i == 1 {
			continue
		}
		MapCellValues[i], err = file.GetCellValue(SheetName, j)
		if err != nil {
			log.Fatal(err)
		}
	}
	return MapCellValues
}

// transpose a 2s slice
func RotateSlice90(Data [][]string) [][]string {
	NData := make([][]string, len(Data[0]))
	fmt.Println(len(Data))
	for i := 0; i < len(NData); i++ {
		NData[i] = make([]string, 2)
	}
	for i := 0; i < len(NData); i++ {
		NData[i][0] = Data[0][i]
		NData[i][1] = Data[1][i]
	}
	return NData
}
