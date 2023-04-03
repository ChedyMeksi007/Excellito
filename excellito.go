package main

import (
    "os"
	"fmt"
    "log"
	"strconv"
	"encoding/csv"
    "github.com/xuri/excelize/v2"
)

func main(){
/***************************************************************************************/
// read the file name that you want to read
	fmt.Println("give me the file name you want to open :\t")
	var filename string
  	fmt.Scan(&filename)
//open the excel file
	file, err := excelize.OpenFile(filename)
    if err != nil {
	
        log.Fatal(err)
    }
//read the sheet number that we are going to use
	fmt.Println("give me the sheet number you want to open (1-> ...) :\t")
	var sheetNumber int 
	fmt.Scan(&sheetNumber)
	firstSheet := file.WorkBook.Sheets.Sheet[sheetNumber-1].Name
//read the numbers of the columns needed
var  k int 
for{
fmt.Println("give me the number of columns  you want to use  :\t")
fmt.Scan(&k)
if k <= 26 {
	break;
 }
}
//read the names of the columns needed
var slColNames []string
fmt.Println("give me the column names you want to use one by one:")
for i:= 0 ; i<k ;i++{
	var str string
	fmt.Scan(&str)
	slColNames = append(slColNames,str)
}	

/***************************************************************************************/    
// This is done for  specific excel sheet, of course it needs to be abstracter
    CellsNamesE:= GetCellNames(file,firstSheet,slColNames[0])
    CellValuesE:= GetCellValues(CellsNamesE,firstSheet,file)
	CellsNamesA:= GetCellNames(file,firstSheet,slColNames[1])
    CellValuesA:= GetCellValues(CellsNamesA,firstSheet,file)
	
	
/***********************************/
defer file.Close()
/**********************************/
  // This is done on a specific excel sheet, of course it needs to be abstracter

    var Data [][]string
	var Data1 []string
	var Data2 []string
	    Data1 = append(Data1, "SerialNumber")
		Data2 = append(Data2,"Asset Tag")
    for  i:=1 ; i<=110 ;i++ {
		Data1 = append(Data1, CellValuesA[i+2])
		Data2 = append(Data2, CellValuesE[i+2])	
	}
	Data = append(Data, Data1)
	Data = append(Data, Data2)
	Data = RotateSlice90(Data)

	// create csv file 
	csvFile, err:= os.Create("inv.csv")
	if err != nil {
		log.Fatal(err)
	}
  // create new writer to be able to write on the csv file created
	csvWr:= csv.NewWriter(csvFile)

	for _, value := range Data {
		_ = csvWr.Write(value)

	}
	csvWr.Flush()
	// close the file
	csvFile.Close()
	/**********************************/
}
/*
*@brief this function creats a map of the cellnames to be used  
*
*/
func GetCellNames(file *excelize.File,SheetName string,CellLetter string) (map [int]string){
	rows, err := file.GetRows(SheetName)
	if err != nil{
		fmt.Println(err)
	} 
	MapCellsNames:= make(map[int]string)
	for k, _ := range rows{
		k++
		kstr := strconv.Itoa(k)
		cellname := CellLetter + kstr
		MapCellsNames[k]= cellname
	}
	return MapCellsNames
}
/*
*@brief this function creats a map of the cellvalues to be used 
*/
func GetCellValues(MapCellsNames map [int]string, SheetName string, file *excelize.File ) (map [int]string){
	MapCellValues :=make(map[int]string)
	var err error
	for i, j := range MapCellsNames{
		if i == 1  {continue} 
		MapCellValues[i],err = file.GetCellValue(SheetName, j)
		if err != nil{
			log.Fatal(err)
		}
       }
	return MapCellValues
}
/*
*@brief this function rotates a 2d slice 90° to the right so if yo have a [2][10] slice the output will be a [10][2] slice  
*
*@param Data is the inputed slice 
*
*@param NData is the slice to be returned by this function after the changes that need to be done 
*/
func RotateSlice90(Data [][]string)[][]string{
	NData := make ([][]string,111)
	for i:= 0 ; i< len(NData); i++{
		NData[i] = make([]string,2)
	}
	for i := 0  ; i< len(NData) ; i++{
	NData[i][0] = Data[0][i]
	NData[i][1] = Data [1][i]
	}
 return NData
}
