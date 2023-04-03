package main

import (
    "fmt"
    "log"
	"strconv"
	"os"
	"encoding/csv"
    "github.com/xuri/excelize/v2"

	
)
const Let = "E"
var cellname string 

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
fmt.Println("give me the column names you want to use one by one  :\t")
for i:= 0 ; i<k ;i++{
	var str string
	fmt.Scan(&str)
	slColNames = append(slColNames,str)
}	
/***************************************************************************************/    

    CellsNamesE:= GetCellNames(file,firstSheet,"E")
    CellValuesE:= GetCellValues(CellsNamesE,firstSheet,file)
	CellsNamesA:= GetCellNames(file,firstSheet,"A")
    CellValuesA:= GetCellValues(CellsNamesA,firstSheet,file)
	
	
/***********************************/
defer file.Close()
/***********************************/
  

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
	csvFile, err:= os.Create("inv.csv")
	if err != nil {
		log.Fatal(err)
	}
  
	csvWr:= csv.NewWriter(csvFile)

	for _, value := range Data {
		_ = csvWr.Write(value)

	}
	csvWr.Flush()
	csvFile.Close()
	/**********************************/
}

func GetCellNames(file *excelize.File,SheetName string,CellLetter string) (map [int]string){
	rows, err := file.GetRows(SheetName)
	if err != nil{
		fmt.Println(err)
	} 
	MapCellsNames:= make(map[int]string)
	for k, _ := range rows{
		k++
		kstr := strconv.Itoa(k)
		cellname = CellLetter + kstr
		MapCellsNames[k]= cellname
	}
	return MapCellsNames
}
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
