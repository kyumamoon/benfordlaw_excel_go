// create spreadsheet

package main

import (
	"fmt"
	"math"
	"strconv"

	"github.com/xuri/excelize"
)

var runner bool = true
var cellcount int = 1
var digitcount = [9]int{0, 0, 0, 0, 0, 0, 0, 0, 0}
var totalcount int
var filename string = "BenfordLawTest.xlsx"

func sort(num string) {
	x := num

	switch string(x[0]) {
	case "1":
		digitcount[0]++
	case "2":
		digitcount[1]++
	case "3":
		digitcount[2]++
	case "4":
		digitcount[3]++
	case "5":
		digitcount[4]++
	case "6":
		digitcount[5]++
	case "7":
		digitcount[6]++
	case "8":
		digitcount[7]++
	case "9":
		digitcount[8]++
	case "-":
		sort(string(x[1]))
	default:

	}

}

func main() {
	f, err := excelize.OpenFile(filename)

	defer func() {
		if err != nil {
			fmt.Println(err)
			return
		}
	}()

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	defer f.Close()

	/*
		for runner == true {
			cellnumber := "A" + strconv.Itoa(cellcount)
			cell, err := f.GetCellValue("Sheet1", cellnumber)
			if err != nil {
				fmt.Println(err)
				return
			} else {

				fmt.Println(cellnumber)

				if cell == "" {
					//runner = false
					cellcount--
					break
				} else {
					sort(cell)
					cellcount++
				}
			}

		}*/

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}
	for _, row := range rows {
		for _, colCell := range row {
			//fmt.Printf("FIRST LINE: " + colCell)
			//fmt.Printf(colCell, "\t")
			if colCell == "" {
				//runner = false
				//cellcount--
				//break
			} else {
				sort(colCell)
				cellcount++
			}
		}
		//fmt.Println()
	}

	cellcount--

	for i := 0; i < 9; i++ {
		totalcount += digitcount[i]
	}

	var frequency = [9]float64{}

	for i := 0; i < 9; i++ {
		frequency[i] = (float64(digitcount[i]) / float64(totalcount)) //* 100
	}

	/*array := make([]string, cellcount)

	for i := 0; i < cellcount; i++ {
		cellnumber := "A" + strconv.Itoa(i+1)
		array[i], err = f.GetCellValue("Sheet1", cellnumber)
		if err != nil {
			fmt.Println(err)
			return
		}

	}*/

	//fmt.Println(len(array))

	fmt.Printf("Number of entries: %d", cellcount)
	fmt.Printf("\n1: %d\n2: %d\n3: %d\n4: %d\n5: %d\n6: %d\n7: %d\n8: %d\n9: %d\n",
		digitcount[0], digitcount[1], digitcount[2], digitcount[3], digitcount[4], digitcount[5], digitcount[6], digitcount[7], digitcount[8])

	fmt.Printf("Total Count: " + strconv.Itoa(totalcount))

	for i := 1; i < 10; i++ {
		fmt.Printf("\n %d Frequency: %g%%", i, (math.Ceil(frequency[i-1]*10000) / 10000))
	}

	index := f.NewSheet("BENFORDS LAW")

	f.SetCellValue("BENFORDS LAW", "A1", "Digit")
	f.SetCellValue("BENFORDS LAW", "B1", "Analyzed Frequency")
	f.SetCellValue("BENFORDS LAW", "C1", "Benford's Frequency")

	for i := 2; i < 11; i++ {
		x := "A" + strconv.Itoa(i)
		//text := strconv.Itoa(i - 1)
		text := i - 1
		f.SetCellValue("BENFORDS LAW", x, text)
	}

	f.SetCellValue("BENFORDS LAW", "D1", "TOTAL COUNT: ")
	f.SetCellValue("BENFORDS LAW", "E1", totalcount)

	for i := 2; i < 11; i++ {
		x := "B" + strconv.Itoa(i)
		f.SetCellValue("BENFORDS LAW", x, (math.Ceil(frequency[i-2]*10000) / 10000))
	}

	var bendefault = [9]float64{0.3010, 0.1760, 0.1250, 0.0960, 0.0790, 0.0670, 0.0580, 0.0510, 0.0460}

	for i := 2; i < 11; i++ {
		x := "C" + strconv.Itoa(i)
		f.SetCellValue("BENFORDS LAW", x, bendefault[i-2])
	}

	/*var format = Style{
		NumFmt: 10,
	}*/

	newstyle, err := f.NewStyle(&excelize.Style{NumFmt: 10})

	if err != nil {
		fmt.Println("ERROR STYLE")
	}

	/*newstyle2, err := f.NewStyle(&excelize.Alignment{
		Horizontal: "",
	})

	if err != nil {
		fmt.Println(err)
		fmt.Println("ERROR ALIGNMENT")
	}*/
	err = f.SetColStyle("BENFORDS LAW", "B", newstyle)
	err = f.SetColStyle("BENFORDS LAW", "C", newstyle)

	/*err = f.SetColStyle("BENFORDS LAW", "A", newstyle2)
	err = f.SetColStyle("BENFORDS LAW", "B", newstyle2)
	err = f.SetColStyle("BENFORDS LAW", "C", newstyle2)
	err = f.SetColStyle("BENFORDS LAW", "D", newstyle2)*/

	err = f.SetColWidth("BENFORDS LAW", "A", "B", 13)
	err = f.SetColWidth("BENFORDS LAW", "B", "C", 19)
	err = f.SetColWidth("BENFORDS LAW", "C", "D", 19)
	err = f.SetColWidth("BENFORDS LAW", "D", "E", 14)
	err = f.SetColWidth("BENFORDS LAW", "E", "F", 11)

	f.SetActiveSheet(index)

	if err := f.SaveAs(filename); err != nil {
		fmt.Println(err)
	}
}
