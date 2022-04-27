package main

// import fyne
import (
	"fmt"
	"io/ioutil"
	"math"
	"strconv"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/storage"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
)

var filename string
var filepath string

var runner bool = true
var cellcount int = 1
var digitcount = [9]int{0, 0, 0, 0, 0, 0, 0, 0, 0}
var totalcount int

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

func benfordlawanalysis() bool {
	f, err := excelize.OpenFile(filepath)

	defer func() {
		if err != nil {
			fmt.Println(err)
			return
		}
	}()

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)

	}
	for _, row := range rows {
		for _, colCell := range row {

			if colCell == "" {

			} else {
				sort(colCell)
				cellcount++
			}
		}

	}

	cellcount--

	for i := 0; i < 9; i++ {
		totalcount += digitcount[i]
	}

	var frequency = [9]float64{}

	for i := 0; i < 9; i++ {
		frequency[i] = (float64(digitcount[i]) / float64(totalcount))
	}

	for i := 1; i < 10; i++ {
		fmt.Printf("\n %d Frequency: %g%%", i, (math.Ceil(frequency[i-1]*10000) / 10000))
	}

	index := f.NewSheet("BENFORDS LAW")

	f.SetCellValue("BENFORDS LAW", "A1", "Digit")
	f.SetCellValue("BENFORDS LAW", "B1", "Analyzed Frequency")
	f.SetCellValue("BENFORDS LAW", "C1", "Benford's Frequency")

	for i := 2; i < 11; i++ {
		x := "A" + strconv.Itoa(i)

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

	newstyle, err := f.NewStyle(&excelize.Style{NumFmt: 10})

	if err != nil {
		fmt.Println("ERROR STYLE")
	}

	err = f.SetColStyle("BENFORDS LAW", "B", newstyle)
	err = f.SetColStyle("BENFORDS LAW", "C", newstyle)

	err = f.SetColWidth("BENFORDS LAW", "A", "B", 13)
	err = f.SetColWidth("BENFORDS LAW", "B", "C", 19)
	err = f.SetColWidth("BENFORDS LAW", "C", "D", 19)
	err = f.SetColWidth("BENFORDS LAW", "D", "E", 14)
	err = f.SetColWidth("BENFORDS LAW", "E", "F", 11)

	f.SetActiveSheet(index)

	if err := f.SaveAs(filepath); err != nil {
		fmt.Println(err)
	}

	return true
}

func main() {
	// New app
	myApp := app.New()
	//New title and window
	myWindow := myApp.NewWindow("Benford's Law")
	// resize window
	myWindow.Resize(fyne.NewSize(500, 50))

	// New Buttton
	btn := widget.NewButton("Choose Excel File", func() {
		// Using dialogs to open files
		// first argument func(fyne.URIReadCloser, error)
		// 2nd is parent window in our case "w"
		// r for reader
		// _ is ignore error
		file_Dialog := dialog.NewFileOpen(
			func(r fyne.URIReadCloser, _ error) {
				// read files
				data, _ := ioutil.ReadAll(r)
				// reader will read file and store data
				// now result
				fmt.Println(r.URI().Name())
				fmt.Println(string(r.URI().Name()[0]))
				//x := strings.TrimRight(r.URI().Name(), ".xlsx")
				//filename = x
				filename = r.URI().Name()
				filepath = r.URI().Path()
				fmt.Println(filename, data)

				x := benfordlawanalysis()
				if x == true {
					myApp.Quit()
				}
			}, myWindow)
		// fiter to open .txt files only
		// array/slice of strings/extensions
		file_Dialog.SetFilter(
			storage.NewExtensionFileFilter([]string{".xlsx"}))
		file_Dialog.Show()
		// Show file selection dialog.
	})
	// lets show button in parent window

	myWindow.SetContent(container.NewVBox(
		btn,
	))
	myWindow.ShowAndRun()
}
