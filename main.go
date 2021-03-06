package main

import (
	"bufio"
	"encoding/csv"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/tealeg/xlsx"
)

func main() {
	if len(os.Args) != 2 {
		failOnError("main.exeにテキストファイルをドラッグ&ドロップしてください", nil)
	}
	//convertTxt2Csv("./testdata/sample_tmp.txt")
	convertTxt2Csv(os.Args[1])
	waitEnter()
}

func failOnError(errMsg string, err error) {
	//errs := errors.WithStack(err)
	fmt.Println(errMsg)
	if err != nil {
		//fmt.Printf("%+v\n", errs) Stack trace
		fmt.Printf("%s\n", err.Error())
	}
	waitEnter()
	os.Exit(1)
}

func waitEnter() {
	fmt.Println("エンターを押すと処理を終了します。")
	scanner := bufio.NewScanner(os.Stdin)
	scanner.Scan()
}

func convertTxt2Csv(fileName string) {

	fp, err := os.Open(fileName)
	if err != nil {
		failOnError("ファイル読込に失敗しました", err)
	}
	defer fp.Close()

	tsv := csv.NewReader(fp)
	tsv.Comma = '\t'
	tsv.FieldsPerRecord = -1

	records, err := tsv.ReadAll()
	if err != nil {
		failOnError("テキスト->TSV変換エラー", err)
	}

	addressExcel := xlsx.NewFile()
	addressSheet, err := addressExcel.AddSheet("Sheet1")
	if err != nil {
		failOnError("address.xlsxへのsheet追加エラー", err)
	}

	taskExcel := xlsx.NewFile()
	taskSheet, err := taskExcel.AddSheet("Sheet1")
	if err != nil {
		failOnError("task.xlsxへのsheet追加エラー", err)
	}

	appendAddressHeader(addressSheet)
	appendTaskHeader(taskSheet)

	addressIndex := 1
	taskIndex := 1
	needNumOfAddressRecord := 0
	addressItems := []string{}
	for _, items := range records {
		switch items[0] {
		case "*Section":
			addressIndex = appendAddressRow(
				addressSheet, addressItems, addressIndex, needNumOfAddressRecord)
			needNumOfAddressRecord = 0
		case "*S":
			addressItems = items
		case "*I":
			appendTaskRow(taskSheet, items, taskIndex, addressItems[1])
			taskIndex++
			needNumOfAddressRecord++
		}
	}

	appendAddressRow(addressSheet, addressItems, addressIndex, needNumOfAddressRecord)

	exe, err := os.Executable()
	if err != nil {
		failOnError("exeファイル実行パス取得失敗", err)
	}

	outputDirPath := filepath.Dir(exe)
	err = addressExcel.Save(outputDirPath + "/address.xlsx")

	if err != nil {
		failOnError("address.xlsxの保存に失敗しました", err)
	}
	fmt.Println(outputDirPath + "\\address.xlsxを出力しました")

	err = taskExcel.Save(outputDirPath + "/task.xlsx")
	if err != nil {
		failOnError("task.xlsxの保存に失敗しました", err)
	}
	fmt.Println(outputDirPath + "\\task.xlsxを出力しました")
}

func appendAddressRow(sheet *xlsx.Sheet, items []string, rowNumber, needNumOfAddressRecord int) int {
	if needNumOfAddressRecord == 0 {
		return rowNumber
	}

	for i := 0; i < needNumOfAddressRecord; i++ {
		sheet.Cell(rowNumber, 0).Value = items[3] // First Name
		sheet.Cell(rowNumber, 1).Value = items[2] // Last Name
		sheet.Cell(rowNumber, 4).Value = "Japan"
		sheet.Cell(rowNumber, 5).Value = items[6]  // City
		sheet.Cell(rowNumber, 6).Value = items[7]  // Address
		sheet.Cell(rowNumber, 8).Value = items[5]  // State
		sheet.Cell(rowNumber, 9).Value = items[4]  // Zip Code
		sheet.Cell(rowNumber, 10).Value = items[8] // TEL
		sheet.Cell(rowNumber, 22).Value = fmt.Sprintf("Profile%d", rowNumber)
		_, err := strconv.Atoi(items[14])
		if err == nil {
			// Card
			sheet.Cell(rowNumber, 23).Value = items[14] // Card Number
			cardMonth, err := strconv.Atoi(items[15])   // Month
			if err != nil {
				failOnError(
					fmt.Sprintf("カード期限月の数値変換エラー。 テキストファイルセクション%s\n", items[1]),
					nil)
			}
			sheet.Cell(rowNumber, 24).SetInt(cardMonth)

			cardYear, err := strconv.Atoi(items[16]) // Year
			if err != nil {
				failOnError(
					fmt.Sprintf("カード期限年の数値変換エラー。 テキストファイルセクション%s\n", items[1]),
					nil)
			}
			sheet.Cell(rowNumber, 25).SetInt(cardYear)

			// Check digit
			_, err = strconv.Atoi(items[17]) // CVV
			if err != nil {
				failOnError(
					fmt.Sprintf("カードCVVの数値変換エラー。 テキストファイルセクション%s\n", items[1]),
					nil)
			}
			sheet.Cell(rowNumber, 26).Value = items[17] // CVV
			sheet.Cell(rowNumber, 33).Value = "false"
		} else {
			// 代金引換
			sheet.Cell(rowNumber, 33).Value = "true"
		}
		sheet.Cell(rowNumber, 28).Value = items[9] // Email
		sheet.Cell(rowNumber, 32).Value = "No checkout limit"
		rowNumber++
	}

	return rowNumber
}

func appendAddressHeader(sheet *xlsx.Sheet) {
	sheet.Cell(0, 0).Value = "First"
	sheet.Cell(0, 1).Value = "Last"
	sheet.Cell(0, 4).Value = "Country"
	sheet.Cell(0, 5).Value = "City"
	sheet.Cell(0, 6).Value = "Address"
	sheet.Cell(0, 7).Value = "Apt/House"
	sheet.Cell(0, 8).Value = "State"
	sheet.Cell(0, 9).Value = "ZipCode"
	sheet.Cell(0, 10).Value = "Phone"
	sheet.Cell(0, 11).Value = "First"
	sheet.Cell(0, 12).Value = "Last"
	sheet.Cell(0, 15).Value = "Country"
	sheet.Cell(0, 16).Value = "City"
	sheet.Cell(0, 17).Value = "Address"
	sheet.Cell(0, 18).Value = "Apt / House"
	sheet.Cell(0, 19).Value = "State"
	sheet.Cell(0, 20).Value = "ZipCode"
	sheet.Cell(0, 21).Value = "Phone"
	sheet.Cell(0, 22).Value = "Profile Name"
	sheet.Cell(0, 23).Value = "Card Number"
	sheet.Cell(0, 24).Value = "Month"
	sheet.Cell(0, 25).Value = "Year"
	sheet.Cell(0, 26).Value = "CVV"
	sheet.Cell(0, 27).Value = "Card Name"
	sheet.Cell(0, 28).Value = "Email"
	sheet.Cell(0, 32).Value = "Checkout Limit"
	sheet.Cell(0, 33).Value = "Use COD"
}

func appendTaskRow(sheet *xlsx.Sheet, items []string, rowNumber int, sectionId string) {
	sheet.Cell(rowNumber, 0).Value = fmt.Sprintf("Task%d", rowNumber)
	sheet.Cell(rowNumber, 1).Value = "Sample"
	sheet.Cell(rowNumber, 2).Value = fmt.Sprintf("Profile%d", rowNumber)
	sheet.Cell(rowNumber, 3).Value = items[4]  // Keyword
	sheet.Cell(rowNumber, 4).SetInt(rowNumber) // ProxyList

	switch items[5] { // Category
	case "tops_sweaters":
		sheet.Cell(rowNumber, 5).Value = "Tops/Sweaters"
	case "t-shirts":
		sheet.Cell(rowNumber, 5).Value = "T-Shirts"
	default:
		sheet.Cell(rowNumber, 5).Value = strings.ToUpper(items[5][:1]) + items[5][1:]
	}

	sheet.Cell(rowNumber, 6).Value = items[3] // Color

	switch strings.ToLower(items[2]) { // Size
	case "":
		sheet.Cell(rowNumber, 7).Value = "Random"
	case "s":
		fallthrough
	case "sランダム":
		sheet.Cell(rowNumber, 7).Value = "Small"
	case "m":
		fallthrough
	case "mランダム":
		sheet.Cell(rowNumber, 7).Value = "Medium"
	case "l":
		fallthrough
	case "lランダム":
		sheet.Cell(rowNumber, 7).Value = "Large"
	case "xl":
		fallthrough
	case "xlランダム":
		sheet.Cell(rowNumber, 7).Value = "XLarge"
	case "s/m":
		sheet.Cell(rowNumber, 7).Value = "S/M"
	case "l/xl":
		sheet.Cell(rowNumber, 7).Value = "L/XL"
	default:
		fmt.Printf(
			"サイズ'%s'は、規定外のため、そのままtask.xlsxに入力します。Excel行%d テキストファイルセクション%s\n",
			items[2], rowNumber, sectionId)
		sheet.Cell(rowNumber, 7).Value = items[2]
	}

	sheet.Cell(rowNumber, 8).Value = "super"
	sheet.Cell(rowNumber, 9).SetInt(0)
	sheet.Cell(rowNumber, 10).Value = "true"
	sheet.Cell(rowNumber, 11).Value = "true"
	sheet.Cell(rowNumber, 12).Value = "true"
}

func appendTaskHeader(sheet *xlsx.Sheet) {
	sheet.Cell(0, 0).Value = "Task Name"
	sheet.Cell(0, 1).Value = "Sheet Name"
	sheet.Cell(0, 2).Value = "Profile Name"
	sheet.Cell(0, 3).Value = "Keywords/Link"
	sheet.Cell(0, 4).Value = "Proxy List"
	sheet.Cell(0, 5).Value = "Category"
	sheet.Cell(0, 6).Value = "Color"
	sheet.Cell(0, 7).Value = "Size"
	sheet.Cell(0, 8).Value = "Mode"
	sheet.Cell(0, 9).Value = "Delay"
	sheet.Cell(0, 10).Value = "RetryOnFailure"
	sheet.Cell(0, 11).Value = "Restock Mode"
	sheet.Cell(0, 12).Value = "captcha bypass"
}
