package main

import (
	"encoding/csv"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

func findHeaderRow(rows [][]string) (int, int) {
	if len(rows) == 0 {
		return -1, 0
	}

	maxCols := 0
	for _, row := range rows {
		nonEmptyCells := 0
		for _, cell := range row {
			if cell != "" {
				nonEmptyCells++
			}
		}
		if nonEmptyCells > maxCols {
			maxCols = nonEmptyCells
		}
	}

	for i, row := range rows {
		nonEmptyCells := 0
		for _, cell := range row {
			if cell != "" {
				nonEmptyCells++
			}
		}
		if nonEmptyCells == maxCols {
			return i, maxCols
		}
	}
	return -1, maxCols
}

func main() {
	inputFile := "inventory_copy_2.xlsx"
	// Create output filename by changing extension to .csv
	outputFile := strings.TrimSuffix(inputFile, filepath.Ext(inputFile)) + ".csv"

	f, err := excelize.OpenFile(inputFile)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer f.Close()

	sheet := f.GetSheetList()[0]
	rows, err := f.GetRows(sheet)
	if err != nil {
		fmt.Println(err)
		return
	}

	headerRowIndex, maxCols := findHeaderRow(rows)
	if headerRowIndex < 0 {
		fmt.Println("Could not find header row")
		return
	}

	fmt.Printf("Found header row at index %d with %d columns\n", headerRowIndex+1, maxCols)

	// Process header rows (same as original code)
	for colNum := 1; colNum <= maxCols; colNum++ {
		colName, err := excelize.ColumnNumberToName(colNum)
		if err != nil {
			fmt.Printf("Error converting column number: %v\n", err)
			continue
		}

		headerCellRef := fmt.Sprintf("%s%d", colName, headerRowIndex+1)
		headerValue, err := f.GetCellValue(sheet, headerCellRef)
		if err != nil {
			fmt.Printf("Error getting header cell value: %v\n", err)
			continue
		}

		var newHeaderValue string = headerValue
		for rowNum := 1; rowNum <= headerRowIndex; rowNum++ {
			cellRef := fmt.Sprintf("%s%d", colName, rowNum)
			value, err := f.GetCellValue(sheet, cellRef)
			if err != nil {
				fmt.Printf("Error getting cell value: %v\n", err)
				continue
			}
			if value != "" {
				newHeaderValue = newHeaderValue + " " + value
			}
		}

		err = f.SetCellValue(sheet, headerCellRef, newHeaderValue)
		if err != nil {
			fmt.Printf("Error setting new header value: %v\n", err)
		}
	}

	// Remove rows above header
	for i := headerRowIndex; i > 0; i-- {
		err := f.RemoveRow(sheet, i)
		if err != nil {
			fmt.Printf("Error removing row %d: %v\n", i, err)
		}
	}

	// Get all rows after processing
	processedRows, err := f.GetRows(sheet)
	if err != nil {
		fmt.Printf("Error getting processed rows: %v\n", err)
		return
	}

	// Create CSV file
	csvFile, err := os.Create(outputFile)
	if err != nil {
		fmt.Printf("Error creating CSV file: %v\n", err)
		return
	}
	defer csvFile.Close()

	// Create CSV writer
	writer := csv.NewWriter(csvFile)
	defer writer.Flush()

	// Write all rows to CSV
	for _, row := range processedRows {
		err := writer.Write(row)
		if err != nil {
			fmt.Printf("Error writing row to CSV: %v\n", err)
			continue
		}
	}

	fmt.Printf("Successfully converted to CSV: %s\n", outputFile)
}
