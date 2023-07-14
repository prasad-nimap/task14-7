package main

import (
	"fmt"
	"log"
	"os"

	"github.com/xuri/excelize/v2"
)

func columnNumberToLetter(columnNumber int) string {
	columnLetter := ""
	for columnNumber > 0 {
		remainder := (columnNumber - 1) % 26
		columnLetter = string('A'+remainder) + columnLetter
		columnNumber = (columnNumber - 1) / 26
	}
	return columnLetter
}

func main() {
	args := os.Args[1:] // Skip the first argument as it contains the program name
	if len(args) < 2 {
		fmt.Println("Usage: go run main.go <source workbook> <destination workbook>")
		return
	}

	sourceWorkbook := args[0]
	destinationWorkbook := args[1]
	fmt.Println("Source workbook:", sourceWorkbook)
	fmt.Println("Destination workbook:", destinationWorkbook)

	// Open the source workbook
	sourceFile, err := excelize.OpenFile(sourceWorkbook)
	if err != nil {
		fmt.Println("Failed to open source workbook:", err)
		return
	}
	defer func() {
		if err := sourceFile.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// Create a new destination workbook
	destFile := excelize.NewFile()

	// Copy sheets from the source workbook to the destination workbook
	sourceSheets := []string{"Jan", "Feb"}
	for _, sourceSheet := range sourceSheets {
		// Read data from the source workbook
		rows, err := sourceFile.GetRows(sourceSheet)
		if err != nil {
			fmt.Printf("Failed to get rows from source sheet '%s': %v\n", sourceSheet, err)
			continue
		}

		// Create a new sheet in the destination workbook
		destSheet := sourceSheet
		destFile.NewSheet(destSheet)

		// Write data to the destination workbook
		for rowIndex, row := range rows {
			for colIndex, cellValue := range row {
				colLetter := columnNumberToLetter(colIndex + 1)
				destFile.SetCellValue(destSheet, colLetter+fmt.Sprint(rowIndex+1), cellValue)

				cellCoords, err := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)
				if err != nil {
					log.Fatal(err)
				}

				//for formula
				formula, err := sourceFile.GetCellFormula(sourceSheets, cellCoords)
				if formula == "" {
					// Set the cell value in the destination sheet
					err = destFile.SetCellValue(destSheet, cellCoords, cellValue)
					if err != nil {
						log.Fatal(err)
					}
				} else {
					err = destFile.SetCellFormula(destSheet, cellCoords, formula)
					if err != nil {
						log.Fatal(err)
					}
				}

			}

		}

	}

	// Save the destination workbook
	if err := destFile.SaveAs(destinationWorkbook); err != nil {
		fmt.Println("Failed to save destination workbook:", err)
		return
	}

	fmt.Println("Sheets copied and pasted successfully.")
}
