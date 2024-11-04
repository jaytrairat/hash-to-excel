/*
Copyright © 2024 JAYTRAIRAT jay.trairat@gmail.com
*/
package cmd

import (
	"encoding/csv"
	"fmt"
	"os"
	"path/filepath"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

// Constants for column headers and widths
const (
	outputFileName = "formatted_hash.xlsx"
	headerRow      = 1
	startDataRow   = 2
	fileNameIndex  = 0
	hashIndex      = 1
	fileSizeIndex  = 11
)

// Column headers for the Excel file
var headers = []string{"ลำดับ", "File name", "SHA-256", "File size"}

// Widths for each column in the Excel file
var columnWidths = map[string]float64{
	"A": 7,
	"B": 40,
	"C": 75,
	"D": 15,
}

// rootCmd represents the base command
var rootCmd = &cobra.Command{
	Use:   "hash-to-excel",
	Short: "Read a CSV file and write specific fields to a new Excel file",
	Run: func(cmd *cobra.Command, args []string) {
		inputFile, err := cmd.Flags().GetString("input")
		if err != nil {
			fmt.Println("Error retrieving input flag:", err)
			return
		}

		// If no input file is provided, search for the first CSV file in the current directory
		if inputFile == "" {
			inputFile, err = findFirstCSVFile()
			if err != nil {
				fmt.Println(err)
				return
			}
		}

		if err := parseFile(inputFile); err != nil {
			fmt.Println("Failed to process files:", err)
		} else {
			fmt.Printf("Successfully processed the CSV file (%s) and created %s", inputFile, outputFileName)
		}
	},
}

func init() {
	rootCmd.Flags().StringP("input", "i", "", "Path to the input CSV file")
}

// parseFile reads the input CSV file and generates the Excel file
func parseFile(inputFilename string) error {
	records, err := readCSV(inputFilename)
	if err != nil {
		return err
	}

	if len(records) == 0 {
		fmt.Println("No records found in the input CSV file.")
		return nil
	}

	f := excelize.NewFile()
	if err := createExcelFile(f, records); err != nil {
		return err
	}

	return saveExcelFile(f)
}

// readCSV reads a CSV file and returns the records
func readCSV(filename string) ([][]string, error) {
	file, err := os.Open(filename)
	if err != nil {
		return nil, fmt.Errorf("error opening input file: %w", err)
	}
	defer file.Close()

	records, err := csv.NewReader(file).ReadAll()
	if err != nil {
		return nil, fmt.Errorf("error reading the CSV file: %w", err)
	}

	return records, nil
}

// findFirstCSVFile searches the current directory for the first CSV file and returns its name
func findFirstCSVFile() (string, error) {
	files, err := os.ReadDir(".")
	if err != nil {
		return "", fmt.Errorf("error reading current directory: %w", err)
	}

	for _, file := range files {
		if !file.IsDir() && filepath.Ext(file.Name()) == ".csv" {
			return file.Name(), nil
		}
	}

	return "", fmt.Errorf("no CSV files found in the current directory")
}

// createExcelFile populates the Excel file with data from the CSV records
func createExcelFile(f *excelize.File, records [][]string) error {
	f.SetSheetRow("Sheet1", "A1", &headers)

	for i, record := range records {
		if len(record) >= 12 {
			newRecord := []interface{}{i + 1, record[fileNameIndex], record[hashIndex], record[fileSizeIndex]}
			f.SetSheetRow("Sheet1", fmt.Sprintf("A%d", i+startDataRow), &newRecord)
		} else {
			fmt.Println("Skipping record due to insufficient fields:", record)
		}
	}

	if err := setColumnWidths(f); err != nil {
		return err
	}

	return setStyles(f, len(records))
}

// setColumnWidths sets the widths for the specified columns
func setColumnWidths(f *excelize.File) error {
	for col, width := range columnWidths {
		if err := f.SetColWidth("Sheet1", col, col, width); err != nil {
			return fmt.Errorf("error setting column width: %w", err)
		}
	}
	return nil
}

// setStyles applies the styles to the cells in the Excel file
func setStyles(f *excelize.File, recordCount int) error {
	headerStyle, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})
	if err != nil {
		return fmt.Errorf("error creating header style: %w", err)
	}
	f.SetCellStyle("Sheet1", "A1", "D1", headerStyle)

	centerStyle, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
	})
	if err != nil {
		return fmt.Errorf("error creating center style: %w", err)
	}

	rightStyle, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "right",
		},
	})
	if err != nil {
		return fmt.Errorf("error creating right style: %w", err)
	}

	for i := startDataRow; i <= recordCount+1; i++ {
		f.SetCellStyle("Sheet1", fmt.Sprintf("A%d", i), fmt.Sprintf("A%d", i), centerStyle)
		f.SetCellStyle("Sheet1", fmt.Sprintf("D%d", i), fmt.Sprintf("D%d", i), rightStyle)
	}

	return nil
}

// saveExcelFile saves the Excel file to the output filename
func saveExcelFile(f *excelize.File) error {
	return f.SaveAs(outputFileName)
}

// Execute executes the root command
func Execute() {
	if err := rootCmd.Execute(); err != nil {
		os.Exit(1)
	}
}
