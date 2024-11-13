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

const (
	outputFileName = "formatted_hash.xlsx"
	headerRow      = 1
	fileNameIndex  = 0
	hashIndex      = 1
	fileSizeIndex  = 11
)

var headers = []string{"ลำดับ", "File name", "SHA-256", "File size"}

var columnWidths = map[string]float64{
	"A": 7,
	"B": 40,
	"C": 75,
	"D": 15,
}

var rootCmd = &cobra.Command{
	Use:   "hash-to-excel",
	Short: "Read a CSV file and write specific fields to a new Excel file",
	Run: func(cmd *cobra.Command, args []string) {
		inputFile, err := cmd.Flags().GetString("input")
		if err != nil {
			fmt.Println("Error retrieving input flag:", err)
			return
		}

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

func createExcelFile(f *excelize.File, records [][]string) error {
	f.SetSheetRow("Sheet1", "A1", &headers)

	for i, record := range records {
		if len(record) >= 12 {
			halfLengthOfHash := len(record[hashIndex]) / 2
			splitedHash := record[hashIndex][:halfLengthOfHash] + "\r\n" + record[hashIndex][halfLengthOfHash:]

			newRecord := []interface{}{i + 1, record[fileNameIndex], splitedHash, record[fileSizeIndex]}
			f.SetSheetRow("Sheet1", fmt.Sprintf("A%d", i+startDataRow), &newRecord)
		} else {
			fmt.Println("Skipping record due to insufficient fields:", record)
		}
	}

	if err := setColumnWidths(f); err != nil {
		return err
	}

	if err := cfuncs.setStyles(f, len(records)); err != nil {
		return err
	}

	return nil
}

func setColumnWidths(f *excelize.File) error {
	for col, width := range columnWidths {
		if err := f.SetColWidth("Sheet1", col, col, width); err != nil {
			return fmt.Errorf("error setting column width: %w", err)
		}
	}
	return nil
}

func saveExcelFile(f *excelize.File) error {
	return f.SaveAs(outputFileName)
}

func Execute() {
	if err := rootCmd.Execute(); err != nil {
		os.Exit(1)
	}
}
