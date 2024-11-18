/*
Copyright © 2024 JAYTRAIRAT jay.trairat@gmail.com
*/
package cmd

import (
	"encoding/csv"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/jaytrairat/hash-to-excel/cmd/cfuncs"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

const (
	outputExcelFileName = "formatted_hash.xlsx"
	outputTextFileName  = "formatted_hash.txt"
	headerRow           = 1
	fileNameIndex       = 0
	hashIndex           = 1
	fileSizeIndex       = 11
	startDataRow        = 2
)

var headers = []string{"ลำดับ", "File name", "SHA-256", "File size"}

var rootCmd = &cobra.Command{
	Use:   "hash-to-excel",
	Short: "Read a CSV file and write specific fields to a new Excel file",
	Run: func(cmd *cobra.Command, args []string) {
		inputFile, _ := cmd.Flags().GetString("input")

		if inputFile == "" {
			inputFile, _ = findFirstCSVFile()
		}

		if err := parseFile(inputFile); err != nil {
			fmt.Println("Failed to process files:", err)
		} else {
			fmt.Printf("Successfully processed the CSV file (%s) and created %s", inputFile, outputExcelFileName)
		}
	},
}

func init() {
	rootCmd.Flags().StringP("input", "i", "", "Path to the input CSV file")
}

func parseFile(inputFilename string) error {
	records, _ := readCSV(inputFilename)
	if len(records) == 0 {
		fmt.Println("No records found in the input CSV file.")
	}

	f := excelize.NewFile()
	if err := createExcelFile(f, records); err != nil {
		return err
	}

	if err := createTextFile(records); err != nil {
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

		splitedFileNameField := strings.Split(record[fileNameIndex], "_")
		formattedFileNameField := strings.Join(splitedFileNameField[0:3], "_") + "_\n" + strings.Join(splitedFileNameField[3:], "_")

		halfLengthOfHashField := len(record[hashIndex]) / 2
		splitedHashField := record[hashIndex][:halfLengthOfHashField] + "\r\n" + record[hashIndex][halfLengthOfHashField:]

		columnData := []interface{}{i + 1, formattedFileNameField, splitedHashField, record[fileSizeIndex]}
		f.SetSheetRow("Sheet1", fmt.Sprintf("A%d", i+startDataRow), &columnData)
	}

	cfuncs.SetColumnWidths(f)
	cfuncs.SetStyles(f, len(records))
	return nil
}

func createTextFile(records [][]string) error {
	var formatted []string
	for _, record := range records {
		fileNameField := record[fileNameIndex]
		hashFile := record[hashIndex]
		formatted = append(formatted, fmt.Sprintf(
			"รายละเอียดปรากฏตามไฟล์ประกอบรายงาน ชื่อไฟล์ %s ค่า Hash SHA256: %s",
			fileNameField,
			hashFile,
		))
	}

	content := strings.Join(formatted, "\n")

	err := os.WriteFile(outputTextFileName, []byte(content), 0644)
	if err != nil {
		return fmt.Errorf("failed to write to file %s: %w", outputTextFileName, err)
	}

	return nil
}

func saveExcelFile(f *excelize.File) error {
	return f.SaveAs(outputExcelFileName)
}

func Execute() {
	if err := rootCmd.Execute(); err != nil {
		os.Exit(1)
	}
}
