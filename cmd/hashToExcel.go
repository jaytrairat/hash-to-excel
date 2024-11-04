package cmd

import (
	"encoding/csv"
	"fmt"
	"os"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

var hashToExcelCmd = &cobra.Command{
	Use:   "hash-to-excel [input-file]",
	Short: "Read a CSV file and write specific fields to a new Excel file",
	Args:  cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		inputFilename := args[0]
		if err := parseFile(inputFilename); err != nil {
			fmt.Println("Failed to process files:", err)
		}
	},
}

func parseFile(inputFilename string) error {
	file, err := os.Open(inputFilename)
	if err != nil {
		return fmt.Errorf("error opening input file: %w", err)
	}
	defer file.Close()

	reader := csv.NewReader(file)
	records, err := reader.ReadAll()
	if err != nil {
		return fmt.Errorf("error reading the CSV file: %w", err)
	}

	if len(records) == 0 {
		fmt.Println("No records found in the input CSV file.")
		return nil
	}

	f := excelize.NewFile()

	headers := []string{"ลำดับ", "File name", "SHA-256", "File size"}
	if err := f.SetSheetRow("Sheet1", "A1", &headers); err != nil {
		return fmt.Errorf("error setting headers: %w", err)
	}

	for i, record := range records {
		if len(record) >= 8 {
			newRecord := []interface{}{
				i + 1,
				record[0],
				record[1],
				record[11],
			}
			cell, err := excelize.CoordinatesToCellName(1, i+2)
			if err != nil {
				return fmt.Errorf("error converting coordinates: %w", err)
			}
			if err := f.SetSheetRow("Sheet1", cell, &newRecord); err != nil {
				return fmt.Errorf("error writing to output file: %w", err)
			}
		} else {
			fmt.Println("Skipping record due to insufficient fields:", record)
		}
	}

	if err := f.SaveAs("output.xlsx"); err != nil {
		return fmt.Errorf("error saving output file: %w", err)
	}

	return nil
}

func init() {
	rootCmd.AddCommand(hashToExcelCmd)
}
