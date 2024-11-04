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
		err := parseFile(args[0])
		if err != nil {
			fmt.Println("Failed to process files:", err)
		} else {
			fmt.Println("Successfully processed the CSV file and created formatted_hash file.")
		}
	},
}

func parseFile(inputFilename string) error {
	file, err := os.Open(inputFilename)
	if err != nil {
		return fmt.Errorf("error opening input file: %w", err)
	}
	defer file.Close()

	records, err := csv.NewReader(file).ReadAll()
	if err != nil {
		return fmt.Errorf("error reading the CSV file: %w", err)
	}

	if len(records) == 0 {
		fmt.Println("No records found in the input CSV file.")
		return nil
	}

	f := excelize.NewFile()
	headers := []string{"ลำดับ", "File name", "SHA-256", "File size"}
	f.SetSheetRow("Sheet1", "A1", &headers)

	for i, record := range records {
		if len(record) >= 12 {
			newRecord := []interface{}{i + 1, record[0], record[1], record[11]}
			f.SetSheetRow("Sheet1", fmt.Sprintf("A%d", i+2), &newRecord)
		} else {
			fmt.Println("Skipping record due to insufficient fields:", record)
		}
	}

	widths := map[string]float64{"A": 7, "B": 40, "C": 75, "D": 15}
	for col, w := range widths {
		if err := f.SetColWidth("Sheet1", col, col, w); err != nil {
			return fmt.Errorf("error setting column width: %w", err)
		}
	}

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

	for i := 2; i <= len(records)+1; i++ {
		f.SetCellStyle("Sheet1", fmt.Sprintf("A%d", i), fmt.Sprintf("A%d", i), centerStyle)
		f.SetCellStyle("Sheet1", fmt.Sprintf("D%d", i), fmt.Sprintf("D%d", i), rightStyle)
	}

	return f.SaveAs("formatted_hash.xlsx")
}

func init() {
	rootCmd.AddCommand(hashToExcelCmd)
}
