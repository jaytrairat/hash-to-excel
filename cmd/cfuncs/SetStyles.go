package cfuncs

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

const (
	startDataRow = 2
)

func SetStyles(f *excelize.File, recordCount int) error {
	headerStyle, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold:   true,
			Size:   16,
			Family: "TH Sarabun New",
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		return fmt.Errorf("error creating header style: %w", err)
	}
	f.SetCellStyle("Sheet1", "A1", "D1", headerStyle)

	indexStyle, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Size:   16,
			Family: "TH Sarabun New",
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})

	contentStyle, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Size:   16,
			Family: "TH Sarabun New",
		},
		Alignment: &excelize.Alignment{
			WrapText: true,
			Vertical: "center",
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})

	rightAlignedStyle, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Size:   16,
			Family: "TH Sarabun New",
		},
		Alignment: &excelize.Alignment{
			Horizontal: "right",
			Vertical:   "center",
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})

	for i := startDataRow; i <= recordCount+1; i++ {
		f.SetCellStyle("Sheet1", fmt.Sprintf("A%d", i), fmt.Sprintf("A%d", i), indexStyle)
		f.SetCellStyle("Sheet1", fmt.Sprintf("B%d", i), fmt.Sprintf("B%d", i), contentStyle)
		f.SetCellStyle("Sheet1", fmt.Sprintf("C%d", i), fmt.Sprintf("C%d", i), contentStyle)
		f.SetCellStyle("Sheet1", fmt.Sprintf("D%d", i), fmt.Sprintf("D%d", i), rightAlignedStyle)
	}

	return nil
}
