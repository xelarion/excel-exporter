package excel_exporter

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// SheetMaxRows defines the maximum number of rows per sheet for Excel 2007 and later versions (.xlsx format).
const SheetMaxRows = 1048576

// MergeCell defines a merged cell data.
type MergeCell struct {
	HCell string
	VCell string
}

// Row represents a row of data in the Excel sheet.
type Row struct {
	Cells      []excelize.Cell    // Cells in the row
	MergeCells []MergeCell        // Merged cells in the row
	RowOpts    []excelize.RowOpts // Options for the row, only useful when useStreamWriter is true
}

// RowDataFunc is a function type that returns the next row of data or nil if no more data.
type RowDataFunc func() *Row

// SheetData represents the data for a single sheet.
type SheetData struct {
	Name    string
	RowFunc RowDataFunc
}

// ExcelExporter wraps *excelize.File and *excelize.StreamWriter for exporting Excel files.
type ExcelExporter struct {
	File            *excelize.File
	StreamWriter    *excelize.StreamWriter
	FileName        string
	UseStreamWriter bool
}

// NewExcelExporter creates a new ExcelExporter.
func NewExcelExporter(fileName string, useStreamWriter bool) *ExcelExporter {
	return &ExcelExporter{
		File:            excelize.NewFile(),
		FileName:        fileName,
		UseStreamWriter: useStreamWriter,
	}
}

// Export exports the Excel file.
func (e *ExcelExporter) Export(sheets []SheetData) error {
	for i, sheet := range sheets {
		if _, err := e.File.NewSheet(sheet.Name); err != nil {
			return fmt.Errorf("failed to create a new sheet: %w", err)
		}

		// delete default sheet
		if i == 0 && e.File.SheetCount > 1 {
			if err := e.File.DeleteSheet("Sheet1"); err != nil {
				return fmt.Errorf("failed to delete default sheet: %w", err)
			}
		}

		if e.UseStreamWriter {
			if err := e.exportWithStreamWriter(sheet); err != nil {
				return err
			}
		} else {
			if err := e.exportWithMemory(sheet); err != nil {
				return err
			}
		}
	}

	return e.File.SaveAs(e.FileName)
}

func (e *ExcelExporter) exportWithStreamWriter(sheet SheetData) error {
	initFunc := func(sheetName string) error {
		var err error
		e.StreamWriter, err = e.File.NewStreamWriter(sheetName)
		return err
	}

	writeRowFunc := func(sheetName string, rowID int, row *Row) error {
		rowCells := make([]interface{}, len(row.Cells))
		for j, cell := range row.Cells {
			rowCells[j] = cell
		}

		cell, _ := excelize.CoordinatesToCellName(1, rowID)
		if err := e.StreamWriter.SetRow(cell, rowCells, row.RowOpts...); err != nil {
			return err
		}

		for _, mergeCell := range row.MergeCells {
			if err := e.StreamWriter.MergeCell(mergeCell.HCell, mergeCell.VCell); err != nil {
				return err
			}
		}

		return nil
	}

	if err := e.exportHelper(sheet, initFunc, writeRowFunc); err != nil {
		return err
	}

	return e.StreamWriter.Flush()
}

func (e *ExcelExporter) exportWithMemory(sheet SheetData) error {
	initFunc := func(sheetName string) error {
		return nil
	}

	writeRowFunc := func(sheetName string, rowID int, row *Row) error {
		for j, cell := range row.Cells {
			cellName, _ := excelize.CoordinatesToCellName(j+1, rowID)
			if err := e.File.SetCellValue(sheetName, cellName, cell.Value); err != nil {
				return err
			}

			if cell.StyleID > 0 {
				if err := e.File.SetCellStyle(sheetName, cellName, cellName, cell.StyleID); err != nil {
					return err
				}
			}

			if cell.Formula != "" {
				if err := e.File.SetCellFormula(sheetName, cellName, cell.Formula); err != nil {
					return err
				}
			}
		}

		for _, mergeCell := range row.MergeCells {
			if err := e.File.MergeCell(sheetName, mergeCell.HCell, mergeCell.VCell); err != nil {
				return err
			}
		}

		return nil
	}

	return e.exportHelper(sheet, initFunc, writeRowFunc)
}

func (e *ExcelExporter) exportHelper(sheet SheetData, initFunc func(string) error, writeRowFunc func(string, int, *Row) error) error {
	rowID := 1
	sheetSuffix := 0
	currentSheetName := sheet.Name

	if err := initFunc(currentSheetName); err != nil {
		return err
	}

	for {
		row := sheet.RowFunc()
		if row == nil || row.Cells == nil {
			break
		}

		if rowID > SheetMaxRows {
			// Create a new sheet if row count exceeds SheetMaxRows
			sheetSuffix++
			currentSheetName = fmt.Sprintf("%s_%d", sheet.Name, sheetSuffix)
			rowID = 1

			if _, err := e.File.NewSheet(currentSheetName); err != nil {
				return fmt.Errorf("failed to create a new sheet: %w", err)
			}

			if err := initFunc(currentSheetName); err != nil {
				return err
			}
		}

		if err := writeRowFunc(currentSheetName, rowID, row); err != nil {
			return err
		}

		rowID++
	}

	return nil
}
