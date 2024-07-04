package excel_exporter

import (
	"fmt"
	"sync"

	"github.com/xuri/excelize/v2"
)

// SheetMaxRows defines the maximum number of rows per sheet for Excel 2007 and later versions (.xlsx format).
const SheetMaxRows = 1048576

// MergeCell defines a merged cell data.
type MergeCell struct {
	TopLeftCell     string
	BottomRightCell string
}

// Row represents a row of data in the Excel sheet.
type Row struct {
	Cells      []excelize.Cell    // Cells in the row
	MergeCells []MergeCell        // Merged cells in the row
	RowOpts    []excelize.RowOpts // Options for the row, only useful when useStreamWriter is true
}

// RowDataFunc is a function type that returns the next row of data or nil if no more data.
type RowDataFunc func() Row

// SheetData represents the data for a single sheet.
type SheetData struct {
	Name    string
	RowFunc RowDataFunc
}

// Exporter provides methods for exporting data to Excel files.
type Exporter struct {
	File            *excelize.File
	FileName        string
	CurrentSheet    string // Current sheet name
	UseStreamWriter bool
	StreamWriter    *excelize.StreamWriter
}

// New creates a new Exporter instance.
func New(fileName string, useStreamWriter bool) *Exporter {
	return &Exporter{
		File:            excelize.NewFile(),
		FileName:        fileName,
		UseStreamWriter: useStreamWriter,
	}
}

// Export exports the Excel file.
func (e *Exporter) Export(sheets []SheetData) error {
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
			if err := e.exportUsingStreamWriter(sheet); err != nil {
				return err
			}
		} else {
			if err := e.exportUsingMemory(sheet); err != nil {
				return err
			}
		}
	}

	return e.File.SaveAs(e.FileName)
}

func (e *Exporter) exportUsingStreamWriter(sheet SheetData) error {
	initFunc := func(sheetName string) error {
		var err error
		e.StreamWriter, err = e.File.NewStreamWriter(sheetName)
		return err
	}

	writeRowFunc := func(sheetName string, rowID int, row Row) error {
		rowCells := make([]interface{}, len(row.Cells))
		for j, cell := range row.Cells {
			rowCells[j] = cell
		}

		cell, _ := excelize.CoordinatesToCellName(1, rowID)
		if err := e.StreamWriter.SetRow(cell, rowCells, row.RowOpts...); err != nil {
			return err
		}

		for _, mergeCell := range row.MergeCells {
			if err := e.StreamWriter.MergeCell(mergeCell.TopLeftCell, mergeCell.BottomRightCell); err != nil {
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

func (e *Exporter) exportUsingMemory(sheet SheetData) error {
	initFunc := func(sheetName string) error {
		return nil
	}

	writeRowFunc := func(sheetName string, rowID int, row Row) error {
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
			if err := e.File.MergeCell(sheetName, mergeCell.TopLeftCell, mergeCell.BottomRightCell); err != nil {
				return err
			}
		}

		return nil
	}

	return e.exportHelper(sheet, initFunc, writeRowFunc)
}

func (e *Exporter) exportHelper(sheet SheetData, initFunc func(string) error, writeRowFunc func(string, int, Row) error) error {
	rowID := 1
	sheetSuffix := 0
	e.CurrentSheet = sheet.Name

	if err := initFunc(e.CurrentSheet); err != nil {
		return err
	}

	for {
		row := sheet.RowFunc()
		if row.Cells == nil {
			break
		}

		if rowID > SheetMaxRows {
			// Create a new sheet if row count exceeds SheetMaxRows
			sheetSuffix++
			rowID = 1

			currentSheetName := fmt.Sprintf("%s_%d", sheet.Name, sheetSuffix)
			if _, err := e.File.NewSheet(currentSheetName); err != nil {
				return fmt.Errorf("failed to create a new sheet: %w", err)
			}

			e.CurrentSheet = currentSheetName
			if err := initFunc(e.CurrentSheet); err != nil {
				return err
			}
		}

		if err := writeRowFunc(e.CurrentSheet, rowID, row); err != nil {
			return err
		}

		rowID++
	}

	return nil
}

// UseRowChan returns a RowDataFunc that will use a channel to send Row objects to the given function
func UseRowChan(sendDataFunc func(dataCh chan Row)) RowDataFunc {
	var once sync.Once
	var dataCh chan Row
	return func() Row {
		once.Do(func() {
			dataCh = make(chan Row)
			go func() {
				defer close(dataCh)
				sendDataFunc(dataCh)
			}()
		})

		row, ok := <-dataCh
		if !ok {
			return Row{}
		}
		return row
	}
}
