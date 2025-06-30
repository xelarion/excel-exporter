package excel_exporter

import (
	"fmt"
	"sync"

	"github.com/xuri/excelize/v2"
)

// SheetMaxRows defines the maximum number of rows per sheet for Excel 2007 and later versions (.xlsx format).
const SheetMaxRows = 1 << 20

// RowDataFunc is a function type that returns the next row of data and an error if any.
// The rowNumber parameter indicates the current Excel row number (starting from 1).
type RowDataFunc func(rowNumber int) (Row, error)

// InitFunc is a function type that will be called at the beginning of each sheet.
type InitFunc func(exporter *Exporter) error

// SheetData represents the data for a single sheet.
type SheetData struct {
	Name     string
	RowFunc  RowDataFunc
	InitFunc InitFunc
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
	// call close to remove temp files
	defer e.File.Close()

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
		// cell merge and style will be lost if no flush
		if e.StreamWriter != nil {
			err = e.StreamWriter.Flush()
			if err != nil {
				return err
			}
		}
		e.StreamWriter, err = e.File.NewStreamWriter(sheetName)
		if err != nil {
			return err
		}
		if sheet.InitFunc != nil {
			err = sheet.InitFunc(e)
			if err != nil {
				return err
			}
		}
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
		if sheet.InitFunc != nil {
			return sheet.InitFunc(e)
		}
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
	var rowIndex = 0
	for {
		row, err := sheet.RowFunc(rowIndex)
		if err != nil {
			return err
		}

		if row.Cells == nil {
			break
		}

		if rowID > SheetMaxRows {
			sheetSuffix++
			rowID = 1

			// Create a new sheet if row count exceeds SheetMaxRows
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
		rowIndex++
	}

	return nil
}

// UseRowChan returns a RowDataFunc that will use a channel to send Row objects to the given function.
func UseRowChan(sendDataFunc func(dataCh chan Row) error) RowDataFunc {
	var once sync.Once
	var dataCh chan Row
	var sendErr error

	return func(rowNumber int) (Row, error) {
		once.Do(func() {
			dataCh = make(chan Row)
			go func() {
				defer close(dataCh)
				sendErr = sendDataFunc(dataCh)
			}()
		})

		row, ok := <-dataCh
		if sendErr != nil {
			return Row{}, sendErr
		}

		if !ok {
			return Row{}, nil
		}
		return row, nil
	}
}
