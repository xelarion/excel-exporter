package excel_exporter

import "github.com/xuri/excelize/v2"

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

// NewRow creates a new Row with the specified cell values.
func NewRow(cellValues ...interface{}) Row {
	cells := make([]excelize.Cell, len(cellValues))
	for i, cellValue := range cellValues {
		cells[i] = excelize.Cell{Value: cellValue}
	}
	return Row{Cells: cells}
}
