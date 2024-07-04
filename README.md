# Excel Exporter

`excel-exporter` is a Go package designed to simplify the process of exporting data to Excel files. It supports both in-memory and stream writing modes, making it efficient for handling large datasets. This package is built on top of the `github.com/xuri/excelize/v2` package for robust Excel file manipulation.

## Features

- Export data to Excel with support for multiple sheets.
- Efficient handling of large datasets using StreamWriter mode.
- Customizable cell styles, merged cells, and row options.
- Automatic handling of sheet row limits by creating new sheets.
- Built on top of the `github.com/xuri/excelize/v2` package for robust Excel file manipulation.

## Installation

To install the package, use:

```sh
go get -u github.com/xelarion/excel-exporter
```

## Usage

### Exporting Data

This example demonstrates how to use the `StreamWriter` mode to export a large dataset to an Excel file.

```go
package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	excelexporter "github.com/xelarion/excel-exporter"
)

func main() {
	// Set useStreamWriter to true for StreamWriter mode, false for in-memory mode
	useStreamWriter := true
	exporter := excelexporter.New("test_export.xlsx", useStreamWriter)
	sheets := []excelexporter.SheetData{
		{Name: "Sheet1", RowFunc: generateLargeData("Sheet1", 1500000)},
		{Name: "Sheet2", RowFunc: generateLargeData("Sheet2", 2000)},
	}

	if err := exporter.Export(sheets); err != nil {
		fmt.Printf("Failed to export Excel file: %v\n", err)
	}
}

func generateLargeData(sheetName string, rowCount int) excelexporter.RowDataFunc {
	currentRow := 0
	return func() excelexporter.Row {
		if currentRow >= rowCount {
			return excelexporter.Row{}
		}
		currentRow++
		return excelexporter.Row{
			Cells: []excelize.Cell{
				{Value: fmt.Sprintf("%s-a%d", sheetName, currentRow)},
				{Value: fmt.Sprintf("%s-b%d", sheetName, currentRow)},
				{Value: fmt.Sprintf("%s-c%d", sheetName, currentRow)},
			},
		}
	}
}
```

### Exporting Data Using Channel

This example demonstrates how to use channels with the `UseRowChan` function to export data to an Excel file.

```go
package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	excelexporter "github.com/xelarion/excel-exporter"
)

func main() {
	exporter := excelexporter.New("test_channel.xlsx", true)
	sheetNames := []string{"SheetA", "SheetB"}
	sheets := make([]excelexporter.SheetData, len(sheetNames))
	for i, name := range sheetNames {
		sheets[i] = excelexporter.SheetData{
			Name:    name,
			RowFunc: excelexporter.UseRowChan(queryDataToChannelFunc(exporter, name)),
		}
	}

	if err := exporter.Export(sheets); err != nil {
		fmt.Printf("Failed to export Excel file with StreamWriter: %v\n", err)
	}
}

func queryDataToChannelFunc(exporter *excelexporter.Exporter, sheetName string) func(dataCh chan excelexporter.Row) {
	return func(dataCh chan excelexporter.Row) {
		titleStyle, _ := exporter.File.NewStyle(
			&excelize.Style{
				Font:      &excelize.Font{Color: "777777", Size: 14},
				Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
			},
		)

		// Set column width
		if exporter.UseStreamWriter {
			// when use StreamWriter
			_ = exporter.StreamWriter.SetColWidth(1, 3, 30)
		} else {
			// when use memory
			_ = exporter.File.SetColWidth(exporter.CurrentSheet, "A", "C", 30)
		}

		dataCh <- excelexporter.Row{
			Cells: []excelize.Cell{
				{Value: "MergedTitle 1", StyleID: titleStyle},
				{Value: "", StyleID: titleStyle},
				{Value: "MergedTitle 2", StyleID: titleStyle},
			},
			// Merge cells
			MergeCells: []excelexporter.MergeCell{
				{TopLeftCell: "A1", BottomRightCell: "B1"},
			},
			// set row style, only useful when useStreamWriter is true
			RowOpts: []excelize.RowOpts{
				{Height: 20, StyleID: titleStyle},
			},
		}

		// Title
		dataCh <- excelexporter.Row{
			Cells: []excelize.Cell{
				{Value: "Title 1", StyleID: titleStyle},
				{Value: "Title 2", StyleID: titleStyle},
				{Value: "Title 3", StyleID: titleStyle},
			},
		}

		// Simulate querying data from the database and sending to channel
		for i := 0; i < 10; i++ {
			dataCh <- excelexporter.Row{
				Cells: []excelize.Cell{
					{Value: fmt.Sprintf("%s-%d-1", sheetName, i)},
					{Value: fmt.Sprintf("%s-%d-2", sheetName, i)},
					{Value: fmt.Sprintf("%s-%d-3", sheetName, i)},
				},
			}
		}
	}
}
```

## API Reference

### `Exporter`

The `Exporter` struct provides methods for exporting data to Excel files.

#### Methods

- `New(fileName string, useStreamWriter bool) *Exporter`: Creates a new `Exporter`.
- `Export(sheets []SheetData) error`: Exports the Excel file with the specified sheets.

### `SheetData`

The `SheetData` struct represents the data for a single sheet.

#### Fields

- `Name string`: The name of the sheet.
- `RowFunc RowDataFunc`: A function that returns the next row of data.

### `Row`

The `Row` struct represents a row of data in the Excel sheet.

#### Fields

- `Cells []excelize.Cell`: Cells in the row.
- `MergeCells []MergeCell`: Merged cells in the row.
- `RowOpts []excelize.RowOpts`: Options for the row.

### `MergeCell`

The `MergeCell` struct defines a merged cell data.

#### Fields

- `TopLeftCell string`: The starting cell to be merged.
- `BottomRightCell string`: The ending cell to be merged.

### `UseRowChan`

The `UseRowChan` function returns a `RowDataFunc` that uses a channel to send `Row` objects to the given function.

```go
func UseRowChan(sendDataFunc func(dataCh chan Row)) RowDataFunc
```

## License

This project is licensed under the MIT License.
