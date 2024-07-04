### README.md

```markdown
# Excel Exporter

`excel_exporter` is a Go package that provides functionality for exporting Excel files with multiple sheets. It supports setting styles on individual cells and offers two export methods: `StreamWriter` and `Memory`. The package utilizes the `github.com/xuri/excelize/v2` library to handle the Excel file operations.

## Installation

To install the package, use the following command:

```sh
go get -u github.com/xelarion/excel-exporter
```

## Usage

### Creating a New ExcelExporter

To create a new `ExcelExporter` instance, specify the file name and whether to use the `StreamWriter` method for exporting:

```go
import "github.com/xelarion/excel-exporter"

exporter := excel_exporter.NewExcelExporter("example.xlsx", true) // Use true for StreamWriter
```

### Defining Sheet Data

Define the data for each sheet using the `SheetData` structure. Implement a `RowDataFunc` function that returns the data for each row:

```go
func rowDataFunc() *excel_exporter.Row {
    cells := []excelize.Cell{
        {Value: "Sample Data 1"},
        {Value: "Sample Data 2"},
    }

    return &excel_exporter.Row{
        Cells: cells,
        MergeCells: []excel_exporter.MergeCell{
            {HCell: "A1", VCell: "B1"},
        },
        RowOpts: []excelize.RowOpts{
            {Height: 20},
        },
    }
}

sheetData := excel_exporter.SheetData{
    Name:    "Sheet1",
    RowFunc: rowDataFunc,
}
```

### Exporting the Excel File

To export the Excel file, call the `Export` method with the sheet data:

```go
sheets := []excel_exporter.SheetData{sheetData}

if err := exporter.Export(sheets); err != nil {
    fmt.Println("Failed to export Excel file:", err)
} else {
    fmt.Println("Excel file exported successfully.")
}
```

### Full Example

```go
package main

import (
    "fmt"
    "github.com/xelarion/excel-exporter"
    "github.com/xuri/excelize/v2"
)

func generateRowData(exporter *excel_exporter.ExcelExporter) *excel_exporter.Row {
    cells := make([]excelize.Cell, 20)
    for i := 0; i < 20; i++ {
        style, _ := exporter.File.NewStyle(&excelize.Style{
            Font: &excelize.Font{Color: "777777", Size: 12},
            Alignment: &excelize.Alignment{
                Horizontal: "center",
                Vertical:   "center",
            },
        })
        cells[i] = excelize.Cell{
            StyleID: style,
            Formula: "",
            Value:   fmt.Sprintf("Data %d", i),
        }
    }

    return &excel_exporter.Row{
        Cells: cells,
        MergeCells: []excel_exporter.MergeCell{
            {HCell: "A1", VCell: "B1"},
        },
        RowOpts: []excelize.RowOpts{
            {Height: 20},
        },
    }
}

func generateLargeData(exporter *excel_exporter.ExcelExporter, rowCount int) excel_exporter.RowDataFunc {
    currentRow := 0
    return func() *excel_exporter.Row {
        if currentRow >= rowCount {
            return nil
        }
        currentRow++
        return generateRowData(exporter)
    }
}

func main() {
    exporter := excel_exporter.NewExcelExporter("example.xlsx", true) // Use StreamWriter

    sheetData1 := excel_exporter.SheetData{
        Name:    "Sheet1",
        RowFunc: generateLargeData(exporter, 1000), // 1000 rows
    }
    sheetData2 := excel_exporter.SheetData{
        Name:    "Sheet2",
        RowFunc: generateLargeData(exporter, 500),  // 500 rows
    }

    sheets := []excel_exporter.SheetData{sheetData1, sheetData2}

    if err := exporter.Export(sheets); err != nil {
        fmt.Println("Failed to export Excel file:", err)
    } else {
        fmt.Println("Excel file exported successfully.")
    }
}
```

### Testing

The package includes tests to verify the functionality of exporting large datasets with both `StreamWriter` and in-memory methods. The tests record the duration, CPU usage, and memory usage.

```
go test -v ./...
```

## License

This project is licensed under the MIT License.
