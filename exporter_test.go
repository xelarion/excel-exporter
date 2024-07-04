package excel_exporter

import (
	"fmt"
	"os"
	"runtime"
	"runtime/pprof"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

func TestExportWithStreamWriter(t *testing.T) {
	exporter := NewExcelExporter("test_streamwriter.xlsx", true)

	sheetData1 := SheetData{
		Name:    "SheetA",
		RowFunc: generateLargeData(20000), // 20k rows
	}
	sheetData2 := SheetData{
		Name:    "SheetB",
		RowFunc: generateLargeData(2000), // 2k rows
	}

	sheets := []SheetData{sheetData1, sheetData2}

	// Start CPU profiling
	cpuProfile, err := os.Create("cpu_profile_streamwriter.prof")
	if err != nil {
		t.Fatalf("could not create CPU profile: %v", err)
	}
	pprof.StartCPUProfile(cpuProfile)
	defer pprof.StopCPUProfile()

	start := time.Now()
	if err := exporter.Export(sheets); err != nil {
		t.Fatalf("Failed to export Excel file with StreamWriter: %v", err)
	}
	duration := time.Since(start)
	t.Logf("Export with StreamWriter took %v", duration)

	// Record memory stats after export
	recordMemoryStats(t, "StreamWriter")

	pprof.StopCPUProfile()
	cpuProfile.Close()

	// Start memory profiling
	memProfile, err := os.Create("mem_profile_streamwriter.prof")
	if err != nil {
		t.Fatalf("could not create memory profile: %v", err)
	}
	defer memProfile.Close()
	runtime.GC() // get up-to-date statistics
	if err := pprof.WriteHeapProfile(memProfile); err != nil {
		t.Fatalf("could not write memory profile: %v", err)
	}
}

func TestExportWithMemory(t *testing.T) {
	exporter := NewExcelExporter("test_memory.xlsx", false)

	sheetData1 := SheetData{
		Name:    "Sheet1",
		RowFunc: generateLargeData(20000), // 20k rows
	}
	sheetData2 := SheetData{
		Name:    "Sheet2",
		RowFunc: generateLargeData(2000), // 2k rows
	}

	sheets := []SheetData{sheetData1, sheetData2}

	// Start CPU profiling
	cpuProfile, err := os.Create("cpu_profile_memory.prof")
	if err != nil {
		t.Fatalf("could not create CPU profile: %v", err)
	}
	pprof.StartCPUProfile(cpuProfile)
	defer pprof.StopCPUProfile()

	start := time.Now()
	if err := exporter.Export(sheets); err != nil {
		t.Fatalf("Failed to export Excel file with memory: %v", err)
	}
	duration := time.Since(start)
	t.Logf("Export with memory took %v", duration)

	// Record memory stats after export
	recordMemoryStats(t, "Memory")

	pprof.StopCPUProfile()
	cpuProfile.Close()

	// Start memory profiling
	memProfile, err := os.Create("mem_profile_memory.prof")
	if err != nil {
		t.Fatalf("could not create memory profile: %v", err)
	}
	defer memProfile.Close()
	runtime.GC() // get up-to-date statistics
	if err := pprof.WriteHeapProfile(memProfile); err != nil {
		t.Fatalf("could not write memory profile: %v", err)
	}
}

func TestExportWithStreamWriterUseChannel(t *testing.T) {
	start := time.Now()

	exporter := NewExcelExporter("test_streamwriter_channel.xlsx", true)

	sheetNames := []string{"SheetA", "SheetB"}
	sheets := make([]SheetData, len(sheetNames))
	for i, name := range sheetNames {
		sheets[i] = SheetData{
			Name:    name,
			RowFunc: UseRowChan(addRowsToChanFunc(exporter, name)),
		}
	}

	if err := exporter.Export(sheets); err != nil {
		t.Fatalf("Failed to export Excel file with StreamWriter channel: %v", err)
	}

	duration := time.Since(start)
	t.Logf("Export with StreamWriter channel took %v", duration)
}

func addRowsToChanFunc(exporter *ExcelExporter, sheetName string) func(dataCh chan Row) {
	return func(dataCh chan Row) {
		titleStyle, _ := exporter.File.NewStyle(
			&excelize.Style{
				Font:      &excelize.Font{Color: "777777", Size: 14},
				Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
			},
		)

		// Set column width
		if exporter.UseStreamWriter {
			_ = exporter.StreamWriter.SetColWidth(1, 3, 30)
		} else {
			_ = exporter.File.SetColWidth(exporter.CurrentSheetName, "A", "C", 30)
		}

		// Merge cells
		dataCh <- Row{
			Cells: []excelize.Cell{
				{Value: "MergedTitle 1"},
				{Value: ""},
				{Value: "MergedTitle 2"},
			},
			MergeCells: []MergeCell{
				{HCell: "A1", VCell: "B1"},
			},
			RowOpts: []excelize.RowOpts{
				{Height: 20, StyleID: titleStyle},
			},
		}

		// Title
		dataCh <- Row{
			Cells: []excelize.Cell{
				{Value: "Title 1", StyleID: titleStyle},
				{Value: "Title 2", StyleID: titleStyle},
				{Value: "Title 3", StyleID: titleStyle},
			},
		}

		// query data and send to channel
		for i := 0; i < 10; i++ {
			dataCh <- Row{
				Cells: []excelize.Cell{
					{Value: fmt.Sprintf("%s-%d-1", sheetName, i)},
					{Value: fmt.Sprintf("%s-%d-2", sheetName, i)},
					{Value: fmt.Sprintf("%s-%d-3", sheetName, i)},
				},
			}
		}

	}
}

func generateLargeData(rowCount int) RowDataFunc {
	currentRow := 0
	return func() Row {
		if currentRow >= rowCount {
			return Row{}
		}
		currentRow++
		return Row{
			Cells: []excelize.Cell{
				{Value: fmt.Sprintf("a%d", currentRow)},
				{Value: fmt.Sprintf("b%d", currentRow)},
				{Value: fmt.Sprintf("c%d", currentRow)},
			},
			MergeCells: nil,
			RowOpts:    nil,
		}
	}
}

func recordMemoryStats(t *testing.T, prefix string) {
	var m runtime.MemStats
	runtime.ReadMemStats(&m)
	t.Logf("%s - Memory Usage: Alloc = %v MiB, TotalAlloc = %v MiB, Sys = %v MiB, NumGC = %v",
		prefix, m.Alloc/1024/1024, m.TotalAlloc/1024/1024, m.Sys/1024/1024, m.NumGC)
}
