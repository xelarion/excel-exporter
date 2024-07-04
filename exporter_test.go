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

func generateRowData(exporter *ExcelExporter) *Row {
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

	return &Row{
		Cells: cells,
		MergeCells: []MergeCell{
			{HCell: "A1", VCell: "B1"},
		},
		RowOpts: []excelize.RowOpts{
			{Height: 20},
		},
	}
}

func generateLargeData(exporter *ExcelExporter, rowCount int) RowDataFunc {
	currentRow := 0
	return func() *Row {
		if currentRow >= rowCount {
			return nil
		}
		currentRow++
		return generateRowData(exporter)
	}
}

func recordMemoryStats(t *testing.T, prefix string) {
	var m runtime.MemStats
	runtime.ReadMemStats(&m)
	t.Logf("%s - Memory Usage: Alloc = %v MiB, TotalAlloc = %v MiB, Sys = %v MiB, NumGC = %v",
		prefix, m.Alloc/1024/1024, m.TotalAlloc/1024/1024, m.Sys/1024/1024, m.NumGC)
}

func TestExportWithStreamWriter(t *testing.T) {
	exporter := NewExcelExporter("test_streamwriter.xlsx", true)

	sheetData1 := SheetData{
		Name:    "SheetA",
		RowFunc: generateLargeData(exporter, 20000), // 20k rows
	}
	sheetData2 := SheetData{
		Name:    "SheetB",
		RowFunc: generateLargeData(exporter, 2000), // 2k rows
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
		RowFunc: generateLargeData(exporter, 20000), // 20k rows
	}
	sheetData2 := SheetData{
		Name:    "Sheet2",
		RowFunc: generateLargeData(exporter, 2000), // 2k rows
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
