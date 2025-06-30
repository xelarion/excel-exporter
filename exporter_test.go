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
	exporter := New("test_streamwriter.xlsx", true)

	sheetData1 := SheetData{
		Name:    "SheetA",
		RowFunc: generateLargeData("SheetA", 20000), // 20k rows
	}
	sheetData2 := SheetData{
		Name:    "SheetB",
		RowFunc: generateLargeData("SheetB", 2000), // 2k rows
	}
	sheetData3 := SheetData{
		Name:    "SheetC",
		RowFunc: generateLargeData("SheetC", SheetMaxRows+1),
	}

	sheets := []SheetData{sheetData1, sheetData2, sheetData3}

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
	exporter := New("test_memory.xlsx", false)

	sheetData1 := SheetData{
		Name:    "Sheet1",
		RowFunc: generateLargeData("Sheet1", 20000), // 20k rows
	}
	sheetData2 := SheetData{
		Name:    "Sheet2",
		RowFunc: generateLargeData("Sheet2", 2000), // 2k rows
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

	exporter := New("test_streamwriter_channel.xlsx", true)

	sheetNames := []string{"SheetA", "SheetB"}
	sheets := make([]SheetData, len(sheetNames))
	for i, name := range sheetNames {
		sheets[i] = SheetData{
			Name:    name,
			RowFunc: UseRowChan(queryDataToChannelFunc(exporter, name)),
			InitFunc: func(exporter *Exporter) (err error) {
				if exporter.UseStreamWriter {
					if err = exporter.StreamWriter.SetColWidth(1, 3, 30); err != nil {
						return err
					}
				} else {
					if err = exporter.File.SetColWidth(exporter.CurrentSheet, "A", "C", 30); err != nil {
						return err
					}
				}
				return nil
			},
		}
	}

	if err := exporter.Export(sheets); err != nil {
		t.Fatalf("Failed to export Excel file with StreamWriter channel: %v", err)
	}

	duration := time.Since(start)
	t.Logf("Export with StreamWriter channel took %v", duration)
}

func queryDataToChannelFunc(exporter *Exporter, sheetName string) func(dataCh chan Row) error {
	return func(dataCh chan Row) error {
		titleStyle, err := exporter.File.NewStyle(
			&excelize.Style{
				Font:      &excelize.Font{Color: "777777", Size: 14},
				Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
			},
		)
		if err != nil {
			return err
		}

		// Simulate querying data from the database and sending to channel
		rowIndex := 0
		for i := 0; i < (SheetMaxRows*2)+3; i++ {
			if i%(SheetMaxRows) == 0 {
				dataCh <- Row{
					Cells: []excelize.Cell{
						{Value: "MergedTitle 1", StyleID: titleStyle},
						{Value: "", StyleID: titleStyle},
						{Value: "MergedTitle 2", StyleID: titleStyle},
					},
					MergeCells: []MergeCell{
						{TopLeftCell: "A1", BottomRightCell: "B1"},
					},
					RowOpts: []excelize.RowOpts{
						{Height: 20, StyleID: titleStyle},
					},
				}
				continue
			}
			if i%(SheetMaxRows) == 1 {
				dataCh <- Row{
					Cells: []excelize.Cell{
						{Value: "Sub Title 1", StyleID: titleStyle},
						{Value: "Sub Title 2", StyleID: titleStyle},
						{Value: "Sub Title 3", StyleID: titleStyle},
					},
				}
				continue
			}
			rowIndex++

			dataCh <- NewRow(
				fmt.Sprintf("%s-%d-1", sheetName, rowIndex),
				fmt.Sprintf("%s-%d-2", sheetName, rowIndex),
				fmt.Sprintf("%s-%d-3", sheetName, rowIndex),
			)
		}

		return nil
	}
}

func generateLargeData(sheetName string, rowCount int) RowDataFunc {
	return func(rowNumber int) (Row, error) {
		if rowNumber > rowCount {
			return Row{}, nil
		}

		return NewRow(
			fmt.Sprintf("%s-a%d", sheetName, rowNumber),
			fmt.Sprintf("%s-b%d", sheetName, rowNumber),
			fmt.Sprintf("%s-c%d", sheetName, rowNumber),
		), nil
	}
}

func recordMemoryStats(t *testing.T, prefix string) {
	var m runtime.MemStats
	runtime.ReadMemStats(&m)
	t.Logf("%s - Memory Usage: Alloc = %v MiB, TotalAlloc = %v MiB, Sys = %v MiB, NumGC = %v",
		prefix, m.Alloc/1024/1024, m.TotalAlloc/1024/1024, m.Sys/1024/1024, m.NumGC)
}
