package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"math"
	"os"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Structs for JSON export
type Discrepancy struct {
	SerialNo   int     `json:"serial_no"`
	StudentID  string  `json:"student_id"`
	Expected   float64 `json:"expected"`
	Calculated float64 `json:"calculated"`
}

type Averages struct {
	ComponentAverages map[string]float64 `json:"component_averages"`
	BranchAverages    map[string]float64 `json:"branch_averages"`
}

type Rank struct {
	Component string `json:"component"`
	TopRanks  []struct {
		Rank   int     `json:"rank"`
		Emplid string  `json:"emplid"`
		Marks  float64 `json:"marks"`
	} `json:"top_ranks"`
}

type ExportData struct {
	Discrepancies []Discrepancy `json:"discrepancies"`
	Averages      Averages      `json:"averages"`
	Ranks         []Rank        `json:"ranks"`
}

func isEmptyRow(row []string) bool {
	for _, cell := range row {
		if cell != "" {
			return false // Found a non-empty cell
		}
	}
	return true // All cells are empty
}

func checkDiscrepancies(validRows [][]string, tolerance float64) {
	for i := 1; i < len(validRows); i++ {
		sum := 0.0
		for j := 4; j <= 9; j++ {
			if j == 8 {
				continue
			}
			value, err := strconv.ParseFloat(validRows[i][j], 64)
			if err != nil {
				fmt.Println("Conversion error:", err)
				continue // Skip this cell if conversion fails
			}
			sum += value
		}
		expectedSum, err := strconv.ParseFloat(validRows[i][10], 64)
		if err != nil {
			fmt.Println("Conversion error:", err)
		} else if math.Abs(sum-expectedSum) > tolerance {
			fmt.Printf("There is a discrepancy in the sum of the student with Sl No: %d and student ID: %s\n", i, validRows[i][3])
			validRows[i][10] = strconv.FormatFloat(sum, 'f', 2, 64)
		}
	}
}

func calculateAverages(validRows [][]string) {
	for i := 4; i < 11; i++ {
		sum := 0.0
		average := 0.0
		for j := 1; j < len(validRows); j++ {
			value, err := strconv.ParseFloat(validRows[j][i], 32)
			if err != nil {
				fmt.Println("Conversion error:", err)
				continue // Skip this cell if conversion fails
			}
			sum += value
		}
		average = sum / float64((len(validRows) - 1))
		fmt.Printf("The average for %s is: %.2f\n", validRows[0][i], average)
	}
}

func computeBranchAverages(validRows [][]string) {
	branches := map[string]struct {
		sum float64
		len float64
	}{
		"2024A3": {0, 0},
		"2024A4": {0, 0},
		"2024A5": {0, 0},
		"2024A7": {0, 0},
		"2024A8": {0, 0},
		"2024AA": {0, 0},
		"2024AD": {0, 0},
	}

	for i := 1; i < len(validRows); i++ {
		for prefix := range branches {
			if strings.HasPrefix(validRows[i][3], prefix) {
				value, err := strconv.ParseFloat(validRows[i][10], 64)
				if err != nil {
					fmt.Println("Conversion error:", err)
					break
				}
				b := branches[prefix]
				b.sum += value
				b.len++
				branches[prefix] = b
				break
			}
		}
	}

	for prefix, data := range branches {
		if data.len > 0 {
			fmt.Printf("Average for branch %s: %.2f\n", prefix, data.sum/data.len)
		}
	}
}

func top3Ranks(validRows [][]string) {
	var Rank_1, Rank_2, Rank_3 int
	for i := 4; i <= 10; i++ {
		maxMarks, secondMaxMarks, thirdMaxMarks := -1.0, -1.0, -1.0
		Rank_1, Rank_2, Rank_3 = -1, -1, -1

		for j := 1; j < len(validRows)-1; j++ {
			marks, err := strconv.ParseFloat(validRows[j][i], 64)
			if err != nil {
				fmt.Println("Conversion error:", err)
				continue // Skip this cell if conversion fails
			}

			if marks > maxMarks {
				// Shift ranks down
				thirdMaxMarks, Rank_3 = secondMaxMarks, Rank_2
				secondMaxMarks, Rank_2 = maxMarks, Rank_1
				maxMarks, Rank_1 = marks, j

			} else if marks > secondMaxMarks {
				thirdMaxMarks, Rank_3 = secondMaxMarks, Rank_2
				secondMaxMarks, Rank_2 = marks, j

			} else if marks > thirdMaxMarks {
				thirdMaxMarks, Rank_3 = marks, j
			}
		}
		fmt.Printf("\nCP %s Rankings:\n", validRows[0][i])
		fmt.Printf("1st: Emplid: %s\tMarks: %s\n", validRows[Rank_1][2], validRows[Rank_1][i])
		fmt.Printf("2nd: Emplid: %s\tMarks: %s\n", validRows[Rank_2][2], validRows[Rank_2][i])
		fmt.Printf("3rd: Emplid: %s\tMarks: %s\n", validRows[Rank_3][2], validRows[Rank_3][i])

	}
}

func exportToJSON(discrepancies []Discrepancy, averages map[string]float64, branchAverages map[string]float64, ranks []Rank, filename string) error {
	data := ExportData{
		Discrepancies: discrepancies,
		Averages: Averages{
			ComponentAverages: averages,
			BranchAverages:    branchAverages,
		},
		Ranks: ranks,
	}

	file, err := os.Create(filename)
	if err != nil {
		return err
	}
	defer file.Close()

	encoder := json.NewEncoder(file)
	encoder.SetIndent("", "  ")
	return encoder.Encode(data)
}

func main() {

	if len(os.Args) < 2 {
		fmt.Println("Please provide the XLSX file path as a command-line argument.")
		return
	}

	filePath := os.Args[1]

	exportFlag := flag.Bool("export", false, "Export summary to JSON file")
	outputFile := flag.String("output", "report.json", "Output JSON filename")
	flag.Parse()

	if len(flag.Args()) < 1 {
		fmt.Println("Usage: go run main.go <file.xlsx> [--export] [--output=filename.json]")
		return
	}

	f, err := excelize.OpenFile(filePath)
	if err != nil {
		fmt.Println(err)
		return
	}
	if !strings.HasSuffix(filePath, ".xlsx") {
		fmt.Println("Error: Provided file is not an XLSX file.")
		return
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// Get all sheet names
	sheets := f.GetSheetList()
	sheetName := sheets[0]
	// Get all the rows in the Sheet1.
	rows, err := f.GetRows(sheetName)
	if err != nil {
		fmt.Println(err)
		return
	}
	var validRows [][]string
	for _, row := range rows {
		if !isEmptyRow(row) {
			validRows = append(validRows, row)
		}
	}
	if len(validRows[0]) < 11 {
		fmt.Println("Error: The Excel file has fewer columns than expected.")
		return
	}
	tolerance := 1e-2
	checkDiscrepancies(validRows, tolerance) // Checks for discrepancy within the totalling

	fmt.Println("\nComponent-wise Averages:")
	calculateAverages(validRows) // Calculates average per component

	fmt.Println("\nBranch-wise Averages:")
	computeBranchAverages(validRows) // Calculates branch-wise average

	fmt.Println("\nTOP 3 RANKS:")
	top3Ranks(validRows) // Gives out top 3 ranks per component

}
