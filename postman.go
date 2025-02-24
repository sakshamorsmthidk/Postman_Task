package main

import (
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func isEmptyRow(row []string) bool {
	for _, cell := range row {
		if cell != "" {
			return false // Found a non-empty cell
		}
	}
	return true // All cells are empty
}

func main() {

	if len(os.Args) < 2 {
		fmt.Println("Please provide the XLSX file path as a command-line argument.")
		return
	}

	filePath := os.Args[1]

	f, err := excelize.OpenFile(filePath)
	if err != nil {
		fmt.Println(err)
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
	fmt.Println("📄 Available sheets:", sheets)
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
	// Checks for discrepancy within the totalling
	for i := 1; i < len(validRows); i++ {
		total := 0.0
		pre_compre, err := strconv.ParseFloat(validRows[i][8], 32)
		if err != nil {
			fmt.Println("Conversion error:", err)
			continue // Skip this cell if conversion fails
		}
		compre, err := strconv.ParseFloat(validRows[i][9], 32)
		if err != nil {
			fmt.Println("Conversion error:", err)
			continue // Skip this cell if conversion fails
		}
		total = pre_compre + compre
		expectedSum, err := strconv.ParseFloat(validRows[i][10], 32)
		if err != nil {
			fmt.Println("Conversion error:", err)
		} else if total != expectedSum {
			fmt.Printf("There is a discrepancy in the sum of the student with Sl No: %d and student ID: %s\n", i, validRows[i][3])
		}

	}
	// Gives average per eval component
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
		fmt.Printf("The average for %s is: %f\n", validRows[0][i], average)
	}

	var len_A3, sum_A3, len_A4, sum_A4, len_A5, sum_A5, len_A7, sum_A7, len_A8, sum_A8, len_AA, sum_AA, len_AD, sum_AD float64

	for i := 1; i < len(validRows); i++ {
		switch {

		case strings.HasPrefix(validRows[i][3], "2024A3"):
			value, err := strconv.ParseFloat(validRows[i][10], 32)
			if err != nil {
				fmt.Println("Conversion error:", err)
				continue // Skip this cell if conversion fails
			}
			sum_A3 += value
			len_A3++
		case strings.HasPrefix(validRows[i][3], "2024A4"):
			value, err := strconv.ParseFloat(validRows[i][10], 32)
			if err != nil {
				fmt.Println("Conversion error:", err)
				continue // Skip this cell if conversion fails
			}
			sum_A4 += value
			len_A4++
		case strings.HasPrefix(validRows[i][3], "2024A5"):
			value, err := strconv.ParseFloat(validRows[i][10], 32)
			if err != nil {
				fmt.Println("Conversion error:", err)
				continue // Skip this cell if conversion fails
			}
			sum_A5 += value
			len_A5++
		case strings.HasPrefix(validRows[i][3], "2024A7"):
			value, err := strconv.ParseFloat(validRows[i][10], 32)
			if err != nil {
				fmt.Println("Conversion error:", err)
				continue // Skip this cell if conversion fails
			}
			sum_A7 += value
			len_A7++
		case strings.HasPrefix(validRows[i][3], "2024A8"):
			value, err := strconv.ParseFloat(validRows[i][10], 32)
			if err != nil {
				fmt.Println("Conversion error:", err)
				continue // Skip this cell if conversion fails
			}
			sum_A8 += value
			len_A8++
		case strings.HasPrefix(validRows[i][3], "2024AA"):
			value, err := strconv.ParseFloat(validRows[i][10], 32)
			if err != nil {
				fmt.Println("Conversion error:", err)
				continue // Skip this cell if conversion fails
			}
			sum_AA += value
			len_AA++
		case strings.HasPrefix(validRows[i][3], "2024AD"):
			value, err := strconv.ParseFloat(validRows[i][10], 32)
			if err != nil {
				fmt.Println("Conversion error:", err)
				continue // Skip this cell if conversion fails
			}
			sum_AD += value
			len_AD++
		}
	}
	fmt.Printf("The Branch average for EEE is: %f\n", sum_A3/len_A3)
	fmt.Printf("The Branch average for Mechanical is: %f\n", sum_A4/len_A4)
	fmt.Printf("The Branch average for B.Pharma is: %f\n", sum_A5/len_A5)
	fmt.Printf("The Branch average for CSE is: %f\n", sum_A7/len_A7)
	fmt.Printf("The Branch average for ENI is: %f\n", sum_A8/len_A8)
	fmt.Printf("The Branch average for ECE is: %f\n", sum_AA/len_AA)
	fmt.Printf("The Branch average for MnC is: %f\n", sum_AD/len_AD)

}
