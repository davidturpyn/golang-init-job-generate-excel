package services

import (
	"JOB_GenerateExcel/database"
	"JOB_GenerateExcel/models"
	"fmt"
	"os"
	"time"

	"github.com/xuri/excelize/v2"
)

// GenerateExcel generates an Excel file containing users created today
func GenerateExcel(filePath string) error {
	// Get the current date (no time part, only date)
	today := time.Now().Format("2006-01-02") // Format as YYYY-MM-DD

	// Fetch users from the database where CreatedAt matches today's date
	var users []models.User
	if err := database.DB.Where("DATE(created_at) = ?", today).Find(&users).Error; err != nil {
		return fmt.Errorf("failed to fetch users: %w", err)
	}

	// Create a new Excel file
	f := excelize.NewFile()

	// Create a sheet
	sheetName := "Users"
	f.SetSheetName("Sheet1", sheetName)

	// Define styles
	headerStyle, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Color: "#FFFFFF", Size: 12},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#4CAF50"}, Pattern: 1},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})

	dataStyle, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "left", Vertical: "center"},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})

	// Set headers
	headers := []string{"ID", "Name", "Email", "Age", "Created At"}
	for i, header := range headers {
		cell := fmt.Sprintf("%s1", string('A'+i))
		f.SetCellValue(sheetName, cell, header)
		f.SetCellStyle(sheetName, cell, cell, headerStyle)
	}

	// Populate data
	for row, user := range users {
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", row+2), user.ID)
		f.SetCellValue(sheetName, fmt.Sprintf("B%d", row+2), user.Name)
		f.SetCellValue(sheetName, fmt.Sprintf("C%d", row+2), user.Email)
		f.SetCellValue(sheetName, fmt.Sprintf("D%d", row+2), user.Age)
		f.SetCellValue(sheetName, fmt.Sprintf("E%d", row+2), user.CreatedAt.Format("2006-01-02 15:04:05"))

		// Apply styles to data rows
		for col := 0; col < len(headers); col++ {
			cell := fmt.Sprintf("%s%d", string('A'+col), row+2)
			f.SetCellStyle(sheetName, cell, cell, dataStyle)
		}
	}

	// Adjust column widths
	columnWidths := map[string]float64{
		"A": 5,  // ID
		"B": 20, // Name
		"C": 30, // Email
		"D": 5,  // Age
		"E": 20, // Created At
	}
	for col, width := range columnWidths {
		f.SetColWidth(sheetName, col, col, width)
	}

	// Ensure the folder exists
	folder := filePath[:len(filePath)-len("/users.xlsx")]
	if err := os.MkdirAll(folder, os.ModePerm); err != nil {
		return fmt.Errorf("failed to create directory: %w", err)
	}

	// Save the file to the specified folder
	if err := f.SaveAs(filePath); err != nil {
		return fmt.Errorf("failed to save Excel file: %w", err)
	}

	return nil
}
