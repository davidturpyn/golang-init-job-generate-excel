package main

import (
	"JOB_GenerateExcel/database"
	"JOB_GenerateExcel/models"
	"JOB_GenerateExcel/services"
	"fmt"
	"log"
	"time"
)

func main() {

	// Start log
	fmt.Println("===========================================================================")
	fmt.Println("GenerateExcel Application START " + time.Now().Format("02-01-2006 03:04:05 PM"))
	fmt.Println("===========================================================================")

	// Try-catch block in Go style (using error handling)
	defer func() {
		if r := recover(); r != nil {
			log.Printf("Recovered from panic: %v", r)
			fmt.Println("===========================================================================")
			fmt.Println("GenerateExcel Application END (WITH ERROR) " + time.Now().Format("02-01-2006 03:04:05 PM"))
			fmt.Println("===========================================================================")
		}
	}()

	// Connect to the database
	if err := database.Connect(); err != nil {
		log.Fatalf("Database connection failed: %v", err)
	}
	fmt.Println("Connected to the database.")

	// Migrate the User model (if needed)
	if err := database.DB.AutoMigrate(&models.User{}); err != nil {
		log.Fatalf("Failed to migrate User model: %v", err)
	}
	fmt.Println("Database migration completed.")

	// Sample data insertion (simulating data fetch as in C# example)
	// users := []models.User{
	// 	{Name: "John Doe", Email: "john@example.com", Age: 30},
	// 	{Name: "Jane Smith", Email: "jane@example.com", Age: 25},
	// }
	// for _, user := range users {
	// 	database.DB.Create(&user)
	// }
	// fmt.Println("Inserted sample user data.")

	// Generate the Excel file
	outputPath := "output/users.xlsx"
	if err := services.GenerateExcel(outputPath); err != nil {
		log.Printf("Error generating Excel file: %v", err)
		fmt.Println("===========================================================================")
		fmt.Println("GenerateExcel Application END (WITH ERROR) " + time.Now().Format("02-01-2006 03:04:05 PM"))
		fmt.Println("===========================================================================")
		return
	}

	// Log successful completion
	fmt.Println("Excel file generated successfully at:", outputPath)

	// Simulate email sending (you can integrate with your actual SMTP email sending)
	sendEmailSimulation()

	// Final log
	fmt.Println("===========================================================================")
	fmt.Println("GenerateExcel Application END " + time.Now().Format("02-01-2006 03:04:05 PM"))
	fmt.Println("===========================================================================")
}

func sendEmailSimulation() {
	// Example of email sending logic (can be replaced with actual email logic)
	fmt.Println("Simulating sending email...")

	// Email content (this would be more sophisticated in a real application)
	emailBody := "<html><body><h1>Job Status Notification</h1><p>Excel file has been successfully generated.</p></body></html>"

	// Log email sending status
	fmt.Println("Email body prepared:", emailBody)

	// Simulate email sent
	fmt.Println("Email sent successfully.")
}
