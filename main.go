package main

import (
	"context"
	"fmt"
	"log"

	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

func main() {
	// Load Service Account credentials
	ctx := context.Background()
	credentialsFile := "aufguss-dm-169088e9f544.json" // Replace with your JSON file
	credentialsFile = "aufguss-scoring-6ad4a3b3519f.json"
	sheetsService, err := sheets.NewService(ctx, option.WithCredentialsFile(credentialsFile))
	if err != nil {
		log.Fatalf("Unable to create Sheets client: %v", err)
	}

	driveService, err := drive.NewService(ctx, option.WithCredentialsFile(credentialsFile))
	if err != nil {
		log.Fatalf("Unable to create Drive client: %v", err)
	}

	// ID of the source Google Sheet
	sourceSpreadsheetID := "1VZGcUjgKaYv4yo5LEeHsUX-xLQDwfT9b6S6aO4uxVds" // Replace with your source sheet ID

	// ID of the parent folder where the new folder will be created
	parentFolderID := "166oePFGeddWcho14g5bXULF6B6Bb5vO6" // Replace with the ID of the parent folder

	// 1. Create a new empty folder
	newFolder := &drive.File{
		Name:     "Aufguss VM 2025 Classic",
		MimeType: "application/vnd.google-apps.folder",
		Parents:  []string{parentFolderID},
	}
	createdFolder, err := driveService.Files.Create(newFolder).Do()
	if err != nil {
		log.Fatalf("Unable to create folder: %v", err)
	}
	newFolderID := createdFolder.Id
	fmt.Printf("Created new folder with ID: %s\n", newFolderID)

	// 2. Get the name of the source spreadsheet
	sourceFile, err := driveService.Files.Get(sourceSpreadsheetID).Fields("name").Do()
	if err != nil {
		log.Fatalf("Unable to retrieve source spreadsheet name: %v", err)
	}
	newSpreadsheetName := fmt.Sprintf("%s [copy]", sourceFile.Name)

	// 3. Copy the Google Sheet using Drive API
	newFile := &drive.File{
		Name:     newSpreadsheetName,
		MimeType: "application/vnd.google-apps.spreadsheet",
		Parents:  []string{newFolderID},
	}
	copiedFile, err := driveService.Files.Copy(sourceSpreadsheetID, newFile).Do()
	if err != nil {
		log.Fatalf("Unable to copy spreadsheet: %v", err)
	}
	copiedSpreadsheetID := copiedFile.Id
	fmt.Printf("Copied spreadsheet ID: %s\n", copiedSpreadsheetID)

	// 4. Copy the "Board" sheet to "Board 2" within the copied spreadsheet
	// Get the source spreadsheet details
	sourceSpreadsheet, err := sheetsService.Spreadsheets.Get(copiedSpreadsheetID).Fields("sheets(properties(sheetId,title))").Do()
	if err != nil {
		log.Fatalf("Unable to get spreadsheet details: %v", err)
	}

	// Find the "Board" sheet ID
	var boardSheetID int64
	for _, sheet := range sourceSpreadsheet.Sheets {
		if sheet.Properties.Title == "Board" {
			boardSheetID = sheet.Properties.SheetId
			break
		}
	}
	if boardSheetID == 0 {
		log.Fatalf("Sheet named 'Board' not found in the spreadsheet.")
	}

	// Duplicate the "Board" sheet
	duplicateRequest := &sheets.DuplicateSheetRequest{
		SourceSheetId: boardSheetID,
		NewSheetName:  "Board 2",
	}
	request := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{
			{DuplicateSheet: duplicateRequest},
		},
	}
	_, err = sheetsService.Spreadsheets.BatchUpdate(copiedSpreadsheetID, request).Do()
	if err != nil {
		log.Fatalf("Unable to duplicate sheet: %v", err)
	}
	fmt.Println("Sheet 'Board' duplicated to 'Board 2'.")

	// 5. Write a value to a cell in the copied sheet
	writeRange := "Board 2!B1" // Adjust as necessary
	value := "My Name!"
	values := &sheets.ValueRange{
		Values: [][]interface{}{{value}},
	}
	_, err = sheetsService.Spreadsheets.Values.Update(copiedSpreadsheetID, writeRange, values).
		ValueInputOption("RAW").Do()
	if err != nil {
		log.Fatalf("Unable to write to sheet: %v", err)
	}
	fmt.Println("Written value to cell A1 in the copied sheet")

	// 6. Insert a row in the copied sheet
	insertRange := "Board 2!A9:A9" // Adjust as necessary
	newRow := &sheets.ValueRange{
		Values: [][]interface{}{{"", "(New Row)"}},
	}
	_, err = sheetsService.Spreadsheets.Values.Append(copiedSpreadsheetID, insertRange, newRow).
		ValueInputOption("RAW").InsertDataOption("INSERT_ROWS").Do()
	if err != nil {
		log.Fatalf("Unable to insert row: %v", err)
	}
	fmt.Println("Inserted new row in the copied sheet")
}
