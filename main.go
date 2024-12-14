package main

import (
	"bytes"
	_ "embed"
	"image/color"
	"os"
	"path/filepath"
	"strings"

	"context"
	"fmt"
	"image"

	log "github.com/s00500/env_logger"

	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/container"
)

// Embed the logo file
//
//go:embed embedded/logo.png
var logoPNG []byte

var competitionsDir string = "competitions"

// Competition struct
type Competition struct {
	Name          string        `json:"name"`
	SourceSheetID string        `json:"source_sheet_id"`
	Jury          []*Juror      `json:"jury"`
	Contestants   []*Contestant `json:"contestants"`
}

type Juror struct {
	Name   string `json:"name"`
	Weight int    `json:"weight"`
}

type Contestant struct {
	Name string `json:"name"`
}

// Directory to store competition files
const appWidth = 800

var dataDir string

func main() {

	// Setting the path.
	execPath, err := getExecutableDir()
	if err != nil {
		log.Fatalf("Failed to get executable directory: %v", err)
	}

	// For a bundled application, it will be the absolute path, and we split by the application name.
	exePathParts := strings.Split(execPath, "Aufguss Scoring Generator.app")
	if len(exePathParts) != 2 {
		exePathParts[0] = "" // Otherwise (when running in development) we will use relative path
		//log.Fatalf("Aufguss Scoring Generator.app was not in path")
	}

	dataDir = filepath.Join(exePathParts[0], "competitions")
	if _, err := os.Stat(dataDir); os.IsNotExist(err) {
		if err := os.MkdirAll(dataDir, 0755); err != nil {
			log.Fatalf("Failed to create competitions directory: %v", err)
		}
	}
	//log.Println(dataDir)

	// Initialize app and window
	myApp := app.NewWithID("com.example.aufgussscoring")

	// Apply the custom theme
	myApp.Settings().SetTheme(&CustomTheme{})

	mainWindow := myApp.NewWindow("Scoring Sheet Generator")
	mainWindow.Resize(fyne.NewSize(appWidth, 600))

	// Decode the embedded PNG
	img, _, err := image.Decode(bytes.NewReader(logoPNG))
	if err != nil {
		log.Fatalf("Failed to decode embedded image: %v", err)
	}

	// Create a Fyne image from the decoded Go image
	logo := canvas.NewImageFromImage(img)
	logo.FillMode = canvas.ImageFillContain
	logo.SetMinSize(fyne.Size{1675 / (1675 / appWidth), 163 / (1675 / appWidth)})

	spaceAbove := canvas.NewRectangle(color.Transparent)
	spaceAbove.SetMinSize(fyne.NewSize(0, 10)) // 20 pixels of space

	appLayout := container.NewVBox(
		spaceAbove,
		logo,
		spaceAbove,
		fynelayout(mainWindow, myApp),
	)

	// Set content and run the app
	mainWindow.SetContent(appLayout)

	// Add Application Menu for macOS
	mainWindow.SetMainMenu(
		fyne.NewMainMenu(
			fyne.NewMenu("",
				fyne.NewMenuItem("Preferences", func() {
					showPreferences(myApp, mainWindow) // Show Preferences dialog
				}),
			),
		),
	)

	mainWindow.ShowAndRun()
}

func getExecutableDir() (string, error) {
	execPath, err := os.Executable()
	if err != nil {
		return "", err
	}
	return filepath.Dir(execPath), nil
}

// Helper Functions

func main_2() {

	// Name of JSON file with credentials to access these Google Docs resources
	credentialsFile := "aufguss-dm-169088e9f544.json" // Replace with your JSON file
	credentialsFile = "aufguss-scoring-6ad4a3b3519f.json"

	// ID of the parent folder where the new folder for the competition will be created
	parentFolderID := "166oePFGeddWcho14g5bXULF6B6Bb5vO6" // Replace with the ID of the parent folder

	// Competition setup:
	competition := Competition{
		Name:          "Aufguss VM 2025 - Classic",
		SourceSheetID: "1VZGcUjgKaYv4yo5LEeHsUX-xLQDwfT9b6S6aO4uxVds",
		Jury: []*Juror{
			{
				Name:   "Lone",
				Weight: 100,
			},
			{
				Name:   "Bent",
				Weight: 50,
			},
			// {
			// 	Name:   "Kurt",
			// 	Weight: 50,
			// },
		},
		Contestants: []*Contestant{
			{
				Name: "Mr. White",
			},
			{
				Name: "Donald Duck",
			},
			{
				Name: "Superman",
			},
		},
	}

	// Initialize Service Account services
	ctx := context.Background()
	sheetsService, err := sheets.NewService(ctx, option.WithCredentialsFile(credentialsFile))
	if err != nil {
		log.Fatalf("Unable to create Sheets client: %v", err)
	}

	driveService, err := drive.NewService(ctx, option.WithCredentialsFile(credentialsFile))
	if err != nil {
		log.Fatalf("Unable to create Drive client: %v", err)
	}

	// Create a new empty folder for the competition
	newFolder := &drive.File{
		Name:     competition.Name,
		MimeType: "application/vnd.google-apps.folder",
		Parents:  []string{parentFolderID},
	}
	createdFolder, err := driveService.Files.Create(newFolder).Do()
	if err != nil {
		log.Fatalf("Unable to create folder: %v", err)
	}
	newFolderID := createdFolder.Id
	fmt.Printf("Created new folder '%s' with ID: %s\n", competition.Name, newFolderID)

	// Get the name of the source spreadsheet
	// sourceFile, err := driveService.Files.Get(competition.SourceSheetID).Fields("name").Do()
	// if err != nil {
	// 	log.Fatalf("Unable to retrieve source spreadsheet name: %v", err)
	// }
	// newSpreadsheetName := fmt.Sprintf("%s [copy]", sourceFile.Name)

	// Create Overview sheet:
	newFile := &drive.File{
		Name:     fmt.Sprintf("%s - Overview", competition.Name),
		MimeType: "application/vnd.google-apps.spreadsheet",
		Parents:  []string{newFolderID},
	}
	copiedFile, err := driveService.Files.Copy(competition.SourceSheetID, newFile).Do()
	if err != nil {
		log.Fatalf("Unable to copy spreadsheet: %v", err)
	}
	adminSheetId := copiedFile.Id
	fmt.Printf("Copied spreadsheet ID %s to %s for Overview\n", competition.SourceSheetID, adminSheetId)

	// Copy the "Board" sheet the number of times there is a contestant
	// Get the source spreadsheet details
	sourceSpreadsheet, err := sheetsService.Spreadsheets.Get(adminSheetId).Fields("sheets(properties(sheetId,title))").Do()
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

	pointsAndTotal, err := findPointsAndTotal(sheetsService, adminSheetId, "Board")
	if err != nil {
		log.Fatalf("Error finding Points and Total: %v", err)
	}
	//log.Println(log.Indent(pointsAndTotal))

	// Step 1: Duplicate the "Board" sheet for each contestant
	batchRequest := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{},
	}

	for i, _ := range competition.Contestants {
		// Duplicate the "Board" sheet
		newSheetName := fmt.Sprintf("AM%d", len(competition.Contestants)-i)
		duplicateRequest := &sheets.DuplicateSheetRequest{
			SourceSheetId: boardSheetID,
			NewSheetName:  newSheetName,
		}
		batchRequest.Requests = append(batchRequest.Requests, &sheets.Request{
			DuplicateSheet: duplicateRequest,
		})
	}

	// Execute the batch request to duplicate sheets
	batchResponse, err := sheetsService.Spreadsheets.BatchUpdate(adminSheetId, batchRequest).Do()
	if err != nil {
		log.Fatalf("Unable to duplicate sheets: %v", err)
	}
	fmt.Printf("Sheet 'Board' duplicated %d times.\n", len(competition.Contestants))

	// Step 2: Map duplicated sheet IDs and set contestant names
	sheetIDMap := make(map[string]int64)
	for _, reply := range batchResponse.Replies {
		if reply.DuplicateSheet != nil {
			newSheetID := reply.DuplicateSheet.Properties.SheetId
			newSheetName := reply.DuplicateSheet.Properties.Title
			sheetIDMap[newSheetName] = newSheetID
		}
	}

	// Step 3: Insert contestant names into B1 of each duplicated sheet
	batchRequest = &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{},
	}

	sheetNames := []string{}
	for i, contestant := range competition.Contestants {
		sheetName := fmt.Sprintf("AM%d", len(competition.Contestants)-i)
		sheetNames = append(sheetNames, sheetName)
		sheetID, exists := sheetIDMap[sheetName]
		if !exists {
			log.Fatalf("Could not find sheet ID for %s", sheetName)
		}

		// Prepare value update for B1
		valueUpdateRequest := &sheets.UpdateCellsRequest{
			Start: &sheets.GridCoordinate{
				SheetId:     sheetID,
				RowIndex:    0,
				ColumnIndex: 1, // Column B (0-based indexing)
			},
			Rows: []*sheets.RowData{
				{
					Values: []*sheets.CellData{
						{
							UserEnteredValue: &sheets.ExtendedValue{
								StringValue: &contestant.Name,
							},
						},
					},
				},
			},
			Fields: "userEnteredValue",
		}
		batchRequest.Requests = append(batchRequest.Requests, &sheets.Request{
			UpdateCells: valueUpdateRequest,
		})
	}

	// Execute the batch request to update contestant names
	_, err = sheetsService.Spreadsheets.BatchUpdate(adminSheetId, batchRequest).Do()
	if err != nil {
		log.Fatalf("Unable to update contestant names: %v", err)
	}
	fmt.Println("Inserted contestant names into each duplicated sheet.")

	// Create a batch update request to delete the Board sheet
	deleteSheetRequest := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{
			{
				DeleteSheet: &sheets.DeleteSheetRequest{
					SheetId: boardSheetID,
				},
			},
		},
	}

	// Execute the batch update
	_, err = sheetsService.Spreadsheets.BatchUpdate(adminSheetId, deleteSheetRequest).Do()
	if err != nil {
		log.Fatalf("Unable to delete sheet: %v", err)
	}
	fmt.Println("Sheet 'Board' deleted in the Overview sheet.")

	// Copy the Admin Google Sheet using Drive API
	jurorSheets := []string{}
	for i, juror := range competition.Jury {
		newFile := &drive.File{
			Name:     fmt.Sprintf("%s - Scoring Juror #%d (%s)", competition.Name, i+1, juror.Name),
			MimeType: "application/vnd.google-apps.spreadsheet",
			Parents:  []string{newFolderID},
		}
		copiedFile, err := driveService.Files.Copy(adminSheetId, newFile).Do()
		if err != nil {
			log.Fatalf("Unable to copy spreadsheet: %v", err)
		}
		jurorSheets = append(jurorSheets, copiedFile.Id)
		fmt.Printf("Copied Overview spreadsheet (ID %s) for Juror #%d (%s)\n", copiedFile.Id, i+1, juror.Name)
	}

	// Duplicate Juror rows in the Overview Spreadsheet:
	processJurorRows(sheetsService, adminSheetId, sheetNames, pointsAndTotal, competition.Jury, jurorSheets)
}

type RowColumnInfo struct {
	Row       int    // Row index (1-based)
	EndColumn string // Column letter with "Total:" (e.g., "B", "C"), or empty if not found
}

func findPointsAndTotal(sheetsService *sheets.Service, spreadsheetID, sheetName string) ([]RowColumnInfo, error) {
	// Range to scan: Rows 1-200, Columns A-Z
	readRange := fmt.Sprintf("%s!A1:Z200", sheetName)
	resp, err := sheetsService.Spreadsheets.Values.Get(spreadsheetID, readRange).Do()
	if err != nil {
		return nil, fmt.Errorf("unable to read data from sheet: %v", err)
	}

	results := []RowColumnInfo{}

	// Iterate over the rows in the response
	for rowIndex, row := range resp.Values {
		if len(row) > 0 && row[0] == "Points:" { // Check if column A has "Points:"
			info := RowColumnInfo{Row: rowIndex + 1} // 1-based index for rows

			// Search for "Total:" in the row (columns B-Z)
			for colIndex := 1; colIndex < len(row) && colIndex <= 25; colIndex++ {
				if row[colIndex] == "Total:" {
					info.EndColumn = string('A' + colIndex - 1) // Convert to column letter
					break
				}
			}

			results = append(results, info)
		}
	}

	return results, nil
}

func processJurorRows(
	sheetsService *sheets.Service,
	spreadsheetID string,
	sheetNames []string,
	pointsData []RowColumnInfo,
	jurors []*Juror,
	jurorSheetIDs []string,
) error {
	// Step 1: Retrieve the metadata to map sheet names to SheetIds
	sheetMetadata, err := sheetsService.Spreadsheets.Get(spreadsheetID).Fields("sheets(properties(sheetId,title))").Do()
	if err != nil {
		return fmt.Errorf("unable to retrieve sheet metadata: %v", err)
	}

	// Create a map of sheet names to SheetIds
	sheetNameToId := make(map[string]int64)
	for _, sheet := range sheetMetadata.Sheets {
		sheetNameToId[sheet.Properties.Title] = sheet.Properties.SheetId
	}

	for _, sheetName := range sheetNames {
		sheetId, exists := sheetNameToId[sheetName]
		if !exists {
			return fmt.Errorf("sheet with name %s not found in spreadsheet", sheetName)
		}

		fmt.Printf("Processing sheet: %s\n", sheetName)

		// Reverse the order of pointsData to process rows from bottom to top
		for i := len(pointsData) - 1; i >= 0; i-- {
			rowInfo := pointsData[i]

			// Read the row to be copied
			readRange := fmt.Sprintf("%s!A%d:Z%d", sheetName, rowInfo.Row, rowInfo.Row)
			resp, err := sheetsService.Spreadsheets.Values.Get(spreadsheetID, readRange).Do()
			if err != nil {
				return fmt.Errorf("failed to read row %d from sheet %s: %w", rowInfo.Row, sheetName, err)
			}

			if len(resp.Values) == 0 {
				continue // Skip if the row is empty
			}

			// Create a batch request for the current row
			batchRequest := &sheets.BatchUpdateSpreadsheetRequest{
				Requests: []*sheets.Request{},
			}

			if len(jurors) > 1 {

				// Insert a row below the current row
				insertRowRequest := &sheets.InsertRangeRequest{
					Range: &sheets.GridRange{
						SheetId:       sheetId,            // SheetId is needed; should come from the sheet name mapping
						StartRowIndex: int64(rowInfo.Row), // Zero-based index of the row
						EndRowIndex:   int64(rowInfo.Row + len(jurors) - 1),
					},
					ShiftDimension: "ROWS",
				}
				batchRequest.Requests = append(batchRequest.Requests, &sheets.Request{
					InsertRange: insertRowRequest,
				})

				// Copy the original row to the newly inserted row
				copyPasteRequest := &sheets.CopyPasteRequest{
					Source: &sheets.GridRange{
						SheetId:          sheetId,
						StartRowIndex:    int64(rowInfo.Row - 1),
						EndRowIndex:      int64(rowInfo.Row),
						StartColumnIndex: 0,
						EndColumnIndex:   26, // Copy up to column Z (adjust if needed)
					},
					Destination: &sheets.GridRange{
						SheetId:          sheetId,
						StartRowIndex:    int64(rowInfo.Row),
						EndRowIndex:      int64(rowInfo.Row + len(jurors) - 1),
						StartColumnIndex: 0,
						EndColumnIndex:   26, // Copy up to column Z (adjust if needed)
					},
					PasteType: "PASTE_NORMAL", // Preserve formatting and values
				}
				batchRequest.Requests = append(batchRequest.Requests, &sheets.Request{
					CopyPaste: copyPasteRequest,
				})
			}

			for jurorIndex := 0; jurorIndex < len(jurors); jurorIndex++ {
				jurorName := jurors[jurorIndex].Name
				jurorSheetID := jurorSheetIDs[jurorIndex] // Assume you have a map or slice with SheetIDs for each juror
				fValue := fmt.Sprintf(`=IMPORTRANGE("https://docs.google.com/spreadsheets/d/%s"; "%s!B%d:%s%d")`, jurorSheetID, sheetName, rowInfo.Row, rowInfo.EndColumn, rowInfo.Row)

				// Update the newly inserted row's column A with the juror's name
				updateCellRequest := &sheets.UpdateCellsRequest{
					Start: &sheets.GridCoordinate{
						SheetId:     sheetId,
						RowIndex:    int64(rowInfo.Row + jurorIndex - 1),
						ColumnIndex: 0,
					},
					Rows: []*sheets.RowData{
						{
							Values: []*sheets.CellData{
								{
									UserEnteredValue: &sheets.ExtendedValue{
										StringValue: &jurorName,
									},
								},
								{
									UserEnteredValue: &sheets.ExtendedValue{
										FormulaValue: &fValue,
									},
								},
							},
						},
					},
					Fields: "userEnteredValue",
				}
				batchRequest.Requests = append(batchRequest.Requests, &sheets.Request{
					UpdateCells: updateCellRequest,
				})

				f2Value := fmt.Sprintf(`=IMPORTRANGE("https://docs.google.com/spreadsheets/d/%s"; "%s!%s%d")`, jurorSheetID, sheetName, columnIndexToLetter(columnLetterToIndex(rowInfo.EndColumn)+3+1), rowInfo.Row)
				updateCellRequest = &sheets.UpdateCellsRequest{
					Start: &sheets.GridCoordinate{
						SheetId:     sheetId,
						RowIndex:    int64(rowInfo.Row + jurorIndex - 1),
						ColumnIndex: int64(columnLetterToIndex(rowInfo.EndColumn) + 3), // +3 is passing three columns to get to the feedback column
					},
					Rows: []*sheets.RowData{
						{
							Values: []*sheets.CellData{
								{
									UserEnteredValue: &sheets.ExtendedValue{
										FormulaValue: &f2Value,
									},
								},
							},
						},
					},
					Fields: "userEnteredValue",
				}
				batchRequest.Requests = append(batchRequest.Requests, &sheets.Request{
					UpdateCells: updateCellRequest,
				})
			}

			// Execute the batch request for the current row
			_, err = sheetsService.Spreadsheets.BatchUpdate(spreadsheetID, batchRequest).Do()
			if err != nil {
				return fmt.Errorf("failed to process sheet %s, row %d: %w", sheetName, rowInfo.Row, err)
			}
		}
	}
	return nil
}

func columnLetterToIndex(col string) int {
	result := 0
	for _, char := range col {
		result = result*26 + int(char-'A') + 1
	}
	return result
}

func columnIndexToLetter(index int) string {
	result := ""
	for index > 0 {
		index-- // Adjust for 1-based index
		result = string(rune('A'+(index%26))) + result
		index /= 26
	}
	return result
}
