package main

import (
	_ "embed"

	"context"
	"fmt"

	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

func generateGoogleSheets(ctx context.Context, credentials string, parentFolderID string, competition Competition, logStatus func(message string)) error {
	// Initialize services
	services, err := initializeGoogleServices(ctx, credentials)
	if err != nil {
		return err
	}

	// Create a folder for the competition

	if err := checkContext(ctx); err != nil {
		return err
	}
	newFolderID, err := createFolder(ctx, services.Drive, parentFolderID, competition.Name, logStatus)
	if err != nil {
		return err
	}

	// Create an overview sheet

	if err := checkContext(ctx); err != nil {
		return err
	}
	adminSheetID, err := copyTemplateSheet(ctx, services, newFolderID, competition, logStatus)
	if err != nil {
		return err
	}

	// Find and process the "Board" sheet

	if err := checkContext(ctx); err != nil {
		return err
	}
	boardSheetID, pointsAndTotal, err := findBoardSheet(ctx, services.Sheets, adminSheetID, logStatus)
	if err != nil {
		return err
	}

	// Duplicate sheets and insert contestant names

	if err := checkContext(ctx); err != nil {
		return err
	}
	sheetNames, err := duplicateAndNameSheets(ctx, services.Sheets, adminSheetID, boardSheetID, competition, logStatus)
	if err != nil {
		return err
	}

	// Insert contestant names

	if err := checkContext(ctx); err != nil {
		return err
	}
	if err := insertContestantNames(ctx, services.Sheets, adminSheetID, competition, sheetNames, logStatus); err != nil {
		return err
	}

	// Delete the original "Board" sheet

	if err := checkContext(ctx); err != nil {
		return err
	}
	if err := deleteBoardSheet(ctx, services.Sheets, adminSheetID, boardSheetID, logStatus); err != nil {
		return err
	}

	// Create spreadsheets for jurors

	if err := checkContext(ctx); err != nil {
		return err
	}
	if err := createJurorSheets(ctx, services, adminSheetID, newFolderID, competition, sheetNames, pointsAndTotal, logStatus); err != nil {
		return err
	}

	return nil
}

func initializeGoogleServices(ctx context.Context, credentials string) (*GoogleServices, error) {
	sheetsService, err := sheets.NewService(ctx, option.WithCredentialsJSON([]byte(credentials)))
	if err != nil {
		return nil, fmt.Errorf("unable to create Sheets client: %v", err)
	}
	driveService, err := drive.NewService(ctx, option.WithCredentialsJSON([]byte(credentials)))
	if err != nil {
		return nil, fmt.Errorf("unable to create Drive client: %v", err)
	}
	return &GoogleServices{Sheets: sheetsService, Drive: driveService}, nil
}

func createFolder(ctx context.Context, driveService *drive.Service, parentFolderID, folderName string, logStatus func(message string)) (string, error) {
	logStatus(fmt.Sprintf("Creating new folder '%s'...\n", folderName))
	newFolder := &drive.File{
		Name:     folderName,
		MimeType: "application/vnd.google-apps.folder",
		Parents:  []string{parentFolderID},
	}
	createdFolder, err := driveService.Files.Create(newFolder).Do()
	if err != nil {
		return "", fmt.Errorf("unable to create folder: %v", err)
	}
	logStatus(fmt.Sprintf("Done. New folder '%s' has ID: %s\n", folderName, createdFolder.Id))
	return createdFolder.Id, nil
}

func copyTemplateSheet(ctx context.Context, services *GoogleServices, newFolderID string, competition Competition, logStatus func(message string)) (string, error) {
	logStatus(fmt.Sprintf("Copying Template Spreadsheet ID %s for Overview...\n", competition.SourceSheetID))
	newFile := &drive.File{
		Name:     fmt.Sprintf("%s - Overview", competition.Name),
		MimeType: "application/vnd.google-apps.spreadsheet",
		Parents:  []string{newFolderID},
	}
	copiedFile, err := services.Drive.Files.Copy(competition.SourceSheetID, newFile).Do()
	if err != nil {
		return "", fmt.Errorf("unable to copy spreadsheet: %v", err)
	}
	logStatus(fmt.Sprintf("Done. Spreadsheet ID %s was copied to ID %s\n", competition.SourceSheetID, copiedFile.Id))
	return copiedFile.Id, nil
}

func findBoardSheet(ctx context.Context, sheetsService *sheets.Service, adminSheetID string, logStatus func(message string)) (int64, []RowColumnInfo, error) {
	logStatus("Looking for sheet named 'Board' in new spreadsheet...\n")
	sourceSpreadsheet, err := sheetsService.Spreadsheets.Get(adminSheetID).Fields("sheets(properties(sheetId,title))").Do()
	if err != nil {
		return 0, nil, fmt.Errorf("unable to get spreadsheet details: %v", err)
	}
	var boardSheetID int64
	for _, sheet := range sourceSpreadsheet.Sheets {
		if sheet.Properties.Title == "Board" {
			boardSheetID = sheet.Properties.SheetId
			break
		}
	}
	if boardSheetID == 0 {
		return 0, nil, fmt.Errorf("sheet named 'Board' not found in the spreadsheet")
	}

	pointsAndTotal, err := findPointsAndTotalTokens(sheetsService, adminSheetID, "Board")
	if err != nil {
		return 0, nil, fmt.Errorf("error finding Points and Total: %v", err)
	}

	return boardSheetID, pointsAndTotal, nil
}

func duplicateAndNameSheets(ctx context.Context, sheetsService *sheets.Service, adminSheetID string, boardSheetID int64, competition Competition, logStatus func(message string)) ([]string, error) {
	logStatus(fmt.Sprintf("Duplicating sheet 'Board' %d times...\n", len(competition.Contestants)))
	batchRequest := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{},
	}
	sheetNames := make([]string, len(competition.Contestants)) // Preallocate for known length

	for i := range competition.Contestants {
		sheetName := fmt.Sprintf("AM%d", len(competition.Contestants)-i)
		sheetNames[i] = sheetName
		batchRequest.Requests = append(batchRequest.Requests, &sheets.Request{
			DuplicateSheet: &sheets.DuplicateSheetRequest{
				SourceSheetId: boardSheetID,
				NewSheetName:  sheetName,
			},
		})
	}
	_, err := sheetsService.Spreadsheets.BatchUpdate(adminSheetID, batchRequest).Do()
	if err != nil {
		return nil, fmt.Errorf("unable to duplicate sheets: %v", err)
	}
	logStatus(fmt.Sprintf("Done. Sheet 'Board' duplicated %d times.\n", len(competition.Contestants)))

	return sheetNames, nil
}

func insertContestantNames(ctx context.Context, sheetsService *sheets.Service, adminSheetID string, competition Competition, sheetNames []string, logStatus func(message string)) error {
	logStatus("Inserting contestant names into each duplicated sheet...\n")
	batchRequest := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{},
	}
	sheetIDMap := make(map[string]int64)

	// Map sheet names to IDs
	resp, err := sheetsService.Spreadsheets.Get(adminSheetID).Fields("sheets(properties(sheetId,title))").Do()
	if err != nil {
		return fmt.Errorf("unable to fetch sheets for ID mapping: %v", err)
	}
	for _, sheet := range resp.Sheets {
		sheetIDMap[sheet.Properties.Title] = sheet.Properties.SheetId
	}

	// Prepare updates for each contestant
	for i, contestant := range competition.Contestants {
		sheetName := sheetNames[len(competition.Contestants)-i-1]
		sheetID, exists := sheetIDMap[sheetName]
		if !exists {
			return fmt.Errorf("could not find sheet ID for %s", sheetName)
		}
		valueUpdateRequest := &sheets.UpdateCellsRequest{
			Start: &sheets.GridCoordinate{
				SheetId:     sheetID,
				RowIndex:    1, // Row 2  (0-based indexing)
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

	// Execute batch update
	_, err = sheetsService.Spreadsheets.BatchUpdate(adminSheetID, batchRequest).Do()
	if err != nil {
		return fmt.Errorf("unable to update contestant names: %v", err)
	}
	logStatus("Contestant names inserted successfully.\n")
	return nil
}

func deleteBoardSheet(ctx context.Context, sheetsService *sheets.Service, adminSheetID string, boardSheetID int64, logStatus func(message string)) error {
	logStatus("Deleting 'Board' sheet...\n")
	deleteSheetRequest := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{
			{
				DeleteSheet: &sheets.DeleteSheetRequest{
					SheetId: boardSheetID,
				},
			},
		},
	}
	_, err := sheetsService.Spreadsheets.BatchUpdate(adminSheetID, deleteSheetRequest).Do()
	if err != nil {
		return fmt.Errorf("unable to delete sheet: %v", err)
	}
	logStatus("Sheet 'Board' deleted in the Overview sheet.\n")
	return nil
}

func createJurorSheets(ctx context.Context, services *GoogleServices, adminSheetID, newFolderID string, competition Competition, sheetNames []string, pointsAndTotal []RowColumnInfo, logStatus func(message string)) error {
	logStatus("Creating the spreadsheet for each juror...\n")
	jurorSheets := []string{}
	for i, juror := range competition.Jury {

		if err := checkContext(ctx); err != nil {
			return err
		}

		newFile := &drive.File{
			Name:     fmt.Sprintf("%s - Scoring Juror #%d (%s)", competition.Name, i+1, juror.Name),
			MimeType: "application/vnd.google-apps.spreadsheet",
			Parents:  []string{newFolderID},
		}
		copiedFile, err := services.Drive.Files.Copy(adminSheetID, newFile).Do()
		if err != nil {
			return fmt.Errorf("unable to copy spreadsheet: %v", err)
		}
		jurorSheets = append(jurorSheets, copiedFile.Id)
		logStatus(fmt.Sprintf("Copied Overview spreadsheet for Juror #%d (%s) (Sheet ID %s)\n", i+1, juror.Name, copiedFile.Id))
	}

	if err := checkContext(ctx); err != nil {
		return err
	}
	processJurorRows(ctx, services.Sheets, adminSheetID, sheetNames, pointsAndTotal, competition.Jury, jurorSheets, logStatus)
	return nil
}

// Supporting structs
type RowColumnInfo struct {
	Row       int    // Row index (1-based)
	EndColumn string // Column letter with "Total:" (e.g., "B", "C"), or empty if not found
}

type GoogleServices struct {
	Sheets *sheets.Service
	Drive  *drive.Service
}

func processJurorRows(
	ctx context.Context,
	sheetsService *sheets.Service,
	spreadsheetID string,
	sheetNames []string,
	pointsData []RowColumnInfo,
	jurors []*Juror,
	jurorSheetIDs []string,
	logStatus func(message string),
) error {
	logStatus("Duplicating juror rows in the Overview spreadsheet...\n")

	// Retrieve sheet metadata to map sheet names to IDs
	sheetMetadata, err := sheetsService.Spreadsheets.Get(spreadsheetID).Fields("sheets(properties(sheetId,title))").Do()
	if err != nil {
		return fmt.Errorf("unable to retrieve sheet metadata: %v", err)
	}
	sheetNameToID := map[string]int64{}
	for _, sheet := range sheetMetadata.Sheets {
		sheetNameToID[sheet.Properties.Title] = sheet.Properties.SheetId
	}

	// Process each sheet
	for i, sheetName := range sheetNames {

		if err := checkContext(ctx); err != nil {
			return err
		}

		sheetID, exists := sheetNameToID[sheetName]
		if !exists {
			return fmt.Errorf("sheet with name %s not found in spreadsheet", sheetName)
		}

		logStatus(fmt.Sprintf("Processing sheet: %s (%d/%d)\n", sheetName, i+1, len(sheetNames)))

		// Process rows from bottom to top
		for j := len(pointsData) - 1; j >= 0; j-- {

			if err := checkContext(ctx); err != nil {
				return err
			}

			rowInfo := pointsData[j]

			// Read the row to be copied
			readRange := fmt.Sprintf("%s!A%d:Z%d", sheetName, rowInfo.Row, rowInfo.Row)
			resp, err := sheetsService.Spreadsheets.Values.Get(spreadsheetID, readRange).Do()
			if err != nil {
				return fmt.Errorf("failed to read row %d from sheet %s: %w", rowInfo.Row, sheetName, err)
			}
			if len(resp.Values) == 0 {
				continue // Skip empty rows
			}

			// Batch request for the row
			batchRequest := &sheets.BatchUpdateSpreadsheetRequest{Requests: []*sheets.Request{}}

			// Insert rows for jurors if needed
			if len(jurors) > 1 {
				insertRowRequest := &sheets.Request{
					InsertRange: &sheets.InsertRangeRequest{
						Range: &sheets.GridRange{
							SheetId:       sheetID,
							StartRowIndex: int64(rowInfo.Row),
							EndRowIndex:   int64(rowInfo.Row + len(jurors) - 1),
						},
						ShiftDimension: "ROWS",
					},
				}
				batchRequest.Requests = append(batchRequest.Requests, insertRowRequest)

				copyPasteRequest := &sheets.Request{
					CopyPaste: &sheets.CopyPasteRequest{
						Source: &sheets.GridRange{
							SheetId:          sheetID,
							StartRowIndex:    int64(rowInfo.Row - 1),
							EndRowIndex:      int64(rowInfo.Row),
							StartColumnIndex: 0,
							EndColumnIndex:   26,
						},
						Destination: &sheets.GridRange{
							SheetId:          sheetID,
							StartRowIndex:    int64(rowInfo.Row),
							EndRowIndex:      int64(rowInfo.Row + len(jurors) - 1),
							StartColumnIndex: 0,
							EndColumnIndex:   26,
						},
						PasteType: "PASTE_NORMAL",
					},
				}
				batchRequest.Requests = append(batchRequest.Requests, copyPasteRequest)
			}

			// Update each juror's name, points, and feedback
			for jurorIndex, juror := range jurors {
				rowOffset := int64(rowInfo.Row + jurorIndex - 1)

				// Column A: Juror's name
				batchRequest.Requests = append(batchRequest.Requests,
					createUpdateRequest(sheetID, rowOffset, 0, juror.Name, "userEnteredValue"))

				// Column B: Points formula
				pointsFormula := fmt.Sprintf(`=IMPORTRANGE("https://docs.google.com/spreadsheets/d/%s"; "%s!B%d:%s%d")`,
					jurorSheetIDs[jurorIndex], sheetName, rowInfo.Row, rowInfo.EndColumn, rowInfo.Row)
				batchRequest.Requests = append(batchRequest.Requests,
					createUpdateRequest(sheetID, rowOffset, 1, pointsFormula, "userEnteredValue"))

				// Feedback formula
				feedbackFormula := fmt.Sprintf(`=IMPORTRANGE("https://docs.google.com/spreadsheets/d/%s"; "%s!%s%d")`,
					jurorSheetIDs[jurorIndex], sheetName,
					columnIndexToLetter(columnLetterToIndex(rowInfo.EndColumn)+3+1), rowInfo.Row)
				batchRequest.Requests = append(batchRequest.Requests,
					createUpdateRequest(sheetID, rowOffset, int64(columnLetterToIndex(rowInfo.EndColumn)+3), feedbackFormula, "userEnteredValue"))

				// Juror's weight
				batchRequest.Requests = append(batchRequest.Requests,
					createUpdateRequest(sheetID, rowOffset, int64(columnLetterToIndex(rowInfo.EndColumn)+2), float64(juror.Weight)/100, "userEnteredValue"))
			}

			// Execute batch request
			if _, err := sheetsService.Spreadsheets.BatchUpdate(spreadsheetID, batchRequest).Do(); err != nil {
				return fmt.Errorf("failed to process sheet %s, row %d: %w", sheetName, rowInfo.Row, err)
			}
		}
	}

	logStatus("Finished duplicating juror rows in the Overview spreadsheet.\n")
	return nil
}

// createUpdateRequest creates a Sheets API update request for a specific cell
func createUpdateRequest(sheetID int64, rowIndex, colIndex int64, value interface{}, field string) *sheets.Request {
	var cellValue *sheets.ExtendedValue
	switch v := value.(type) {
	case string:
		if len(v) > 0 && v[0] == '=' { // Check if the string is a formula
			cellValue = &sheets.ExtendedValue{FormulaValue: &v}
		} else {
			cellValue = &sheets.ExtendedValue{StringValue: &v}
		}
	case float64:
		cellValue = &sheets.ExtendedValue{NumberValue: &v}
	}
	return &sheets.Request{
		UpdateCells: &sheets.UpdateCellsRequest{
			Start: &sheets.GridCoordinate{
				SheetId:     sheetID,
				RowIndex:    rowIndex,
				ColumnIndex: colIndex,
			},
			Rows: []*sheets.RowData{
				{
					Values: []*sheets.CellData{
						{UserEnteredValue: cellValue},
					},
				},
			},
			Fields: field,
		},
	}
}

func findPointsAndTotalTokens(sheetsService *sheets.Service, spreadsheetID, sheetName string) ([]RowColumnInfo, error) {
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

func checkContext(ctx context.Context) error {
	select {
	case <-ctx.Done():
		return ctx.Err()
	default:
		return nil
	}
}
