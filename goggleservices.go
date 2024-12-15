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

	// Helper function to check for context cancellation
	checkContext := func() error {
		select {
		case <-ctx.Done():
			return ctx.Err()
		default:
			return nil
		}
	}

	// Initialize Service Account services
	sheetsService, err := sheets.NewService(ctx, option.WithCredentialsJSON([]byte(credentials)))
	if err != nil {
		return fmt.Errorf("unable to create Sheets client: %v", err)
	}
	driveService, err := drive.NewService(ctx, option.WithCredentialsJSON([]byte(credentials)))
	if err != nil {
		return fmt.Errorf("unable to create Drive client: %v", err)
	}

	// Create a new folder for the competition
	if err := checkContext(); err != nil {
		return err
	}
	logStatus(fmt.Sprintf("Creating new folder '%s'...\n", competition.Name))
	newFolder := &drive.File{
		Name:     competition.Name,
		MimeType: "application/vnd.google-apps.folder",
		Parents:  []string{parentFolderID},
	}
	createdFolder, err := driveService.Files.Create(newFolder).Do()
	if err != nil {
		return fmt.Errorf("unable to create folder: %v", err)
	}
	newFolderID := createdFolder.Id
	logStatus(fmt.Sprintf("Done. New folder '%s' has ID: %s\n", competition.Name, newFolderID))

	// Create Overview sheet
	if err := checkContext(); err != nil {
		return err
	}
	logStatus(fmt.Sprintf("Copying Template Spreadsheet ID %s for Overview...\n", competition.SourceSheetID))
	newFile := &drive.File{
		Name:     fmt.Sprintf("%s - Overview", competition.Name),
		MimeType: "application/vnd.google-apps.spreadsheet",
		Parents:  []string{newFolderID},
	}
	copiedFile, err := driveService.Files.Copy(competition.SourceSheetID, newFile).Do()
	if err != nil {
		return fmt.Errorf("unable to copy spreadsheet: %v", err)
	}
	adminSheetId := copiedFile.Id
	logStatus(fmt.Sprintf("Done. Spreadsheet ID %s was copied to ID %s\n", competition.SourceSheetID, adminSheetId))

	// Find and process the "Board" sheet
	if err := checkContext(); err != nil {
		return err
	}
	logStatus("Looking for sheet named 'Board' in new spreadsheet...\n")
	sourceSpreadsheet, err := sheetsService.Spreadsheets.Get(adminSheetId).Fields("sheets(properties(sheetId,title))").Do()
	if err != nil {
		return fmt.Errorf("unable to get spreadsheet details: %v", err)
	}
	var boardSheetID int64
	for _, sheet := range sourceSpreadsheet.Sheets {
		if sheet.Properties.Title == "Board" {
			boardSheetID = sheet.Properties.SheetId
			break
		}
	}
	if boardSheetID == 0 {
		return fmt.Errorf("sheet named 'Board' not found in the spreadsheet")
	}

	pointsAndTotal, err := findPointsAndTotal(sheetsService, adminSheetId, "Board")
	if err != nil {
		return fmt.Errorf("error finding Points and Total: %v", err)
	}

	// Duplicate the "Board" sheet for each contestant
	if err := checkContext(); err != nil {
		return err
	}
	logStatus(fmt.Sprintf("Duplicating sheet 'Board' %d times...\n", len(competition.Contestants)))
	batchRequest := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{},
	}
	for i := range competition.Contestants {
		newSheetName := fmt.Sprintf("AM%d", len(competition.Contestants)-i)
		duplicateRequest := &sheets.DuplicateSheetRequest{
			SourceSheetId: boardSheetID,
			NewSheetName:  newSheetName,
		}
		batchRequest.Requests = append(batchRequest.Requests, &sheets.Request{
			DuplicateSheet: duplicateRequest,
		})
	}
	batchResponse, err := sheetsService.Spreadsheets.BatchUpdate(adminSheetId, batchRequest).Do()
	if err != nil {
		return fmt.Errorf("unable to duplicate sheets: %v", err)
	}
	logStatus(fmt.Sprintf("Done. Sheet 'Board' duplicated %d times.\n", len(competition.Contestants)))

	// Map duplicated sheet IDs
	sheetIDMap := make(map[string]int64)
	for _, reply := range batchResponse.Replies {
		if reply.DuplicateSheet != nil {
			sheetIDMap[reply.DuplicateSheet.Properties.Title] = reply.DuplicateSheet.Properties.SheetId
		}
	}

	// Insert contestant names
	if err := checkContext(); err != nil {
		return err
	}
	logStatus("Inserting contestant names into each duplicated sheet...\n")
	batchRequest = &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{},
	}
	sheetNames := []string{}
	for i, contestant := range competition.Contestants {
		sheetName := fmt.Sprintf("AM%d", len(competition.Contestants)-i)
		sheetNames = append(sheetNames, sheetName)
		sheetID, exists := sheetIDMap[sheetName]
		if !exists {
			return fmt.Errorf("could not find sheet ID for %s", sheetName)
		}
		valueUpdateRequest := &sheets.UpdateCellsRequest{
			Start: &sheets.GridCoordinate{
				SheetId:     sheetID,
				RowIndex:    0,
				ColumnIndex: 1,
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
	_, err = sheetsService.Spreadsheets.BatchUpdate(adminSheetId, batchRequest).Do()
	if err != nil {
		return fmt.Errorf("unable to update contestant names: %v", err)
	}
	logStatus("Done.\n")

	// Delete "Board" sheet
	if err := checkContext(); err != nil {
		return err
	}
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
	_, err = sheetsService.Spreadsheets.BatchUpdate(adminSheetId, deleteSheetRequest).Do()
	if err != nil {
		return fmt.Errorf("unable to delete sheet: %v", err)
	}
	logStatus("Sheet 'Board' deleted in the Overview sheet.\n")

	// Create spreadsheets for jurors
	if err := checkContext(); err != nil {
		return err
	}
	logStatus("Creating the spreadsheet for each juror...\n")
	jurorSheets := []string{}
	for i, juror := range competition.Jury {
		if err := checkContext(); err != nil {
			return err
		}
		newFile := &drive.File{
			Name:     fmt.Sprintf("%s - Scoring Juror #%d (%s)", competition.Name, i+1, juror.Name),
			MimeType: "application/vnd.google-apps.spreadsheet",
			Parents:  []string{newFolderID},
		}
		copiedFile, err := driveService.Files.Copy(adminSheetId, newFile).Do()
		if err != nil {
			return fmt.Errorf("unable to copy spreadsheet: %v", err)
		}
		jurorSheets = append(jurorSheets, copiedFile.Id)
		logStatus(fmt.Sprintf("Copied Overview spreadsheet for Juror #%d (%s) (Sheet ID %s)\n", i+1, juror.Name, copiedFile.Id))
	}

	// Duplicate Juror rows in the Overview Spreadsheet
	processJurorRows(sheetsService, adminSheetId, sheetNames, pointsAndTotal, competition.Jury, jurorSheets, logStatus)

	return nil
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
	logStatus func(message string),
) error {

	logStatus(fmt.Sprintf("Duplicating juror rows in the Overview spreadsheet...\n"))

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

	for i, sheetName := range sheetNames {
		sheetId, exists := sheetNameToId[sheetName]
		if !exists {
			return fmt.Errorf("sheet with name %s not found in spreadsheet", sheetName)
		}

		logStatus(fmt.Sprintf("Processing sheet: %s (%d/%d)\n", sheetName, i+1, len(sheetName)))

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
