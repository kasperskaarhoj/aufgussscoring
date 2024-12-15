package main

import (
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"fmt"
	"image/color"
	"strconv"

	log "github.com/s00500/env_logger"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"
)

func fynelayout(myWindow fyne.Window, myApp fyne.App) *fyne.Container {
	// Declare the fileMap and its mutex in the main scope
	fileMap := make(map[string]string)
	var fileMapMutex sync.RWMutex

	// Competition Name
	nameEntry := widget.NewEntry()
	nameEntry.SetPlaceHolder("Enter Competition Name")

	// Jurors Slice
	jurors := []*Juror{}

	// Create Jury Table
	juryTableComposition, juryTable := createJuryTable(&jurors)

	// Contestants Slice
	contestants := []*Contestant{}

	// Create Contestants Table
	contestantTableComposition, contestantTable := createContestantsTable(&contestants)

	// Create Template Sheet Selector
	templateSheetSelectorContainer, templateSheetSelect := createTemplateSheetSelector(myApp, &fileMap, &fileMapMutex)

	// Create a read-only log field with black text
	logField := NewCustomLogField(color.Black)

	// Wrap the log field in a scrollable container
	scrollableLog := container.NewVScroll(logField)
	scrollableLog.SetMinSize(fyne.NewSize(400, 150)) // Set a minimum size for the log window

	var right, left *fyne.Container

	// Dropdown for loading competitions
	files, err := loadCompetitionFiles()
	if err != nil {
		log.Fatalf("Failed to load files: %v", err)
	}
	files = append(files, "[Create New]")

	fileSelect := widget.NewSelect(files, func(selected string) {
		if selected != "" {
			loadCompetition(selected, nameEntry, templateSheetSelect, &jurors, &contestants, fileMap, &fileMapMutex, juryTable, contestantTable)
			right.Show()
			left.Show()
		}
	})

	// Save button
	saveButton := widget.NewButton("Save", func() {
		// Validate the competition name
		if strings.TrimSpace(nameEntry.Text) == "" {
			dialog.ShowError(fmt.Errorf("Competition name cannot be empty."), myWindow)
			return
		}

		fileMapMutex.RLock()
		defer fileMapMutex.RUnlock()
		var sheetId string
		for name, id := range fileMap {
			if name == templateSheetSelect.Selected {
				sheetId = id
				break
			}
		}

		competition := buildCompetition(strings.TrimSpace(nameEntry.Text), sheetId, jurors, contestants)
		if err := saveCompetition(competition); err != nil {
			dialog.ShowError(err, myWindow)
		} else {
			dialog.ShowInformation("Success", "Competition saved successfully", myWindow)
			files, _ := loadCompetitionFiles() // Reload files
			fileSelect.Options = files
			fileSelect.SetSelected(fmt.Sprintf("%s.json", competition.Name))
		}
	})

	var cancelFunc context.CancelFunc
	var cancelButton *widget.Button
	var generateButton *widget.Button
	generateButton = widget.NewButton("Generate!", func() {

		// Validate the competition name
		if strings.TrimSpace(nameEntry.Text) == "" {
			dialog.ShowError(fmt.Errorf("Competition name cannot be empty."), myWindow)
			return
		}

		fileMapMutex.RLock()
		defer fileMapMutex.RUnlock()
		var sheetId string
		for name, id := range fileMap {
			if name == templateSheetSelect.Selected {
				sheetId = id
				break
			}
		}

		competition := buildCompetition(strings.TrimSpace(nameEntry.Text), sheetId, jurors, contestants)

		// Validate competition fields
		if err := validateCompetition(competition); err != nil {
			dialog.ShowError(err, myWindow)
			return
		}

		if err := saveCompetition(competition); err != nil {
			dialog.ShowError(err, myWindow)
		} else {
			cancelButton.Enable()
			generateButton.Disable()

			// Reload files
			files, _ := loadCompetitionFiles()
			fileSelect.Options = files
			fileSelect.SetSelected(fmt.Sprintf("%s.json", competition.Name))

			logField.SetText("Generating...\n")
			logFunction := func(message string) {
				logField.SetText(fmt.Sprintf("%s%s", logField.Text, message))
			}

			// Create a context with cancellation
			ctx, cancel := context.WithCancel(context.Background())
			cancelFunc = cancel

			// Run the sheet generation asynchronously
			go func() {
				err := generateGoogleSheets(ctx, myApp.Preferences().String("credentials"), myApp.Preferences().String("folder_id"), competition, logFunction)
				if err != nil {
					logFunction(fmt.Sprintf("Error: %v\n", err))
				} else {
					logFunction("Generation completed successfully.\n")
				}
				cancelButton.Disable()
				generateButton.Enable()
			}()
		}
	})
	cancelButton = widget.NewButton("Cancel", func() {
		cancelFunc()
	})
	cancelButton.Disable()

	// Delete button
	deleteButton := widget.NewButton("Delete", func() {
		// Ensure a file is selected in fileSelect
		if fileSelect.Selected == "" || fileSelect.Selected == "[Create New]" {
			dialog.ShowInformation("No Selection", "Please select a valid file to delete.", myWindow)
			return
		}

		// Find the index of the selected file
		selectedIndex := -1
		for i, file := range fileSelect.Options {
			if file == fileSelect.Selected {
				selectedIndex = i
				break
			}
		}

		if selectedIndex == -1 {
			dialog.ShowInformation("File Not Found", "The selected file could not be found in the list.", myWindow)
			return
		}

		// Show confirmation dialog
		dialog.ShowConfirm("Confirm Delete",
			fmt.Sprintf("Are you sure you want to delete '%s'?", fileSelect.Selected),
			func(confirm bool) {
				if confirm {
					// Construct the file path using filepath.Join
					filePath := filepath.Join(dataDir, fileSelect.Selected)

					// Attempt to delete the file
					err := os.Remove(filePath)
					if err != nil {
						dialog.ShowError(fmt.Errorf("Failed to delete file: %w", err), myWindow)
						return
					}

					// Refresh the fileSelect options
					files, _ := loadCompetitionFiles()
					files = append(files, "[Create New]")
					fileSelect.Options = files
					fileSelect.Refresh()

					// Determine the next file to select
					nextIndex := selectedIndex
					if nextIndex >= len(files)-1 {
						// If deleted file was the last in the list, move to the previous file
						nextIndex = len(files) - 2
					}

					if len(files) == 1 { // Only [Create New] remains
						fileSelect.SetSelected("[Create New]")
						right.Hide()
						left.Hide()
					} else if nextIndex >= 0 && nextIndex < len(files) {
						// Automatically select the next file if available
						fileSelect.SetSelected(files[nextIndex])
						loadCompetition(
							files[nextIndex],
							nameEntry,
							templateSheetSelect,
							&jurors,
							&contestants,
							fileMap,
							&fileMapMutex,
							juryTable,
							contestantTable,
						)
						right.Show()
						left.Show()
					}
				}
			},
			myWindow,
		)
	})

	spaceAbove := canvas.NewRectangle(color.Transparent)
	spaceAbove.SetMinSize(fyne.NewSize(0, 10))

	// Layout in two columns:
	left = container.NewVBox(
		widget.NewLabelWithStyle("Competition Name:", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		nameEntry,

		widget.NewLabelWithStyle("Template Sheet:", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		templateSheetSelectorContainer,

		spaceAbove,

		container.NewHBox(
			saveButton,
			deleteButton,
			layout.NewSpacer(),
			generateButton,
			cancelButton,
		),

		container.NewVBox(
			container.NewHBox(
				widget.NewLabelWithStyle("Log:", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
				layout.NewSpacer(),
				widget.NewButtonWithIcon("", theme.ContentClearIcon(), func() {
					logField.SetText("") // Clear the log
				}),
			),
			scrollableLog,
		),
	)
	left.Hide()

	right = container.NewVBox(
		juryTableComposition,
		spaceAbove,
		contestantTableComposition,
	)
	right.Hide()

	twoColumnLayout := container.NewGridWithColumns(2,
		NewPaddedContainer(
			container.NewVBox(
				widget.NewLabelWithStyle("Select Competition:", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
				fileSelect,
				left,
			), 10),
		NewPaddedContainer(right, 10),
	)

	return twoColumnLayout
}

var jurorsMutex sync.RWMutex

func createJuryTable(jurors *[]*Juror) (*fyne.Container, *widget.Table) {

	// Create the jury table
	juryTable := widget.NewTable(
		func() (int, int) {
			jurorsMutex.RLock()
			defer jurorsMutex.RUnlock()
			return len(*jurors), 2 // Rows and columns: jurors count and 2 columns (Name, Weight)
		},
		func() fyne.CanvasObject {
			// Create a single Entry for each cell
			return widget.NewEntry()
		},
		func(id widget.TableCellID, obj fyne.CanvasObject) {
			entry := obj.(*widget.Entry)

			jurorsMutex.RLock()
			defer jurorsMutex.RUnlock()

			if id.Col == 0 { // Name column
				entry.OnChanged = nil // Must release the old OnChanged function before setting the text below, otherwise it would lock up!
				entry.SetText((*jurors)[id.Row].Name)
				entry.OnChanged = func(newText string) {
					jurorsMutex.Lock()
					defer jurorsMutex.Unlock()
					(*jurors)[id.Row].Name = strings.TrimSpace(newText)
				}
			} else if id.Col == 1 { // Weight column
				entry.OnChanged = nil // Must release the old OnChanged function before setting the text below, otherwise it would lock up!
				entry.SetText(fmt.Sprintf("%d", (*jurors)[id.Row].Weight))
				entry.OnChanged = func(newText string) {
					jurorsMutex.Lock()
					defer jurorsMutex.Unlock()
					weight, err := strconv.Atoi(strings.TrimSpace(newText))
					if err == nil {
						(*jurors)[id.Row].Weight = weight
					}
				}
			}
		},
	)

	// Set column widths for proper sizing
	juryTable.SetColumnWidth(0, 300) // Name column width
	juryTable.SetColumnWidth(1, 80)  // Weight column width

	// Add a bounding rectangle to enforce table size
	boundingBox := canvas.NewRectangle(nil)
	boundingBox.SetMinSize(fyne.NewSize(400, 150)) // Set desired size for the table

	// Layer the table on top of the bounding box
	tableWithBounds := container.NewStack(boundingBox, juryTable)

	// Header with buttons for adding and removing jurors
	header := container.NewHBox(
		widget.NewLabelWithStyle("Jury:", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		layout.NewSpacer(),
		widget.NewButton("Add", func() {
			// Append a new juror with default values
			*jurors = append(*jurors, &Juror{Name: "", Weight: 0})
			juryTable.Refresh()
		}),
		widget.NewButton("Remove", func() {
			// Remove the last juror if there are any
			if len(*jurors) > 0 {
				*jurors = (*jurors)[:len(*jurors)-1]
				juryTable.Refresh()
			}
		}),
	)

	// Return the complete layout and the table widget
	return container.NewVBox(
		header,
		tableWithBounds,
	), juryTable
}

var contestantsMutex sync.RWMutex

func createContestantsTable(contestants *[]*Contestant) (*fyne.Container, *widget.Table) {

	// Create the contestants table
	contestantsTable := widget.NewTable(
		func() (int, int) {
			contestantsMutex.RLock()
			defer contestantsMutex.RUnlock()
			return len(*contestants), 1 // Rows: contestants count, Columns: 1 (Name)
		},
		func() fyne.CanvasObject {
			// Create a single Entry for each cell
			return widget.NewEntry()
		},
		func(id widget.TableCellID, obj fyne.CanvasObject) {
			entry := obj.(*widget.Entry)

			contestantsMutex.RLock()
			defer contestantsMutex.RUnlock()

			// Populate the name field
			entry.OnChanged = nil
			entry.SetText((*contestants)[id.Row].Name)
			entry.OnChanged = func(newText string) {
				contestantsMutex.Lock()
				defer contestantsMutex.Unlock()
				(*contestants)[id.Row].Name = strings.TrimSpace(newText)
			}
		},
	)

	// Set column width for the name column
	contestantsTable.SetColumnWidth(0, 380) // Name column width

	// Add a bounding rectangle to enforce table size
	boundingBox := canvas.NewRectangle(nil)
	boundingBox.SetMinSize(fyne.NewSize(400, 250)) // Set desired size for the table

	// Layer the table on top of the bounding box
	tableWithBounds := container.NewStack(boundingBox, contestantsTable)

	// Header with buttons for adding and removing contestants
	header := container.NewHBox(
		widget.NewLabelWithStyle("Contestants:", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		layout.NewSpacer(),
		widget.NewButton("Add", func() {
			// Append a new contestant with default values
			*contestants = append(*contestants, &Contestant{Name: ""})
			contestantsTable.Refresh()
		}),
		widget.NewButton("Remove", func() {
			// Remove the last contestant if there are any
			if len(*contestants) > 0 {
				*contestants = (*contestants)[:len(*contestants)-1]
				contestantsTable.Refresh()
			}
		}),
	)

	// Return the complete layout and the table widget
	return container.NewVBox(
		header,
		tableWithBounds,
	), contestantsTable
}

func createTemplateSheetSelector(
	myApp fyne.App,
	fileMap *map[string]string,
	fileMapMutex *sync.RWMutex,
) (fyne.CanvasObject, *widget.Select) {
	// Styled Warning Label
	warningText := canvas.NewText("", color.RGBA{R: 255, G: 0, B: 0, A: 255}) // Red text
	warningText.TextSize = 12
	warningText.Alignment = fyne.TextAlignCenter

	warningContainer := container.NewVBox(warningText)
	warningContainer.Hide() // Initially hidden

	// Template Sheet Selector
	templateSheetSelect := widget.NewSelect([]string{}, nil)
	templateSheetSelect.PlaceHolder = "Select a Template"

	// Fetch Google Sheets from Drive
	fetchFiles := func() {
		// Fetch the folder ID and credentials
		folderID := myApp.Preferences().String("folder_id")
		credentials := myApp.Preferences().String("credentials")

		// Initialize the Drive API client
		ctx := context.Background()
		service, err := drive.NewService(ctx, option.WithCredentialsJSON([]byte(credentials)))
		if err != nil {
			warningText.Text = "Failed to connect to Google Drive. Check credentials."
			warningContainer.Show()
			warningText.Refresh()
			return
		}

		// Query for Google Sheets in the specified folder
		query := fmt.Sprintf("'%s' in parents and mimeType='application/vnd.google-apps.spreadsheet'", folderID)
		filesList, err := service.Files.List().Q(query).Fields("files(id, name)").Do()
		if err != nil {
			warningText.Text = "Failed to fetch files from Google Drive folder."
			warningContainer.Show()
			warningText.Refresh()
			return
		}

		// Update the fileMap safely with the mutex
		fileMapMutex.Lock()
		defer fileMapMutex.Unlock()

		warningContainer.Hide()
		*fileMap = make(map[string]string) // Clear the old map
		options := []string{}
		for _, file := range filesList.Files {
			options = append(options, file.Name)
			(*fileMap)[file.Name] = file.Id
		}

		if len(filesList.Files) == 0 {
			warningText.Text = "No files were found in the folder or you don't have access"
			warningContainer.Show()
			warningText.Refresh()
		}

		templateSheetSelect.Options = options
		templateSheetSelect.Refresh()
	}

	// Refresh every 5 seconds
	go func() {
		for {
			fetchFiles()
			time.Sleep(5 * time.Second)
		}
	}()

	// Handle selection change
	templateSheetSelect.OnChanged = func(selected string) {
		fileMapMutex.RLock()
		defer fileMapMutex.RUnlock()

		// if id, exists := (*fileMap)[selected]; exists {
		// 	fmt.Printf("Selected File Name: %s, ID: %s\n", selected, id)
		// }
	}

	// Return the layout
	return container.NewVBox(
		templateSheetSelect,
		warningContainer,
	), templateSheetSelect
}

// Validation function
func validateCompetition(comp Competition) error {
	// Check for empty competition name
	if strings.TrimSpace(comp.Name) == "" {
		return fmt.Errorf("Competition name cannot be empty.")
	}

	// Check for at least one juror
	if len(comp.Jury) == 0 {
		return fmt.Errorf("There must be at least one juror.")
	}

	// Check for empty juror fields and valid weights
	for i, juror := range comp.Jury {
		if strings.TrimSpace(juror.Name) == "" {
			return fmt.Errorf("Juror #%d has an empty name.", i+1)
		}
		if juror.Weight < 0 || juror.Weight > 100 {
			return fmt.Errorf("Juror #%d has an invalid weight (%d). Must be between 0 and 100.", i+1, juror.Weight)
		}
	}

	// Check for at least one contestant
	if len(comp.Contestants) == 0 {
		return fmt.Errorf("There must be at least one contestant.")
	}

	// Check for empty contestant names
	for i, contestant := range comp.Contestants {
		if strings.TrimSpace(contestant.Name) == "" {
			return fmt.Errorf("Contestant #%d has an empty name.", i+1)
		}
	}

	// Check if a template sheet is defined
	if strings.TrimSpace(comp.SourceSheetID) == "" {
		return fmt.Errorf("A template sheet must be defined.")
	}

	// If no errors, return nil
	return nil
}

func buildCompetition(name, sourceSheetID string, jurors []*Juror, contestants []*Contestant) Competition {
	return Competition{
		Name:          name,
		SourceSheetID: sourceSheetID,
		Jury:          jurors,
		Contestants:   contestants,
	}
}

func saveCompetition(comp Competition) error {
	if _, err := os.Stat(dataDir); os.IsNotExist(err) {
		os.Mkdir(dataDir, 0755)
	}
	data, err := json.MarshalIndent(comp, "", "  ")
	if err != nil {
		return fmt.Errorf("failed to serialize competition: %w", err)
	}
	filePath := fmt.Sprintf("%s/%s.json", dataDir, comp.Name)
	return os.WriteFile(filePath, data, 0644)
}

func loadCompetition(
	filename string,
	nameEntry *widget.Entry,
	templateSheetSelector *widget.Select,
	jurors *[]*Juror,
	contestants *[]*Contestant,
	fileMap map[string]string,
	fileMapMutex *sync.RWMutex,
	jurorsTable *widget.Table,
	contestantsTable *widget.Table,
) {
	var comp Competition

	if filename == "[Create New]" {
		comp = Competition{}
	} else {
		filePath := fmt.Sprintf("%s/%s", dataDir, filename)
		data, err := os.ReadFile(filePath)
		if err != nil {
			log.Printf("Failed to load competition: %v", err)
			return
		}

		if err := json.Unmarshal(data, &comp); err != nil {
			log.Printf("Failed to parse competition: %v", err)
			return
		}
	}

	// Populate competition details
	nameEntry.SetText(comp.Name)

	// Look up SourceSheetID in fileMap to set the display name
	fileMapMutex.RLock()
	defer fileMapMutex.RUnlock()
	var displayName string
	for name, id := range fileMap {
		if id == comp.SourceSheetID {
			displayName = name
			break
		}
	}

	if displayName != "" {
		templateSheetSelector.SetSelected(displayName)
	} else {
		log.Printf("SourceSheetID %s not found in fileMap", comp.SourceSheetID)
		templateSheetSelector.ClearSelected() // Clear the selection if no match
	}

	// Clear and populate contestants
	contestantsMutex.Lock()
	*contestants = comp.Contestants // Update the contestants slice directly
	contestantsMutex.Unlock()
	contestantsTable.Refresh() // Refresh the contestants table to reflect the new data

	// Clear and populate jurors
	jurorsMutex.Lock()
	*jurors = comp.Jury // Update the slice directly
	jurorsMutex.Unlock()
	jurorsTable.Refresh() // Refresh the table to reflect the new data
}

func splitLines(text string) []string {
	lines := strings.Split(text, "\n")
	for i := range lines {
		lines[i] = strings.TrimSpace(lines[i])
	}
	return lines
}

func formatContestants(contestants []*Contestant) string {
	var result string
	for _, contestant := range contestants {
		result += fmt.Sprintf("%s\n", contestant.Name)
	}
	return result
}

func loadCompetitionFiles() ([]string, error) {
	if _, err := os.Stat(dataDir); os.IsNotExist(err) {
		return nil, nil
	}
	files, err := os.ReadDir(dataDir)
	if err != nil {
		return nil, err
	}
	var names []string
	for _, file := range files {
		if !file.IsDir() {
			names = append(names, file.Name())
		}
	}
	return names, nil
}
