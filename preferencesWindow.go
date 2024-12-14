package main

import (
	"context"
	"encoding/json"
	"fmt"
	"image/color"
	"io/ioutil"

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

// Preferences Dialog Function
func showPreferences(myApp fyne.App, parent fyne.Window) {

	preferencesWindow := myApp.NewWindow("Preferences")
	preferencesWindow.Resize(fyne.NewSize(600, 400))

	// Folder ID Header with Help Icon
	folderIDHeader := container.NewHBox(
		widget.NewLabelWithStyle("Google Drive Folder ID:", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		widget.NewButtonWithIcon("", theme.InfoIcon(), func() {
			showHelpDialog(preferencesWindow)
		}),
	)

	// Folder ID Entry
	folderIDEntry := widget.NewEntry()
	folderIDEntry.SetPlaceHolder("Enter Google Drive Folder ID")
	folderIDEntry.SetText(myApp.Preferences().StringWithFallback("folder_id", ""))

	spaceAbove := canvas.NewRectangle(color.Transparent)
	spaceAbove.SetMinSize(fyne.NewSize(0, 20)) // 20 pixels of space

	// Status Text for Credentials
	statusText := canvas.NewText("No credentials uploaded", color.RGBA{R: 255, G: 0, B: 0, A: 255})
	updateStatusText(statusText, myApp.Preferences().StringWithFallback("credentials", "") != "")

	// Warning Text for Errors
	warningText := canvas.NewText("", color.RGBA{R: 255, G: 0, B: 0, A: 255})
	warningText.TextSize = 12
	warningText.Alignment = fyne.TextAlignCenter
	warningContainer := container.NewVBox(warningText)
	warningContainer.Hide()

	// Upload Button
	uploadButton := widget.NewButton("Upload JSON Key", func() {
		dialog.ShowFileOpen(func(reader fyne.URIReadCloser, err error) {
			if err != nil || reader == nil {
				if err != nil {
					log.Printf("File selection error: %v", err)
				}
				return
			}
			defer reader.Close()

			// Read and validate the file
			content, readErr := ioutil.ReadAll(reader)
			if readErr != nil {
				dialog.ShowError(fmt.Errorf("failed to read file: %w", readErr), preferencesWindow)
				return
			}

			if ok, valErr := IsValidServiceAccountJSON(string(content)); ok {
				myApp.Preferences().SetString("credentials", string(content))
				updateStatusText(statusText, string(content) != "")
				warningContainer.Hide()
			} else {
				warningText.Text = fmt.Sprintf("Invalid credentials: %v", valErr)
				warningContainer.Show()
				warningText.Refresh()
				log.Printf("Error validating credentials: %v", valErr)
			}
		}, preferencesWindow)
	})

	// Save Button
	saveButton := widget.NewButton("Save", func() {
		myApp.Preferences().SetString("folder_id", folderIDEntry.Text)
		dialog.ShowInformation("Success", "Preferences saved successfully.", preferencesWindow)
		preferencesWindow.Close()
	})

	// Preferences Layout
	preferencesWindow.SetContent(container.NewVBox(
		folderIDHeader,
		folderIDEntry,
		spaceAbove,
		widget.NewLabelWithStyle("Upload Google Service Account JSON key:", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		uploadButton,
		warningContainer,
		statusText,
		container.NewHBox(
			layout.NewSpacer(),
			saveButton,
		),
	))
	preferencesWindow.Show()
}

// Updates the status text and color
func updateStatusText(statusText *canvas.Text, credentialsNotEmpty bool) {
	if credentialsNotEmpty {
		statusText.Text = "Credentials Uploaded."
		statusText.Color = color.RGBA{R: 0, G: 128, B: 0, A: 255} // Green
	} else {
		statusText.Text = "No credentials uploaded."
		statusText.Color = color.RGBA{R: 255, G: 0, B: 0, A: 255} // Red
	}
	statusText.Refresh()
}

// Shows the help dialog for Folder ID
func showHelpDialog(parent fyne.Window) {
	richText := widget.NewRichText(
		&widget.TextSegment{Text: "The Google Drive Folder ID is the unique identifier\nfor the folder where your spreadsheets are stored.\nExample:\n"},
		&widget.TextSegment{Text: "https://drive.google.com/drive/u/1/folders/"},
		&widget.TextSegment{
			Text:  "567oeIUYeddWcho14g5bXULFJD7Db5vO6",
			Style: widget.RichTextStyleStrong, // Bold
		},
		&widget.TextSegment{Text: "?ths=true\n\nwhere the part in bold is the folder ID."},
	)
	dialog.ShowCustom("What is the Google Drive Folder ID?", "Close", container.NewVBox(richText), parent)
}

// Validates the service account JSON
func IsValidServiceAccountJSON(input string) (bool, error) {
	if len(input) > 10000 {
		return false, fmt.Errorf("file too long")
	}
	if err := testGoogleLogin(input); err != nil {
		return false, err
	}
	return true, nil
}

// Tests Google Service Account login
func testGoogleLogin(credentialsJSON string) error {
	var credentials map[string]interface{}
	if err := json.Unmarshal([]byte(credentialsJSON), &credentials); err != nil {
		return fmt.Errorf("invalid JSON structure: %v", err)
	}
	if credentials["type"] != "service_account" {
		return fmt.Errorf("invalid credentials type: %v", credentials["type"])
	}

	ctx := context.Background()
	driveService, err := drive.NewService(ctx, option.WithCredentialsJSON([]byte(credentialsJSON)))
	if err != nil {
		return fmt.Errorf("unable to create Drive client: %v", err)
	}

	_, err = driveService.Files.List().PageSize(1).Do()
	if err != nil {
		return fmt.Errorf("authentication failed: %v", err)
	}
	return nil
}
