package main

import (
	"bytes"
	_ "embed"
	"image"
	"image/color"
	"os"
	"path/filepath"
	"strings"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/container"
	log "github.com/s00500/env_logger"
)

// Embed the logo file
//
//go:embed embedded/logo.png
var logoPNG []byte

const (
	competitionsDir = "competitions" // Directory to store competition files
	appWidth        = 800            // Default application width
)

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

var dataDir string

func main() {
	initializeDataDir()

	myApp := app.NewWithID("com.example.aufgussscoring")
	myApp.Settings().SetTheme(&CustomTheme{})

	mainWindow := myApp.NewWindow("Scoring Sheet Generator")
	mainWindow.Resize(fyne.NewSize(appWidth, 600))

	logo := loadLogo()

	appLayout := container.NewVBox(
		spaceAbove(10),
		logo,
		spaceAbove(10),
		fynelayout(mainWindow, myApp),
	)

	mainWindow.SetContent(appLayout)
	mainWindow.SetMainMenu(createAppMenu(myApp, mainWindow))
	mainWindow.ShowAndRun()
}

func initializeDataDir() {
	execPath, err := getExecutableDir()
	if err != nil {
		log.Fatalf("Failed to get executable directory: %v", err)
	}

	exePathParts := strings.Split(execPath, "Aufguss Scoring Generator.app")
	basePath := ""
	if len(exePathParts) == 2 {
		basePath = exePathParts[0]
	}

	dataDir = filepath.Join(basePath, competitionsDir)
	if _, err := os.Stat(dataDir); os.IsNotExist(err) {
		if err := os.MkdirAll(dataDir, 0755); err != nil {
			log.Fatalf("Failed to create competitions directory: %v", err)
		}
	}
	log.Printf("Data directory: %s", dataDir)
}

func getExecutableDir() (string, error) {
	execPath, err := os.Executable()
	if err != nil {
		return "", err
	}
	return filepath.Dir(execPath), nil
}

func loadLogo() *canvas.Image {
	img, _, err := image.Decode(bytes.NewReader(logoPNG))
	if err != nil {
		log.Fatalf("Failed to decode embedded image: %v", err)
	}

	logo := canvas.NewImageFromImage(img)
	logo.FillMode = canvas.ImageFillContain
	ratio := 1675 / appWidth
	logo.SetMinSize(fyne.NewSize(float32(1675/ratio), float32(163/ratio)))

	return logo
}

func createAppMenu(myApp fyne.App, mainWindow fyne.Window) *fyne.MainMenu {
	return fyne.NewMainMenu(
		fyne.NewMenu("",
			fyne.NewMenuItem("Preferences", func() {
				showPreferences(myApp, mainWindow)
			}),
		),
	)
}

func spaceAbove(height int) *canvas.Rectangle {
	spaceAbove := canvas.NewRectangle(color.Transparent)
	spaceAbove.SetMinSize(fyne.NewSize(0, 10))
	return spaceAbove
}
