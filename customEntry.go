package main

import (
	"image/color"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/widget"
)

type CustomLogField struct {
	widget.Entry
	textColor color.Color
}

func NewCustomLogField(textColor color.Color) *CustomLogField {
	logField := &CustomLogField{textColor: textColor}
	logField.ExtendBaseWidget(logField)
	logField.Disable() // Make it read-only
	return logField
}

// CreateRenderer overrides the default renderer to customize text color
func (l *CustomLogField) CreateRenderer() fyne.WidgetRenderer {
	renderer := l.Entry.CreateRenderer()
	for _, obj := range renderer.Objects() {
		if text, ok := obj.(*canvas.Text); ok {
			text.Color = l.textColor // Set text color to black
		}
	}
	return renderer
}
