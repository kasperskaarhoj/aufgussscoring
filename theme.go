package main

import (
	"image/color"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/theme"
)

// CustomTheme defines a custom theme for the application
type CustomTheme struct{}

// Color overrides theme colors
func (c *CustomTheme) Color(name fyne.ThemeColorName, variant fyne.ThemeVariant) color.Color {
	switch name {
	case theme.ColorNameBackground:
		return color.White // Background color
	case theme.ColorNameForeground:
		return color.Black // Foreground (text) color
	case theme.ColorNameDisabled:
		return color.RGBA{R: 128, G: 128, B: 128, A: 255} // Disabled text color
	case theme.ColorNameButton:
		return color.RGBA{R: 220, G: 220, B: 220, A: 255} // Button background
	default:
		return theme.DefaultTheme().Color(name, variant) // Fallback to default theme
	}
}

// Font overrides theme fonts
func (c *CustomTheme) Font(style fyne.TextStyle) fyne.Resource {
	return theme.DefaultTheme().Font(style)
}

// Icon overrides theme icons
func (c *CustomTheme) Icon(name fyne.ThemeIconName) fyne.Resource {
	return theme.DefaultTheme().Icon(name)
}

// Size overrides theme sizes
func (c *CustomTheme) Size(name fyne.ThemeSizeName) float32 {
	return theme.DefaultTheme().Size(name)
}
