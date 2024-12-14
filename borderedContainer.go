package main

import (
	"image/color"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/widget"
)

// BorderedContainer wraps a fyne.CanvasObject and draws a border around it
type BorderedContainer struct {
	widget.BaseWidget
	content fyne.CanvasObject
	color   color.Color
	width   float32
}

// NewBorderedContainer creates a new BorderedContainer
func NewBorderedContainer(content fyne.CanvasObject, borderColor color.Color, borderWidth float32) *BorderedContainer {
	b := &BorderedContainer{
		content: content,
		color:   borderColor,
		width:   borderWidth,
	}
	b.ExtendBaseWidget(b)
	return b
}

// CreateRenderer creates the custom renderer for the BorderedContainer
func (b *BorderedContainer) CreateRenderer() fyne.WidgetRenderer {
	border := canvas.NewRectangle(b.color)
	return &borderedContainerRenderer{
		container: b,
		border:    border,
		objects:   []fyne.CanvasObject{border, b.content},
	}
}

type borderedContainerRenderer struct {
	container *BorderedContainer
	border    *canvas.Rectangle
	objects   []fyne.CanvasObject
}

func (r *borderedContainerRenderer) Layout(size fyne.Size) {
	r.border.Resize(size)
	r.border.StrokeWidth = r.container.width
	innerSize := size.Subtract(fyne.NewSize(r.container.width*2, r.container.width*2))
	r.container.content.Resize(innerSize)
	r.container.content.Move(fyne.NewPos(r.container.width, r.container.width))
}

func (r *borderedContainerRenderer) MinSize() fyne.Size {
	return r.container.content.MinSize().Add(fyne.NewSize(r.container.width*2, r.container.width*2))
}

func (r *borderedContainerRenderer) Refresh() {
	r.border.FillColor = r.container.color
	r.border.StrokeWidth = r.container.width
	r.border.Refresh()
}

func (r *borderedContainerRenderer) Objects() []fyne.CanvasObject {
	return r.objects
}

func (r *borderedContainerRenderer) Destroy() {}
