package main

import (
	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/widget"
)

// PaddedContainer wraps a fyne.CanvasObject and adds a border-like padding around it
type PaddedContainer struct {
	widget.BaseWidget
	content fyne.CanvasObject
	width   float32
}

// NewPaddedContainer creates a new PaddedContainer
func NewPaddedContainer(content fyne.CanvasObject, borderWidth float32) *PaddedContainer {
	p := &PaddedContainer{
		content: content,
		width:   borderWidth,
	}
	p.ExtendBaseWidget(p)
	return p
}

// CreateRenderer creates the custom renderer for the PaddedContainer
func (p *PaddedContainer) CreateRenderer() fyne.WidgetRenderer {
	return &paddedContainerRenderer{
		container: p,
		objects:   []fyne.CanvasObject{p.content},
	}
}

type paddedContainerRenderer struct {
	container *PaddedContainer
	objects   []fyne.CanvasObject
}

func (r *paddedContainerRenderer) Layout(size fyne.Size) {
	// Calculate the inner size after accounting for the padding (border width)
	innerSize := size.Subtract(fyne.NewSize(r.container.width*2, r.container.width*2))
	r.container.content.Resize(innerSize)
	r.container.content.Move(fyne.NewPos(r.container.width, r.container.width))
}

func (r *paddedContainerRenderer) MinSize() fyne.Size {
	// Minimum size includes the content size plus padding
	return r.container.content.MinSize().Add(fyne.NewSize(r.container.width*2, r.container.width*2))
}

func (r *paddedContainerRenderer) Refresh() {
	// No visual updates required for padding; refresh the content
	r.container.content.Refresh()
}

func (r *paddedContainerRenderer) Objects() []fyne.CanvasObject {
	return r.objects
}

func (r *paddedContainerRenderer) Destroy() {}
