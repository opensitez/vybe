// vybe-test: go/cover_image/color_rgb_to_rgba
// origin: languages/go/tests/go/test_cover_image.rs
// vybe-test-mode: compile

package main
import "image/color"
func main() { _ = color.RGBA{R: 1, G: 2, B: 3, A: 255} }
