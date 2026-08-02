// vybe-test: go/stdlib_mime_runtime/color_rgba_model
// origin: languages/go/tests/go/test_stdlib_mime_runtime.rs
// vybe-test-mode: compile

package main
import "image/color"
func main() { _ = color.RGBAModel.Convert(color.RGBA{R: 255}) }
