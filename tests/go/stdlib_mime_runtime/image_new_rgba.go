// vybe-test: go/stdlib_mime_runtime/image_new_rgba
// origin: languages/go/tests/go/test_stdlib_mime_runtime.rs
// vybe-test-mode: compile

package main
import "image"
func main() { _ = image.NewRGBA(image.Rect(0, 0, 2, 2)) }
