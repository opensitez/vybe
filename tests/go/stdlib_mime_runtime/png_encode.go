// vybe-test: go/stdlib_mime_runtime/png_encode
// origin: languages/go/tests/go/test_stdlib_mime_runtime.rs
// vybe-test-mode: compile

package main
import "image/png"
import "image"
import "bytes"
func main() { _ = png.Encode(bytes.NewBuffer(nil), image.NewRGBA(image.Rect(0, 0, 1, 1))) }
