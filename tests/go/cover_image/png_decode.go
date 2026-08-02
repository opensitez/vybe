// vybe-test: go/cover_image/png_decode
// origin: languages/go/tests/go/test_cover_image.rs
// vybe-test-mode: compile

package main
import "image/png"
import "bytes"
func main() { _, _ = png.Decode(bytes.NewReader(nil)) }
