// vybe-test: go/cover_image/jpeg_encode
// origin: languages/go/tests/go/test_cover_image.rs
// vybe-test-mode: compile

package main
import "image/jpeg"
import "image"
import "bytes"
func main() { _ = jpeg.Encode(bytes.NewBuffer(nil), image.NewRGBA(image.Rect(0, 0, 1, 1)), nil) }
