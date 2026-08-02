// vybe-test: go/cover_image/jpeg_decode_config
// origin: languages/go/tests/go/test_cover_image.rs
// vybe-test-mode: compile

package main
import "image/jpeg"
import "bytes"
func main() { _, _ = jpeg.DecodeConfig(bytes.NewReader(nil)) }
