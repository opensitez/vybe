// vybe-test: go/cover_image/png_decode_config
// origin: languages/go/tests/go/test_cover_image.rs
// vybe-test-mode: compile

package main
import "image/png"
import "bytes"
func main() { _, _ = png.DecodeConfig(bytes.NewReader(nil)) }
