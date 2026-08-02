// vybe-test: go/cover_image/gif_decode_config
// origin: languages/go/tests/go/test_cover_image.rs
// vybe-test-mode: compile

package main
import "image/gif"
import "bytes"
func main() { _, _ = gif.DecodeConfig(bytes.NewReader(nil)) }
