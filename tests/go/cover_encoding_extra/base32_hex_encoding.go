// vybe-test: go/cover_encoding_extra/base32_hex_encoding
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/base32"
func main() { _ = base32.HexEncoding.EncodeToString([]byte("go")) }
