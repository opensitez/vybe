// vybe-test: go/encoding_hex_base64/hex_invalid_byte_error_value
// origin: languages/go/tests/go/test_encoding_hex_base64.rs
// vybe-test-mode: compile

package main
import "encoding/hex"
func main() { _ = hex.InvalidByte }
