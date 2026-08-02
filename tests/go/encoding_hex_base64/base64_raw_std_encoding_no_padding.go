// vybe-test: go/encoding_hex_base64/base64_raw_std_encoding_no_padding
// origin: languages/go/tests/go/test_encoding_hex_base64.rs
// vybe-test-mode: compile

package main
import "encoding/base64"
func main() { _ = base64.RawStdEncoding.EncodeToString([]byte("f")) }
