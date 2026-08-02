// vybe-test: go/encoding_hex_base64/base64_std_decoded_len_estimate
// origin: languages/go/tests/go/test_encoding_hex_base64.rs
// vybe-test-mode: compile

package main
import "encoding/base64"
func main() { _ = base64.StdEncoding.DecodedLen(4) }
