// vybe-test: go/encoding_hex_base64/hex_append_encode_slice
// origin: languages/go/tests/go/test_encoding_hex_base64.rs
// vybe-test-mode: compile

package main
import "encoding/hex"
func main() { b := []byte{}
_ = hex.AppendEncode(b, []byte("a")) }
