// vybe-test: go/encoding_hex_base64/hex_decode_into_dst_buffer
// origin: languages/go/tests/go/test_encoding_hex_base64.rs
// vybe-test-mode: compile

package main
import "encoding/hex"
func main() { dst := make([]byte, 2)
_, _ = hex.Decode(dst, []byte("6162")) }
