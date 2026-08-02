// vybe-test: go/encoding_hex_base64/base64_std_decode_into_dst_buffer
// origin: languages/go/tests/go/test_encoding_hex_base64.rs
// vybe-test-mode: compile

package main
import "encoding/base64"
func main() { dst := make([]byte, 4)
_, _ = base64.StdEncoding.Decode(dst, []byte("Zm9v")) }
