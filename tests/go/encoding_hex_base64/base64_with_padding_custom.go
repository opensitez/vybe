// vybe-test: go/encoding_hex_base64/base64_with_padding_custom
// origin: languages/go/tests/go/test_encoding_hex_base64.rs
// vybe-test-mode: compile

package main
import "encoding/base64"
func main() { enc := base64.StdEncoding.WithPadding('*')
_ = enc.EncodeToString([]byte("f")) }
