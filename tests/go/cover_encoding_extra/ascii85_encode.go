// vybe-test: go/cover_encoding_extra/ascii85_encode
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/ascii85"
func main() { dst := make([]byte, ascii85.MaxEncodedLen(4))
_ = ascii85.Encode(dst, []byte("go")) }
