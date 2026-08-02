// vybe-test: go/cover_encoding_extra/ascii85_alphabet_len
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/ascii85"
func main() { _ = len(ascii85.Encode(make([]byte, 4), []byte("go"))) }
