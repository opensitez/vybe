// vybe-test: go/cover_encoding_extra/ascii85_max_encoded_len
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/ascii85"
func main() { _ = ascii85.MaxEncodedLen(8) }
