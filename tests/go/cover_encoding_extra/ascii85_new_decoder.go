// vybe-test: go/cover_encoding_extra/ascii85_new_decoder
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/ascii85"
import "bytes"
func main() { _ = ascii85.NewDecoder(bytes.NewBufferString("<~00~>")) }
