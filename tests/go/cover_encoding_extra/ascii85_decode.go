// vybe-test: go/cover_encoding_extra/ascii85_decode
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/ascii85"
func main() { dst := make([]byte, 4)
_, _, _ = ascii85.Decode(dst, []byte("<~00~>"), true) }
