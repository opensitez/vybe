// vybe-test: go/cover_encoding_extra/gob_decoder_decode
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/gob"
import "bytes"
func main() { var v int
d := gob.NewDecoder(bytes.NewBuffer(nil))
_ = d.Decode(&v) }
