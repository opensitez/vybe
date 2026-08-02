// vybe-test: go/cover_encoding_extra/gob_encoder_encode
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/gob"
import "bytes"
func main() { e := gob.NewEncoder(bytes.NewBuffer(nil))
_ = e.Encode(42) }
