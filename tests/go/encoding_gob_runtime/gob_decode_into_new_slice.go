// vybe-test: go/encoding_gob_runtime/gob_decode_into_new_slice
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs
// vybe-test-mode: compile

package main
import "encoding/gob"
import "bytes"
func main() { var s []string
_ = gob.NewDecoder(bytes.NewBuffer(nil)).Decode(&s) }
