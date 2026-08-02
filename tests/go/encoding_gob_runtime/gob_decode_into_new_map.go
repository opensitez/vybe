// vybe-test: go/encoding_gob_runtime/gob_decode_into_new_map
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs
// vybe-test-mode: compile

package main
import "encoding/gob"
import "bytes"
func main() { var m map[string]int
_ = gob.NewDecoder(bytes.NewBuffer(nil)).Decode(&m) }
