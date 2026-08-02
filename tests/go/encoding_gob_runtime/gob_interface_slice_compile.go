// vybe-test: go/encoding_gob_runtime/gob_interface_slice_compile
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs
// vybe-test-mode: compile

package main
import "encoding/gob"
import "bytes"
func main() { _ = gob.NewEncoder(bytes.NewBuffer(nil)).Encode([]interface{}{1, "a"}) }
