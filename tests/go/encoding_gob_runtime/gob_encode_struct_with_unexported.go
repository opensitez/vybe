// vybe-test: go/encoding_gob_runtime/gob_encode_struct_with_unexported
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs
// vybe-test-mode: compile

package main
import "encoding/gob"
import "bytes"
type S struct { pub int
priv string }
func main() { _ = gob.NewEncoder(bytes.NewBuffer(nil)).Encode(S{pub: 1, priv: "x"}) }
