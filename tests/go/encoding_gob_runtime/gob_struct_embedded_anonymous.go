// vybe-test: go/encoding_gob_runtime/gob_struct_embedded_anonymous
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs
// vybe-test-mode: compile

package main
import "encoding/gob"
import "bytes"
type Base struct { ID int }
type Derived struct { Base
Name string }
func main() { _ = gob.NewEncoder(bytes.NewBuffer(nil)).Encode(Derived{Base: Base{ID: 1}, Name: "d"}) }
