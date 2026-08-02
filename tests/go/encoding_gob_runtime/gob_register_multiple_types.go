// vybe-test: go/encoding_gob_runtime/gob_register_multiple_types
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs
// vybe-test-mode: compile

package main
import "encoding/gob"
func main() { gob.Register(int(0))
gob.Register(string(""))
gob.Register([]int{}) }
