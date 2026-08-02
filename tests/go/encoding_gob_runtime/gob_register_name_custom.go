// vybe-test: go/encoding_gob_runtime/gob_register_name_custom
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs
// vybe-test-mode: compile

package main
import "encoding/gob"
type Widget struct { X int }
func main() { gob.RegisterName("my.Widget", Widget{}) }
