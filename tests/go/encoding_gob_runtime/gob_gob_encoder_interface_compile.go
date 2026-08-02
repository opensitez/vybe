// vybe-test: go/encoding_gob_runtime/gob_gob_encoder_interface_compile
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs
// vybe-test-mode: compile

package main
import "encoding/gob"
type T struct { N int }
func (t T) GobEncode() ([]byte, error) { return nil, nil }
func main() { var _ gob.GobEncoder = T{} }
