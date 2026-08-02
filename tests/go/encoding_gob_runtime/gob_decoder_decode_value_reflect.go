// vybe-test: go/encoding_gob_runtime/gob_decoder_decode_value_reflect
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs
// vybe-test-mode: compile

package main
import "encoding/gob"
import "bytes"
import "reflect"
func main() { d := gob.NewDecoder(bytes.NewBuffer(nil))
_ = d.DecodeValue(reflect.ValueOf(new(int)).Elem()) }
