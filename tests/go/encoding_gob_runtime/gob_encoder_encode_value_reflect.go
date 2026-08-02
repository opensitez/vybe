// vybe-test: go/encoding_gob_runtime/gob_encoder_encode_value_reflect
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs
// vybe-test-mode: compile

package main
import "encoding/gob"
import "bytes"
import "reflect"
func main() { e := gob.NewEncoder(bytes.NewBuffer(nil))
_ = e.EncodeValue(reflect.ValueOf(1)) }
