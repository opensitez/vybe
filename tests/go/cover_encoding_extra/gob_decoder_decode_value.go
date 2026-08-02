// vybe-test: go/cover_encoding_extra/gob_decoder_decode_value
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/gob"
import "bytes"
import "reflect"
func main() { d := gob.NewDecoder(bytes.NewBuffer(nil))
_ = d.DecodeValue(reflect.ValueOf(new(int)).Elem()) }
