// vybe-test: go/json_unmarshal_advanced/json_decoder_disallow_unknown_array_elem
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs
// vybe-test-mode: compile

package main
import "encoding/json"
import "bytes"
func main() { dec := json.NewDecoder(bytes.NewReader([]byte(`[1,{"x":1,"y":2}]`)))
dec.DisallowUnknownFields()
var v []interface{}
_ = dec.Decode(&v) }
