// vybe-test: go/json_unmarshal_advanced/json_decoder_disallow_unknown_fields
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs
// vybe-test-mode: compile

package main
import "encoding/json"
import "bytes"
type S struct { X int }
func main() { dec := json.NewDecoder(bytes.NewReader([]byte(`{"X":1,"extra":2}`)))
dec.DisallowUnknownFields()
_ = dec.Decode(&S{}) }
