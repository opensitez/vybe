// vybe-test: go/json_unmarshal_advanced/json_decoder_disallow_unknown_nested
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs
// vybe-test-mode: compile

package main
import "encoding/json"
import "bytes"
type Inner struct { N int }
type Outer struct { Inner Inner }
func main() { dec := json.NewDecoder(bytes.NewReader([]byte(`{"Inner":{"N":1,"bad":2}}`)))
dec.DisallowUnknownFields()
var o Outer
_ = dec.Decode(&o) }
