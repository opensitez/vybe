// vybe-test: go/json_unmarshal_advanced/json_decoder_use_number_float_string
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs
// vybe-test-mode: compile

package main
import "encoding/json"
import "bytes"
func main() { dec := json.NewDecoder(bytes.NewReader([]byte(`"3.14"`)))
dec.UseNumber()
var v interface{}
_ = dec.Decode(&v) }
