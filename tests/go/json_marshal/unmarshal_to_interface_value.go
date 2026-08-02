// vybe-test: go/json_marshal/unmarshal_to_interface_value
// origin: languages/go/tests/go/test_json_marshal.rs
// vybe-test-mode: compile

package main
import "encoding/json"
func main() { var v interface{}
_ = json.Unmarshal([]byte("{\"n\":1}"), &v) }
