// vybe-test: go/json_marshal/unmarshal_empty_json_array
// origin: languages/go/tests/go/test_json_marshal.rs
// vybe-test-mode: compile

package main
import "encoding/json"
func main() { var s []string
_ = json.Unmarshal([]byte("[]"), &s) }
