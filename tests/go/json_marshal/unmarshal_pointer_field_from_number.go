// vybe-test: go/json_marshal/unmarshal_pointer_field_from_number
// origin: languages/go/tests/go/test_json_marshal.rs
// vybe-test-mode: compile

package main
import "encoding/json"
type Box struct { N *int }
func main() { var b Box
_ = json.Unmarshal([]byte("{\"N\":8}"), &b) }
