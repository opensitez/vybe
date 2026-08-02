// vybe-test: go/json_marshal/unmarshal_ignores_unknown_fields
// origin: languages/go/tests/go/test_json_marshal.rs
// vybe-test-mode: compile

package main
import "encoding/json"
type S struct { X int }
func main() { var s S
_ = json.Unmarshal([]byte("{\"X\":1,\"extra\":\"y\"}"), &s) }
