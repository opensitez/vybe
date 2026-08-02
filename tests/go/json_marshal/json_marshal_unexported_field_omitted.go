// vybe-test: go/json_marshal/json_marshal_unexported_field_omitted
// origin: languages/go/tests/go/test_json_marshal.rs
// vybe-test-mode: compile

package main
import "encoding/json"
type T struct { pub int
priv int }
func main() { _, _ = json.Marshal(T{pub: 1, priv: 2}) }
