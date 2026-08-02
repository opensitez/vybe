// vybe-test: go/json_marshal/json_string_tag_int_field
// origin: languages/go/tests/go/test_json_marshal.rs
// vybe-test-mode: compile

package main
import "encoding/json"
type N struct { Val int `json:",string"` }
func main() { _, _ = json.Marshal(N{Val: 7}) }
