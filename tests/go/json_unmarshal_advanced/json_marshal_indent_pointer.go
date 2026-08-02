// vybe-test: go/json_unmarshal_advanced/json_marshal_indent_pointer
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs
// vybe-test-mode: compile

package main
import "encoding/json"
func main() { n := 5
_, _ = json.MarshalIndent(&n, "", "  ") }
