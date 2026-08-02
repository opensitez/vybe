// vybe-test: go/json_marshal/json_marshal_indent_prefix
// origin: languages/go/tests/go/test_json_marshal.rs
// vybe-test-mode: compile

package main
import "encoding/json"
func main() { _, _ = json.MarshalIndent(map[string]int{"a": 1}, "", "  ") }
