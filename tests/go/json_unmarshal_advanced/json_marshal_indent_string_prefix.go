// vybe-test: go/json_unmarshal_advanced/json_marshal_indent_string_prefix
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs
// vybe-test-mode: compile

package main
import "encoding/json"
func main() { _, _ = json.MarshalIndent(struct{ A int }{1}, "tab", "  ") }
