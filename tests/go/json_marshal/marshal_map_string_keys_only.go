// vybe-test: go/json_marshal/marshal_map_string_keys_only
// origin: languages/go/tests/go/test_json_marshal.rs
// vybe-test-mode: compile

package main
import "encoding/json"
func main() { _, _ = json.Marshal(map[string]string{"k": "v"}) }
