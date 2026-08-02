// vybe-test: go/json_unmarshal_advanced/json_raw_message_marshal_nil
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs
// vybe-test-mode: compile

package main
import "encoding/json"
func main() { var raw json.RawMessage
_, _ = json.Marshal(raw) }
