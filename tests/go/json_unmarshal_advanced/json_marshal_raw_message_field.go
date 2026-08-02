// vybe-test: go/json_unmarshal_advanced/json_marshal_raw_message_field
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs
// vybe-test-mode: compile

package main
import "encoding/json"
type W struct { Raw json.RawMessage `json:"raw"` }
func main() { _, _ = json.Marshal(W{Raw: json.RawMessage(`{"k":1}`)}) }
