// vybe-test: go/json_marshal/json_raw_message_holder
// origin: languages/go/tests/go/test_json_marshal.rs
// vybe-test-mode: compile

package main
import "encoding/json"
func main() { var raw json.RawMessage
_ = json.Unmarshal([]byte("[1,2]"), &raw) }
