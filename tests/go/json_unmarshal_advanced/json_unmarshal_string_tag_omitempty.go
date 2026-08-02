// vybe-test: go/json_unmarshal_advanced/json_unmarshal_string_tag_omitempty
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs
// vybe-test-mode: compile

package main
import "encoding/json"
type T struct { V int `json:",string,omitempty"` }
func main() { _, _ = json.Marshal(T{}) }
