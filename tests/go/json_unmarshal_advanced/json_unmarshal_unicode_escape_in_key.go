// vybe-test: go/json_unmarshal_advanced/json_unmarshal_unicode_escape_in_key
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs
// vybe-test-mode: compile

package main
import "encoding/json"
func main() { var m map[string]int
_ = json.Unmarshal([]byte(`{"\u0061":1}`), &m) }
