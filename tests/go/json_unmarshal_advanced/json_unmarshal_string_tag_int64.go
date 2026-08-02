// vybe-test: go/json_unmarshal_advanced/json_unmarshal_string_tag_int64
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs
// vybe-test-mode: compile

package main
import "encoding/json"
type T struct { V int64 `json:",string"` }
func main() { var t T
_ = json.Unmarshal([]byte(`{"V":"9223372036854775807"}`), &t) }
