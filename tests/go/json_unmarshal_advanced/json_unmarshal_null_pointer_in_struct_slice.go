// vybe-test: go/json_unmarshal_advanced/json_unmarshal_null_pointer_in_struct_slice
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs
// vybe-test-mode: compile

package main
import "encoding/json"
type Item struct { N *int }
func main() { var s []Item
_ = json.Unmarshal([]byte(`[{"N":null},{"N":1}]`), &s) }
