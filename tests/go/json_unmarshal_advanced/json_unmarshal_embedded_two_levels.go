// vybe-test: go/json_unmarshal_advanced/json_unmarshal_embedded_two_levels
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs
// vybe-test-mode: compile

package main
import "encoding/json"
type A struct { N int }
type B struct { A }
type C struct { B }
func main() { var c C
_ = json.Unmarshal([]byte(`{"N":4}`), &c) }
