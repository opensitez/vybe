// vybe-test: go/json_marshal/marshal_fixed_size_array
// origin: languages/go/tests/go/test_json_marshal.rs
// vybe-test-mode: compile

package main
import "encoding/json"
func main() { _, _ = json.Marshal([2]int{4, 5}) }
