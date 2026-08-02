// vybe-test: go/type_switch_extended/type_switch_map_with_struct_value
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
type kv struct { k string }
func tag(v interface{}) { switch v.(type) { case map[string]kv: fmt.Println(len(v.(map[string]kv)))
default: fmt.Println(0) } }
func main() { tag(map[string]kv{"a": {k: "x"}}) }
