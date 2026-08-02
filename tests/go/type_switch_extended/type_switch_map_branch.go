// vybe-test: go/type_switch_extended/type_switch_map_branch
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case map[string]int: fmt.Println("map") default: fmt.Println("other") } }
func main() { tag(map[string]int{"a": 1}) }
