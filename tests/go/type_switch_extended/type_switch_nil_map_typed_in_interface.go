// vybe-test: go/type_switch_extended/type_switch_nil_map_typed_in_interface
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case map[int]int: fmt.Println("map")
default: fmt.Println("nil-map") } }
func main() { var m map[int]int
tag(m) }
