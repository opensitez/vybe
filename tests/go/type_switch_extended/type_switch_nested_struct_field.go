// vybe-test: go/type_switch_extended/type_switch_nested_struct_field
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
type inner struct { x int }
type outer struct { in inner }
func tag(v interface{}) { switch t := v.(type) { case outer: fmt.Println(t.in.x)
default: fmt.Println(0) } }
func main() { tag(outer{in: inner{x: 8}}) }
