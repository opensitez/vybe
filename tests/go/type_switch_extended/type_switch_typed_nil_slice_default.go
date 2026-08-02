// vybe-test: go/type_switch_extended/type_switch_typed_nil_slice_default
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case []int: fmt.Println("slice") default: fmt.Println("default") } }
func main() { var s []int
tag(s) }
