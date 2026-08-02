// vybe-test: go/type_switch_extended/type_switch_pointer_to_interface_box
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case *interface{}: fmt.Println("ptr-if")
default: fmt.Println("other") } }
func main() { var x interface{} = 1
tag(&x) }
