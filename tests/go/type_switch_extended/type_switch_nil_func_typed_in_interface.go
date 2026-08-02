// vybe-test: go/type_switch_extended/type_switch_nil_func_typed_in_interface
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case func(int): fmt.Println("func")
default: fmt.Println("nil-func") } }
func main() { var f func(int)
tag(f) }
