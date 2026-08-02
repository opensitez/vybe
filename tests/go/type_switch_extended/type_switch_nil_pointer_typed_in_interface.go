// vybe-test: go/type_switch_extended/type_switch_nil_pointer_typed_in_interface
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
type node struct { v int }
func tag(v interface{}) { switch v.(type) { case *node: fmt.Println("ptr")
default: fmt.Println("nil-ptr") } }
func main() { var p *node
tag(p) }
