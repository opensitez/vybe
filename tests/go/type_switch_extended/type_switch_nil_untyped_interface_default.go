// vybe-test: go/type_switch_extended/type_switch_nil_untyped_interface_default
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case int: fmt.Println("int") default: fmt.Println("nil-default") } }
func main() { tag(nil) }
