// vybe-test: go/type_switch_extended/type_switch_interface_holding_slice
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case []interface{}: fmt.Println(len(v.([]interface{})))
default: fmt.Println(0) } }
func main() { tag([]interface{}{1, "a"}) }
