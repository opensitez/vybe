// vybe-test: go/switch_fallthrough_extended/switch_on_interface_holding_int
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { var v interface{} = 8
switch v.(type) { case int: fmt.Println("int")
default: fmt.Println("other") } }
