// vybe-test: go/switch_fallthrough_extended/switch_on_interface_nil
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { var v interface{}
switch v.(type) { case nil: fmt.Println("nil")
default: fmt.Println("other") } }
