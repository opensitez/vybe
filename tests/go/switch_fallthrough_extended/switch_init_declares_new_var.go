// vybe-test: go/switch_fallthrough_extended/switch_init_declares_new_var
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { x := 10
switch y := 5; y { case 5: fmt.Println(y)
default: fmt.Println(0) }
fmt.Println(x) }
