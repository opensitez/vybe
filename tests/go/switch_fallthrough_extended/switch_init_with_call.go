// vybe-test: go/switch_fallthrough_extended/switch_init_with_call
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func val() int { return 3 }
func main() { switch v := val(); v { case 3: fmt.Println(v)
default: fmt.Println(0) } }
