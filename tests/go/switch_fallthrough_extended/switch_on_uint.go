// vybe-test: go/switch_fallthrough_extended/switch_on_uint
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch uint(7) { case 7: fmt.Println("u")
default: fmt.Println("other") } }
